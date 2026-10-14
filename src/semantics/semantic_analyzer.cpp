#include "semantics/semantic_analyzer.hpp"
#include "semantics/semantic_analyzer_internal.h"
#include <algorithm>
#include <cctype>
#include <tuple>
#include <initializer_list>

namespace vb6c3 {

// ============================================================
// 构造/析构
// ============================================================

SemanticAnalyzer::SemanticAnalyzer(Diagnostics& diag, bool verbose)
    : diag_(diag), symTab_(diag), verbose_(verbose) {
    for (int i = 0; i < 26; i++) defTypeMap_[i] = Vb6Type::Variant;
}

// ============================================================
// 两遍扫描入口
// ============================================================

bool SemanticAnalyzer::analyze(Module& module) {
    // 泛型模板类 (tB, G4): 模板本体不进符号表 (类符号/成员/字段全部不注册),
    // 由泛型器把整模块特化克隆注入 modules_ 走完整管线. 空分析器保持
    // modules_/analyzers_ 1:1 对齐.
    if (module.isClassModule && !module.classTypeParams.empty()) return true;

    currentModule_ = &module;

    // 注册内置对象、函数和常量
    registerBuiltins();

    // 如果是类模块，注册类符号 (让其他模块可以通过类名引用)
    if (module.isClassModule) {
        auto classSym = std::make_unique<Symbol>(
            SymbolKind::Class, module.moduleName, Vb6Type::Object,
            SourceLocation{module.filename, 1, 1}, AccessLevel::Public
        );
        classSym->instancing = module.instancing;
        // 收集类成员名称
        for (auto& decl : module.declarations) {
            switch (decl->kind) {
                // Fix 092m: 类数据字段类型表 — 供 cgen 的 With 块字段写按目标字段类型
                // 选择 COM 解包函数 (见 Symbol::memberFieldTypes 注释). 仅记录
                // As <简单类型名> 的字段; 其它 (数组/UDT/未声明) 由 cgen 回退默认.
                case ASTNodeKind::VariableDecl: {
                    auto& v092m = static_cast<VariableDecl&>(*decl);
                    // Fix 092p: 登记字段声明原名 — C 结构体成员名按声明生成, 访问点需
                    // 把源码里的大小写变体 (.socket) 规范化回该名.
                    classSym->memberFieldNames[Symbol::toLower(v092m.name)] = v092m.name;
                    // tB B08a: 记录访问级别 (Private 字段在真 VB6 里等价于 Dim, 也照样记)
                    classSym->memberAccessLevels[Symbol::toLower(v092m.name)] = v092m.access;
                    if (v092m.asType && v092m.asType->kind == ASTNodeKind::SimpleTypeRef) {
                        classSym->memberFieldTypes[Symbol::toLower(v092m.name)] =
                            static_cast<SimpleTypeRef*>(v092m.asType.get())->name;
                    }
                    // Fix 099: 显式 Public 字段 → COM 暴露清单 (见 Symbol::publicFieldNames).
                    // 只认 AccessLevel::Public: 类模块的 `Dim x` 被解析为
                    // AccessLevel::Default (=Public), 但真 VB6 中 Dim 等价 Private,
                    // 不能当成对外可见的字段. 数组/WithEvents 不暴露.
                    // Fix 187: As New 字段同样暴露 —— 实测真 VB6 (MSVBVM60 编译的
                    // VBMAN.dll) 把 `Public HttpClient As New cHttpClient` 暴露为
                    // PropertyGet + PropertyPutRef (同一 dispid) 对; 此前排除 As New
                    // 使晚绑定客户端 vb6_ComGetProp("HttpClient") 取不到属性.
                    if (v092m.access == AccessLevel::Public
                        && !v092m.isWithEvents
                        && !v092m.isDynamicArray
                        && v092m.dimensions.empty()) {
                        classSym->publicFieldNames.push_back(v092m.name);
                    }
                    break;
                }
                case ASTNodeKind::SubDecl: {
                    auto& s = static_cast<SubDecl&>(*decl);
                    classSym->memberNames.push_back(s.name);
                    // Fix 016: Sub 写入 memberProcKinds (覆盖任意前值)
                    classSym->memberProcKinds[Symbol::toLower(s.name)] = ProcKind::Sub;
                    classSym->memberAccessLevels[Symbol::toLower(s.name)] = s.access;  // tB B08a
                    // Fix 033: memberParams 改在 Pass1 填充 (class init 阶段类型解析不完整, 导致 Variant 回归)
                    break;
                }
                case ASTNodeKind::FunctionDecl: {
                    auto& f = static_cast<FunctionDecl&>(*decl);
                    classSym->memberNames.push_back(f.name);
                    // Fix 015: 记录 Function 返回类型名(仅类型为命名类型 SimpleTypeRef 时)
                    // 用于跨模块 storageKey 冲突场景下的 method chaining 返回类型推断
                    if (f.returnType && f.returnType->kind == ASTNodeKind::SimpleTypeRef) {
                        classSym->memberReturnTypes[Symbol::toLower(f.name)] =
                            static_cast<SimpleTypeRef*>(f.returnType.get())->name;
                    }
                    // Fix 016: Function 写入 memberProcKinds (覆盖任意前值 — 同类内
                    // 不允许 Function 与同名 Property 共存, 故此处覆盖无冲突风险)
                    classSym->memberProcKinds[Symbol::toLower(f.name)] = ProcKind::Function;
                    classSym->memberAccessLevels[Symbol::toLower(f.name)] = f.access;  // tB B08a
                    // Fix 033: memberParams 改在 Pass1 填充 (class init 阶段类型解析不完整)
                    break;
                }
                case ASTNodeKind::PropertyDecl: {
                    auto& p = static_cast<PropertyDecl&>(*decl);
                    // 同名Property Get/Let/Set只加一次memberNames
                    {
                        std::string lower = Symbol::toLower(p.name);
                        bool found = false;
                        for (auto& mn : classSym->memberNames) {
                            if (Symbol::toLower(mn) == lower) { found = true; break; }
                        }
                        if (!found) classSym->memberNames.push_back(p.name);
                    }
                    // Fix 015: 仅 Property Get 有返回值; Let/Set 无返回类型不记录.
                    // 用前判 propKind==PropertyGet 避免被后续 Let/Set 覆盖 Get 的返回类型.
                    if (p.propKind == ProcKind::PropertyGet
                        && p.returnType
                        && p.returnType->kind == ASTNodeKind::SimpleTypeRef) {
                        classSym->memberReturnTypes[Symbol::toLower(p.name)] =
                            static_cast<SimpleTypeRef*>(p.returnType.get())->name;
                    }
                    // Fix 016: 按 ProcKind 写入 memberProcKinds, 同名共存时按
                    // 读上下文优先级 Get > Function > Sub > Let > Set 选择, 即:
                    // - Get 总是覆盖 (最高优先级)
                    // - Let 仅在键不存在或现有是 Let/Set 时写入 (不覆盖 Get/Function/Sub)
                    // - Set 仅在键不存在或现有是 Set/Let 时写入 (不覆盖 Get/Function/Sub/Let)
                    // Fix 033: memberParams 改在 Pass1 填充 (此处类型解析不完整, 且 memberProcKinds
                    //         在本阶段结束后已确定最终胜出者, Pass1 可直接据其判定 wins)
                    {
                        std::string lower = Symbol::toLower(p.name);
                        auto it = classSym->memberProcKinds.find(lower);
                        bool wins = false;
                        if (p.propKind == ProcKind::PropertyGet) {
                            wins = true;
                        } else if (p.propKind == ProcKind::PropertyLet) {
                            if (it == classSym->memberProcKinds.end()
                                || it->second == ProcKind::PropertyLet
                                || it->second == ProcKind::PropertySet) {
                                wins = true;
                            }
                        } else if (p.propKind == ProcKind::PropertySet) {
                            if (it == classSym->memberProcKinds.end()
                                || it->second == ProcKind::PropertySet) {
                                wins = true;
                            }
                        }
                        if (wins) {
                            classSym->memberAccessLevels[lower] = p.access;  // tB B08a
                            if (p.propKind == ProcKind::PropertyGet) {
                                classSym->memberProcKinds[lower] = ProcKind::PropertyGet;
                            } else if (p.propKind == ProcKind::PropertyLet) {
                                classSym->memberProcKinds[lower] = ProcKind::PropertyLet;
                            } else if (p.propKind == ProcKind::PropertySet) {
                                classSym->memberProcKinds[lower] = ProcKind::PropertySet;
                            }
                        }
                    }
                    break;
                }
                case ASTNodeKind::EventDecl: {
                    auto& e = static_cast<EventDecl&>(*decl);
                    classSym->memberNames.push_back(e.name);
                    classSym->eventNames.push_back(e.name);  // P6.5: 收集事件名
                    classSym->memberAccessLevels[Symbol::toLower(e.name)] = e.access;  // tB B08a
                    break;
                }
                default:
                    break;
            }
        }
        // 收集 Implements 列表
        for (auto& impl : module.implements) {
            classSym->implementsNames.push_back(impl->interfaceName);
        }
        symTab_.define(std::move(classSym));
    }

    // P22: 收集 DefType 声明 (字母->隐式类型映射)
    for (auto& dt : module.defTypes) {
        defTypeActive_ = true;
        Vb6Type dtType = Vb6Type::Variant;
        switch (dt->defKind) {
            case DefTypeKind::Bool: dtType = Vb6Type::Boolean; break;
            case DefTypeKind::Byte: dtType = Vb6Type::Byte; break;
            case DefTypeKind::Int:  dtType = Vb6Type::Integer; break;
            case DefTypeKind::Lng:  dtType = Vb6Type::Long; break;
            case DefTypeKind::Cur:  dtType = Vb6Type::Currency; break;
            case DefTypeKind::Sng:  dtType = Vb6Type::Single; break;
            case DefTypeKind::Dbl:  dtType = Vb6Type::Double; break;
            case DefTypeKind::Date: dtType = Vb6Type::Date; break;
            case DefTypeKind::Str:  dtType = Vb6Type::String; break;
            case DefTypeKind::Obj:  dtType = Vb6Type::Object; break;
            case DefTypeKind::Var:  dtType = Vb6Type::Variant; break;
        }
        for (auto& range : dt->ranges) {
            char from = toupper(range.from);
            char to = toupper(range.to);
            for (char c = from; c <= to; c++) {
                int idx = c - 'A';
                if (idx >= 0 && idx < 26) defTypeMap_[idx] = dtType;
            }
        }
    }

    // --- Pass 1: 收集所有模块级声明 ---
    pass_ = 1;
    if (verbose_) {
        std::cerr << "[Sem] Pass 1: collecting declarations..." << std::endl;
    }
    // Fix 049: 预扫描模块级 Type/Enum 声明。
    // VB6 允许模块级 Dim/Const/参数 在类型声明 (Private Type/Public Enum) 之前使用该类型,
    // 编译器会先做整模块的类型收集. 若按语句顺序线性处理 (如 cAsyncSocket.cls 中
    // "Private m_uWindowState() As UcsHelperWindowStateType" 位于该 Type 声明之前),
    // resolveTypeRef 找不到符号而静默回退 Variant, 导致: 1) 生成 VB6_SA_AT(vb6_VARIANT,...)
    // 的字段访问; 2) 对 Variant 收件人的 .成员 解析退化为跨模块类查找 (如 .Pos 误解析到
    // cToast.Pos, 参数个数不符, MSVC C2198/C2039/C2223 错误)。
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::TypeDecl || decl->kind == ASTNodeKind::EnumDecl ||
            decl->kind == ASTNodeKind::DelegateDecl) {
            dispatchDecl(*decl, *this);
        }
    }
    for (auto& decl : module.declarations) {
        registerDecl(*decl);
    }
    // 处理 Option 语句
    for (auto& opt : module.options) {
        visit(*opt);
    }

    // --- Pass 2: 分析过程体 ---
    pass_ = 2;
    if (verbose_) {
        std::cerr << "[Sem] Pass 2: analyzing procedure bodies..." << std::endl;
    }
    for (auto& decl : module.declarations) {
        dispatchDecl(*decl, *this);
    }

    // Pass 2 结束后，检查未引用的变量 (仅警告)
    if (verbose_) {
        checkUnreferencedSymbols();
    }

    // Pass 2 结束后，验证 Implements 语句
    // 成员级 `Implements I.M` 子句 (B02b) 由新式比对逐条认领, 认领不到的在下面兜底报错
    std::set<IfaceClauseRef> boundClauses;
    if (module.isClassModule && !module.implements.empty()) {
        auto* classSym = symTab_.lookupModule(module.moduleName);
        if (classSym && classSym->kind == SymbolKind::Class) {
            for (auto& impl : module.implements) {
                const std::string& ifaceName = impl->interfaceName;
                // Interface 契约 (tB, B02): 名字命中新式接口登记表 → 严格比对 (error 级),
                // 不再走下面的 legacy 路径 (D5 分叉); 未命中 → legacy 逻辑一行不动.
                if (ifaceReg_) {
                    auto found = ifaceReg_->find(Symbol::toLower(ifaceName));
                    if (found != ifaceReg_->end()) {
                        // VB6 风格接口 (.cls 宿主, Pass A2 登记进来的): 宿主模块自己就是那个
                        // 接口类, 所以还得照 legacy 路径那样给它打 isInterface 标记 —— 打码层
                        // 靠这个标记跳过接口类的方法体生成 (接口成员本来就是无体的)。
                        if (found->second.clsHost) {
                            if (auto* hostSym = symTab_.lookupModule(found->second.name)) {
                                hostSym->isInterface = true;
                            }
                        }
                        checkNewStyleInterface(module, found->second, ifaceName, impl->loc,
                                               boundClauses);
                        continue;
                    }
                }
                // 查找接口类符号
                auto* ifaceSym = symTab_.lookupModule(ifaceName);
                if (!ifaceSym || ifaceSym->kind != SymbolKind::Class) {
                    diag_.warn(DiagnosticID::SemImplementsInterfaceNotFound, impl->loc,
                        "Implements: interface '" + ifaceName + "' not found");
                    continue;
                }
                // 标记接口类
                ifaceSym->isInterface = true;
                // 收集接口方法名
                std::vector<std::string> ifaceMethodNames;
                for (auto& m : ifaceSym->memberNames) {
                    std::string lower = Symbol::toLower(m);
                    ifaceMethodNames.push_back(lower);
                    ifaceSym->interfaceMethodNames.push_back(lower);
                }
                // 验证实现类包含 InterfaceName_MethodName 方法
                for (auto& m : ifaceSym->memberNames) {
                    std::string required = ifaceName + "_" + m;
                    std::string lowerRequired = Symbol::toLower(required);
                    bool found = false;
                    for (auto& cm : classSym->memberNames) {
                        if (Symbol::toLower(cm) == lowerRequired) {
                            found = true;
                            break;
                        }
                    }
                    if (!found) {
                        // 缺槽这一条与新式路径同一个号 (3012): 两条路报的是同一件事。
                        diag_.warn(DiagnosticID::SemInterfaceNotImplemented, impl->loc,
                            "Implements " + ifaceName + ": method '" + required +
                            "' not implemented in class '" + module.moduleName + "'");
                    }
                }
            }
        }
    }
    checkMemberImplementsClauses(module, boundClauses);

    return !diag_.hasErrors();
}

// 泛型 (tB, G3): 物化器在模块常规分析结束后注入特化副本的增量两遍分析.
// 与 analyze() 的声明表处理同构 (TypeDecl 预扫 → pass1 注册 → pass2 体分析),
// 但不重复 builtins/类符号/DefType/options (模块本体已处理完).
void SemanticAnalyzer::analyzeExtraDecls(const std::vector<Decl*>& decls) {
    pass_ = 1;
    for (auto* d : decls) {
        if (d->kind == ASTNodeKind::TypeDecl || d->kind == ASTNodeKind::EnumDecl ||
            d->kind == ASTNodeKind::DelegateDecl) {
            dispatchDecl(*d, *this);
        }
    }
    for (auto* d : decls) {
        registerDecl(*d);
    }
    pass_ = 2;
    for (auto* d : decls) {
        dispatchDecl(*d, *this);
    }
}

void SemanticAnalyzer::dumpSymbols(std::ostream& os) const {
    os << "=== Symbol Table Dump ===" << std::endl;

    // 遍历所有作用域的符号
    auto dumpScope = [&](const Scope* scope, const char* label) {
        os << "--- " << label << " (depth=" << symTab_.scopeDepth() << ") ---" << std::endl;
        for (auto& [key, sym] : scope->symbols()) {
            os << "  " << sym->kindName() << " " << sym->name;
            os << " : " << TypeSystem::typeToString(sym->type);
            if (sym->isArray) os << "()";
            os << " [" << (sym->access == AccessLevel::Public ? "Public" : "Private") << "]";
            if (sym->isReferenced) os << " (used)";
            else os << " (unused)";
            os << std::endl;
            // 过程参数
            if (!sym->params.empty()) {
                os << "    params: ";
                for (size_t i = 0; i < sym->params.size(); i++) {
                    if (i > 0) os << ", ";
                    auto& p = sym->params[i];
                    os << (p.isByVal ? "ByVal " : "ByRef ") << p.name;
                    os << " As " << TypeSystem::typeToString(p.type);
                    if (p.isOptional) os << " Optional";
                    if (p.isParamArray) os << " ParamArray";
                }
                os << std::endl;
            }
            // 常量值
            if (sym->hasConstValue) {
                os << "    value: ";
                switch (sym->constType) {
                    case Vb6Type::Boolean:
                        os << (sym->constBoolValue ? "True" : "False");
                        break;
                    case Vb6Type::String:
                        os << "\"" << sym->constStringValue << "\"";
                        break;
                    case Vb6Type::Double:
                    case Vb6Type::Single:
                        os << sym->constFloatValue;
                        break;
                    default:
                        os << sym->constIntValue;
                        break;
                }
                os << std::endl;
            }
        }
    };

    // 模块级符号
    dumpScope(symTab_.moduleScope(), "Module Scope");

    os << "=== End Symbol Table ===" << std::endl;
}

} // namespace vb6c3
