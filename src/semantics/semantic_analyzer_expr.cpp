#include "semantics/semantic_analyzer.hpp"
#include <algorithm>
#include <cctype>
#include <tuple>
#include <initializer_list>
#include "semantics/semantic_analyzer_internal.h"

namespace vb6c3 {

// --- semantic_analyzer_expr.cpp: 表达式 Visitor + 类型引用 Visitor ---


// 虚方法 (tB, B08d): 裸名 (无 Me. 前缀) 引用自家可覆盖成员仍然判死。B08d 已经把虚表发了,
// 但只接在**类成员调用**那条解析路上 (resolveClassMemberCall → 间接调用); 裸名走模块作用域的
// 过程符号, 那条路发的是直接调用 —— 留着它会静默绑成本类实现, 与 B07b "继承成员必须写 Me."
// 是同一条边界。B08d 之前这里报的是"类虚表还没发", 现在报的是"这个形状没接派发"。
namespace {
std::string bareVirtualCallMsg(const std::string& cls, const std::string& name) {
    return "Bare (unqualified) call to overridable member '" + name + "' inside class '" + cls +
           "' would statically bind to '" + cls + "." + name + "' (this build dispatches only"
           " Me." + name + " / obj." + name + " calls; ai/022 B08d). Write Me." + name;
}
} // namespace

// ============================================================
// 表达式 Visitor
// ============================================================

Vb6Type SemanticAnalyzer::analyzeExpr(Expr& expr) {
    lastExprType_ = Vb6Type::Unknown;
    if (verbose_) std::cerr << "[Sem]       expr: " << expr.kindName() << std::endl;
    dispatchExpr(expr, *this);
    return lastExprType_;
}

void SemanticAnalyzer::visit(BinaryExpr& node) {
    Vb6Type leftType = analyzeExpr(*node.left);
    Vb6Type rightType = analyzeExpr(*node.right);

    // 根据运算符推导结果类型
    switch (node.op) {
        case BinaryOp::Concat:
            // & 运算符: 字符串连接
            lastExprType_ = Vb6Type::String;
            break;
        case BinaryOp::Eq: case BinaryOp::Neq:
        case BinaryOp::Lt: case BinaryOp::Gt:
        case BinaryOp::Le: case BinaryOp::Ge:
        case BinaryOp::Like: case BinaryOp::Is:
            // 比较运算符 → Boolean
            lastExprType_ = Vb6Type::Boolean;
            break;
        case BinaryOp::And: case BinaryOp::Or: case BinaryOp::Xor:
        case BinaryOp::Eqv: case BinaryOp::Imp:
            // 逻辑运算符: 如果两边都是Boolean → Boolean, 否则 → 数值提升
            if (leftType == Vb6Type::Boolean && rightType == Vb6Type::Boolean) {
                lastExprType_ = Vb6Type::Boolean;
            } else if (TypeSystem::isNumeric(leftType) && TypeSystem::isNumeric(rightType)) {
                lastExprType_ = TypeSystem::promote(leftType, rightType);
            } else {
                lastExprType_ = Vb6Type::Variant;
            }
            break;
        case BinaryOp::Add: case BinaryOp::Sub:
        case BinaryOp::Mul: case BinaryOp::Div:
        case BinaryOp::IntDiv: case BinaryOp::Mod: case BinaryOp::Pow:
            // 算术运算符 → 数值提升
            if (TypeSystem::isNumeric(leftType) && TypeSystem::isNumeric(rightType)) {
                lastExprType_ = TypeSystem::promote(leftType, rightType);
                // 整除和取模 → Long
                if (node.op == BinaryOp::IntDiv || node.op == BinaryOp::Mod) {
                    lastExprType_ = Vb6Type::Long;
                }
            } else if (leftType == Vb6Type::String && rightType == Vb6Type::String && node.op == BinaryOp::Add) {
                // 字符串 + 字符串 → String (VB6允许但推荐用&)
                lastExprType_ = Vb6Type::String;
            } else {
                lastExprType_ = Vb6Type::Variant;
            }
            break;
    }
}

void SemanticAnalyzer::visit(UnaryExpr& node) {
    Vb6Type operandType = analyzeExpr(*node.operand);

    switch (node.op) {
        case UnaryOp::Negate:
            // -x: 数值类型不变, Boolean → Integer
            if (operandType == Vb6Type::Boolean) {
                lastExprType_ = Vb6Type::Integer;
            } else if (TypeSystem::isNumeric(operandType)) {
                lastExprType_ = operandType;
            } else {
                lastExprType_ = Vb6Type::Variant;
            }
            break;
        case UnaryOp::Not:
            // Not x: Boolean → Boolean, 数值 → 数值(按位取反)
            if (operandType == Vb6Type::Boolean) {
                lastExprType_ = Vb6Type::Boolean;
            } else if (TypeSystem::isIntegral(operandType)) {
                lastExprType_ = operandType;
            } else {
                lastExprType_ = Vb6Type::Variant;
            }
            break;
    }
}

void SemanticAnalyzer::visit(LiteralExpr& node) {
    switch (node.literalKind) {
        case LiteralKind::Integer:  lastExprType_ = Vb6Type::Integer; break;
        case LiteralKind::Long:     lastExprType_ = Vb6Type::Long; break;
        case LiteralKind::LongPtr:  lastExprType_ = Vb6Type::LongPtr; break;
        case LiteralKind::Single:   lastExprType_ = Vb6Type::Single; break;
        case LiteralKind::Double:   lastExprType_ = Vb6Type::Double; break;
        case LiteralKind::Currency: lastExprType_ = Vb6Type::Currency; break;
        case LiteralKind::Decimal:  lastExprType_ = Vb6Type::Decimal; break;
        case LiteralKind::String:   lastExprType_ = Vb6Type::String; break;
        case LiteralKind::Date:     lastExprType_ = Vb6Type::Date; break;
        case LiteralKind::Boolean:  lastExprType_ = Vb6Type::Boolean; break;
        case LiteralKind::Nothing:  lastExprType_ = Vb6Type::Object; break;
        case LiteralKind::Empty:    lastExprType_ = Vb6Type::Empty; break;
        case LiteralKind::Null:     lastExprType_ = Vb6Type::Null; break;
    }
}

void SemanticAnalyzer::visit(IdentifierExpr& node) {
    std::string lower = Symbol::toLower(node.name);

    // 在符号表中查找
    auto* sym = symTab_.lookup(node.name);
    if (sym) {
        sym->isReferenced = true;
        // 虚方法 (tB, B08d): 裸名引用自家的可覆盖成员 —— 只有 Me./obj. 那条解析路接了派发。
        // 例外: 函数体里 `Name = …` 是**返回值赋值** (VB6 惯例), 不是调用, 不能算。
        if (pass_ == 2 && sym->kind != SymbolKind::Variable &&
            !(currentProc_ && Symbol::toLower(currentProc_->name) == lower) &&
            virtualCallNeedsDispatch(node.name)) {
            diag_.error(DiagnosticID::SemVirtualNotSupported, node.loc,
                bareVirtualCallMsg(currentModule_ ? currentModule_->moduleName : std::string(),
                                   node.name));
        }
        lastExprType_ = sym->type;
    } else {
        // 未找到标识符
        // <vbeclipse>: 但"工程级已经有这个名字"不算未找到 —— 标准模块的 Public 过程
        // (裸名位) 与模块名 (`Mod.成员` 的限定符位) 都不能落成下面的隐式 Variant 声明，
        // 判据与两种后果都写在 SemanticAnalyzer::namesProjectLevel 的声明处。
        // 放在最前面：Option Explicit 那条 3001 警告同样不该为这两个位出。
        if (pass_ == 2 && namesProjectLevel(node.name)) {
            // 什么都不做 —— 不是变量，也不是未声明标识符；名字的含义由发码层按工程解析。
        } else if (pass_ == 2 && declaredByAncestor(node.name)) {
            diag_.error(DiagnosticID::SemInheritsNotSupported, node.loc,
                "Inherited member '" + node.name + "' cannot be called unqualified in this build"
                " (write Me." + node.name + "; v1 merges inherited members onto the class symbol only)");
        } else if (lower == "mybase" && currentModule_ && currentModule_->isClassModule) {
            // tB Inherits (ai/022 B09): `MyBase` 不是标识符, 由发码层按名字接管 (去虚化基调用),
            // 用错位置在那里报 VB3028。这里不出"未声明的标识符", 否则每条合法写法都配一条噪声。
        } else if (pass_ == 2 && reportIfPackageBlocked(node.name, node.loc)) {
            // ai/023 S03: 包内未导出成员 (VB7006 已报), 不再叠加 3001 警告
        } else if (optionExplicit_ && pass_ == 2) {
            diag_.warn(DiagnosticID::SemUndeclaredIdentifier, node.loc,
                "未声明的标识符: '" + node.name + "' (可能来自其他模块)");
        } else if (pass_ == 2 && !optionExplicit_ && currentProc_ && currentModule_) {
            // Fix <vbeclipse>: VB6 隐式变量声明 — 工程未写 Option Explicit 时,
            // 首次使用的裸标识符按 Variant 局部变量成立 (PopupMenu.cls 的
            // Key/Text/msf_hilite 即真实案例)。登记进本过程作用域与隐式表
            // (供发码侧在过程序言预声明 C 局部), 不再留成未定义裸名。
            auto v = std::make_unique<Symbol>(SymbolKind::Variable, node.name,
                Vb6Type::Variant, node.loc, AccessLevel::Private);
            v->isReferenced = true;
            symTab_.define(std::move(v));
            symTab_.addImplicitVar(Symbol::toLower(currentModule_->moduleName),
                                   Symbol::toLower(currentProc_->name), node.name);
        }
        lastExprType_ = Vb6Type::Variant;  // 宽松模式: 推导为Variant
    }
}

void SemanticAnalyzer::visit(MemberAccessExpr& node) {
    // <vbeclipse>: 接收者子表达式站在"限定符位"上 (见 visit(IdentifierExpr) 的豁免条)。
    // 嵌套 A.B.C 由本处的 save/restore 自然传递: 外层先把 ctx 置真, 内层 MemberAccess
    // 再自己覆盖一次并复原，最左的那个裸标识符最终看到的正是"我在限定符位上"。
    const bool savedMemberObjCtx = memberObjCtx_;
    memberObjCtx_ = true;
    Vb6Type objType = analyzeExpr(*node.object);
    memberObjCtx_ = savedMemberObjCtx;
    // ai/084a M1/M2: 成员访问级别守卫 —— obj 为 Me / 类类型变量时解析接收者类,
    // Private 成员越界报 3028 (家族内放行; 解析不出接收者 → 静默, 维持旧行为)。
    checkMemberAccessGuard(*node.object, node.memberName, node.loc);
    // 虚方法 (tB, B08d): `Me.<可覆盖成员>` 与 `对象变量.<成员>` 都交给发码层的间接调用
    // (CCodeGen::virtDispatchCallee 按 stage 3.4b 的槽表定槽), 语义层不再介入 —— 这里曾有的
    // B08b 拒绝判定已按 D32① 删除: 留着判定再另起一条发码路会得到永远到不了的码。
    // 类继承 (tB, B08c): 这里是 obj.<成员> 唯一的必经点 (读、写、Set/Let、Call 与调用的
    // callee 都从这里过), Protected 越权就判在这。
    if (node.object) checkProtectedVisibility(*node.object, node.memberName, node.loc);
    // P20-21: 如果object是UDT, 查找成员类型
    if (objType == Vb6Type::UserDefinedType) {
        // 查找UDT符号获取成员类型
        if (auto* ident = dynamic_cast<IdentifierExpr*>(node.object.get())) {
            auto* udtSym = symTab_.lookup(ident->name);
            if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                std::string memberLower = node.memberName;
                std::transform(memberLower.begin(), memberLower.end(), memberLower.begin(), ::tolower);
                for (const auto& mi : udtSym->udtMembers) {
                    std::string miLower = mi.name;
                    std::transform(miLower.begin(), miLower.end(), miLower.begin(), ::tolower);
                    if (miLower == memberLower) {
                        lastExprType_ = mi.type;
                        return;
                    }
                }
            }
        }
    }
    // P24-04: ComModule全局函数访问 (VBMAN.Version)
    if (auto* ident = dynamic_cast<IdentifierExpr*>(node.object.get())) {
        auto* modSym = symTab_.lookup(ident->name);
        if (modSym && modSym->kind == SymbolKind::ComModule) {
            std::string memberLower = node.memberName;
            std::transform(memberLower.begin(), memberLower.end(), memberLower.begin(), ::tolower);
            auto it = modSym->comModuleFunctions.find(memberLower);
            if (it != modSym->comModuleFunctions.end()) {
                lastExprType_ = it->second.returnType;
                return;
            }
        }
    }
    // P24-04: ComGlobalNs — VB_GlobalNameSpace promoted函数访问 (如 VBMAN.Version)
    // VBMAN是提升到全局的函数, VBMAN()返回cVBMAN COM对象, .Version是cVBMAN的方法
    // 所以 VBMAN.Version = VBMAN().Version, 即先调用 promoted函数获取对象, 再访问成员
    if (auto* ident = dynamic_cast<IdentifierExpr*>(node.object.get())) {
        auto* gnsSym = symTab_.lookup(ident->name);
        if (gnsSym && gnsSym->kind == SymbolKind::ComGlobalNs) {
            // promoted函数返回Object (COM对象), 成员访问走晚绑定
            lastExprType_ = Vb6Type::Variant;
            return;
        }
    }
    // 简化: 非UDT或未找到成员, 推导为Variant
    lastExprType_ = Vb6Type::Variant;
}

// ai/084a M1/M2: 接收者解析 + Private 越界裁决。
// 数据来自 driver 预计算表 (memberAccess_, 见 symbol_table.hpp) —— 不能 visit 期查符号表:
// 跨模块类符号 (extSym) 的成员表 stage 3.5 才注入, 与 023 S03 预计算屏蔽表同因同解。
// 口径 (v1):
//   - 只裁决 Private; Friend 同工程内等同 Public (真实约束在跨包边界, 留给 M3);
//   - 定义类 == 当前类, 或定义类在当前类继承链上 (tB 继承合并模型) → 家族内放行;
//   - 接收者解析不出 / 类名不在表 / 成员不在表 → 静默, 维持旧行为 (不发明新报错)。
void SemanticAnalyzer::checkMemberAccessGuard(const Expr& obj, const std::string& memberName,
                                              SourceLocation loc) {
    if (!memberAccess_ || memberAccess_->empty() || !currentModule_ || pass_ != 2) return;

    std::string clsLower;
    if (dynamic_cast<const MeExpr*>(&obj)) {
        if (!currentModule_->isClassModule) return;  // SemInvalidUseOfMe 另有判定, 不掺和
        clsLower = Symbol::toLower(currentModule_->moduleName);
    } else if (auto* ident = dynamic_cast<const IdentifierExpr*>(&obj)) {
        const Symbol* varSym = symTab_.lookup(ident->name);
        // 只认带类型名的变量/参数 (Dim x As Foo); 裸类名/UDT/COM 符号无 variableTypeName → 放行
        if (!varSym || varSym->variableTypeName.empty()) return;
        clsLower = Symbol::toLower(varSym->variableTypeName);
    } else {
        return;  // 链式/表达式接收者: v1 不裁决
    }
    if (clsLower.empty()) return;

    auto clsIt = memberAccess_->find(clsLower);
    if (clsIt == memberAccess_->end()) return;   // 非工程内类模块 (UDT/COM/窗体内置类型)
    auto memIt = clsIt->second.find(Symbol::toLower(memberName));
    if (memIt == clsIt->second.end()) return;    // 表里没有 → 旧路径 (未定义名另有诊断)
    const MemberAccessEntry& entry = memIt->second;
    const std::string curLower = Symbol::toLower(currentModule_->moduleName);
    // 家族判定 (Protected 用, 对称): 定义类 == 当前类, 或互在对方继承链上。
    // Private 不用对称版 —— 派生类摸祖先 Private **字段**合法 (B07b 合并), 反向不合法。
    auto inOwnChain = [&](const std::string& defLower) -> bool {
        if (!clsreg_ || clsreg_->empty() || !currentModule_->isClassModule) return false;
        auto self = clsreg_->find(curLower);
        if (self != clsreg_->end()) {
            const auto& chain = self->second.chain;
            if (std::find(chain.begin(), chain.end(), defLower) != chain.end()) return true;
        }
        return false;
    };

    // ai/084a M3: 包导出类 Friend 成员的跨包边界 —— 消费方包 != 定义包 且清单
    // 未开 Friend=True → 7008 (与 S03 的模块级 Friend 屏蔽同构, 闭合 023 v1 边界)。
    // 同包内、包→包 Friend=True 时不受限; 宿主工程模块 (packageName 空) 属包外。
    if (entry.level == AccessLevel::Friend && !entry.definingPkg.empty()
        && Symbol::toLower(currentModule_->packageName) != entry.definingPkg
        && !entry.pkgFriendOpen) {
        diag_.error(DiagnosticID::VbpPackageFriendMember, loc,
            "成员 '" + memberName + "' 在包 '" + entry.definingPkg
            + "' 的类 '" + clsLower + "' 中是 Friend, 清单未声明 Friend=True, 不能从包外访问");
        return;
    }
    if (entry.level == AccessLevel::Protected) {
        // 3023 落地 (084a): 家族内放行, 家族外经 obj. 硬错
        if (entry.definingModuleLower == curLower) return;
        if (inOwnChain(entry.definingModuleLower)) return;
        auto defIt = clsreg_ ? clsreg_->find(entry.definingModuleLower) : clsreg_->end();
        if (clsreg_ && defIt != clsreg_->end()) {
            const auto& dchain = defIt->second.chain;
            if (std::find(dchain.begin(), dchain.end(), curLower) != dchain.end()) return;
        }
        diag_.error(DiagnosticID::SemProtectedOutsideFamily, loc,
            "成员 '" + memberName + "' 是 Protected, 接收者类 '" + clsLower
            + "' 不在当前类的家族里");
        return;
    }
    if (entry.level != AccessLevel::Private) return;

    // 家族内放行 (先比对当前类名, 再查继承链)。仅**字段**: 祖先 Private 字段经
    // B07b inhFields 合并进派生类, 派生类经 Me. 访问合法; Private 过程不在合并集
    // (转发桩剔除 Private), 派生类访问会在发码期落空 → 这里必须拦下。
    if (entry.definingModuleLower == curLower) return;
    if (entry.isField && inOwnChain(entry.definingModuleLower)) return;
    diag_.error(DiagnosticID::SemPrivateOutsideClass, loc,
        "成员 '" + memberName + "' 在类 '" + clsLower + "' 中是 Private, 不能在此访问");
}

void SemanticAnalyzer::visit(DictionaryAccessExpr& node) {
    Vb6Type objType = analyzeExpr(*node.object);
    lastExprType_ = Vb6Type::Variant;
}

void SemanticAnalyzer::visit(IndexOrCallExpr& node) {
    // VB3043 (<vbeclipse>, 用户在真 VB6 里实测确认): 数组槽实参必须是**数组表达式**。
    // `UBound(vbNull)` 在 VB6 是编译期错误 (提示缺少数组), 不是某个返回值。C3 此前把常量
    // 折成整数塞进 vb6_UBound/vb6_Join 的 SafeArray1D* 形参 —— 编译绿、运行期解引用地址 1
    // → 0xC0000005 (实测 UBound(vbNull) 与 Join(vbNull, ",") 两条)。
    // 只拦"字面量 / 常量"这一形 (VB6 里数组不可能是常量), 变量与属性实参一律放过,
    // 避免误伤既有工程。数组槽 = 注册为 Variant|Array 的那三个首参 (builtin_funcs.inc)。
    if (pass_ == 2 && node.callee && node.callee->kind == ASTNodeKind::IdentifierExpr &&
        !node.positional.empty()) {
        std::string fnName = static_cast<IdentifierExpr&>(*node.callee).name;
        std::string fnLower = Symbol::toLower(fnName);
        if (fnLower == "ubound" || fnLower == "lbound" || fnLower == "join") {
            Expr* a0 = node.positional[0].get();
            bool constNotArray = (a0->kind == ASTNodeKind::LiteralExpr);
            std::string a0Name;
            if (!constNotArray && a0->kind == ASTNodeKind::IdentifierExpr) {
                auto& id = static_cast<IdentifierExpr&>(*a0);
                Symbol* s = symTab_.lookup(id.name);
                if (s && !s->isArray &&
                    (s->kind == SymbolKind::Constant || s->kind == SymbolKind::EnumMember)) {
                    constNotArray = true;
                    a0Name = id.name;
                }
            }
            if (constNotArray) {
                diag_.error(DiagnosticID::SemArrayArgExpected, node.loc,
                    fnName + " needs an array expression; got " +
                    (a0Name.empty() ? std::string("a literal") : "constant '" + a0Name + "'") +
                    " (VB6 rejects this at compile time)");
            }
        }
    }

    // 分析被调用者
    Vb6Type calleeType = Vb6Type::Unknown;
    bool argsAnalyzed = false;

    // 判断是函数调用还是数组索引
    if (auto* ident = dynamic_cast<IdentifierExpr*>(node.callee.get())) {
        auto* sym = symTab_.lookup(ident->name);
        if (sym) {
            sym->isReferenced = true;
            // 虚方法 (tB, B08d): 这条分支自己查符号、不走 visit(IdentifierExpr), 所以裸名
            // **带实参**的调用要在这里补同一条判定 (属性/方法的 `Name(…)` 形态)。
            if (pass_ == 2 && sym->kind != SymbolKind::Variable &&
                virtualCallNeedsDispatch(ident->name)) {
                diag_.error(DiagnosticID::SemVirtualNotSupported, node.loc,
                    bareVirtualCallMsg(currentModule_ ? currentModule_->moduleName : std::string(),
                                       ident->name));
            }

            if (sym->ovlCount > 0 &&
                (sym->kind == SymbolKind::Function || sym->kind == SymbolKind::Sub)) {
                // 重载组: 先算实参类型, 打分选变体, 结果后缀记入节点供 cgen 定形
                std::vector<Vb6Type> argT;
                for (auto& arg : node.positional) argT.push_back(analyzeExpr(*arg));
                for (auto& namedArg : node.named) analyzeExpr(*namedArg.value);
                argsAnalyzed = true;
                std::string suffix;
                Symbol* w = resolveOverload(sym, argT, node.loc, suffix);
                node.calleeOvlSuffix = suffix;
                if (w) {
                    w->isReferenced = true;
                    calleeType = (w->kind == SymbolKind::Function)
                                 ? w->type : Vb6Type::Variant;
                }
            } else if (sym->kind == SymbolKind::Variable &&
                lookupDelegateSym(sym->variableTypeName)) {
                // 委托变量直调: op(5, 6) — 按委托签名校验实参, 返回类型取委托声明.
                auto* del = lookupDelegateSym(sym->variableTypeName);
                node.isDelegateCall = true;
                node.delegateTypeName = del->name;
                checkCallArgs(del, node);
                if (!node.named.empty()) {
                    diag_.error(DiagnosticID::SemTypeMismatch, node.loc,
                        "Delegate 直调暂不支持命名参数 (v1)");
                }
                calleeType = (del->delegateProcKind == ProcKind::Function)
                             ? del->delegateReturnType : Vb6Type::Variant;
            } else if (sym->kind == SymbolKind::Function ||
                sym->kind == SymbolKind::DeclareFunc) {
                // 函数调用: 返回类型即为表达式类型
                calleeType = sym->type;
                checkCallArgs(sym, node);
            } else if (sym->kind == SymbolKind::Sub ||
                       sym->kind == SymbolKind::DeclareSub) {
                // Sub作为表达式调用 (VB6允许, 返回Variant)
                calleeType = Vb6Type::Variant;
                checkCallArgs(sym, node);
            } else if (sym->isArray) {
                // 数组索引: 元素类型
                uint16_t baseType = static_cast<uint16_t>(sym->type) &
                                    ~static_cast<uint16_t>(Vb6Type::Array);
                calleeType = static_cast<Vb6Type>(baseType);
            } else {
                calleeType = sym->type;
            }
        } else {
            // 可能是内置函数 (MsgBox, Len, etc.) 或未声明
            // 简化: 推导为Variant
            calleeType = Vb6Type::Variant;

            // ai/023 S03: 包内未导出成员 → 专用诊断 (VB7006), 跳过延后登记
            bool pkgBlocked = reportIfPackageBlocked(ident->name, node.loc);

            // O3 延后: 跨模块 Public 重载组要到 runCrossModuleResolution (本模块
            // 分析之后) 才注入, 此处查无此名. 记下 节点+实参类型, 待 Driver 注入
            // 完成后补跑 resolveOverload 写 calleeOvlSuffix. 仅 pass 2 记录
            // (两遍去重); 无实参的裸名引用不是调用, 不记.
            if (!pkgBlocked && pass_ == 2 &&
                (!node.positional.empty() || !node.named.empty())) {
                std::vector<Vb6Type> argT;
                std::vector<bool> argArr;
                for (auto& arg : node.positional) {
                    Vb6Type t = analyzeExpr(*arg);
                    bool arr = (static_cast<uint16_t>(t) &
                                static_cast<uint16_t>(Vb6Type::Array)) != 0;
                    if (!arr) {
                        if (auto* aid = dynamic_cast<IdentifierExpr*>(arg.get())) {
                            Symbol* as = symTab_.lookup(aid->name);
                            if (as && as->isArray) arr = true;
                        }
                    }
                    argT.push_back(t);
                    argArr.push_back(arr);
                }
                for (auto& namedArg : node.named) analyzeExpr(*namedArg.value);
                deferredXmodCalls_.push_back(
                    {&node, ident->name, std::move(argT), std::move(argArr), node.loc});
                argsAnalyzed = true;  // 上一步已分析, 跳过函数尾的重复遍历
            }

            // 检查Option Explicit
            if (optionExplicit_ && pass_ == 2 &&
                node.positional.empty()) {
                // 非数组索引、无参数 → 可能是未声明变量
                // 但无法确定, 这里不做报错
            }
        }
    } else {
        // 成员调用 (obj.method args)
        analyzeExpr(*node.callee);
        calleeType = Vb6Type::Variant;
    }

    // 分析参数 (重载分支已在选择前分析过, 避免双重求值的副作用)
    if (!argsAnalyzed) {
        for (auto& arg : node.positional) {
            analyzeExpr(*arg);
        }
        for (auto& namedArg : node.named) {
            analyzeExpr(*namedArg.value);
        }
    }

    lastExprType_ = calleeType;
}

void SemanticAnalyzer::visit(NewExpr& node) {
    // ai/023 S04: 非导出包类 → VB7006。必须在此拦: 类符号不在表里时
    // codegen 会静默退化成后期绑定 COM (`vb6_NewObject(L"Cls")`),
    // 编译通过、运行期才失败 —— 比符号未解析更难查。
    reportIfPackageClassBlocked(node.className, node.loc);
    // 查找类名是否在符号表中注册(类符号或作为对象类型引用)
    auto* sym = symTab_.lookup(node.className);
    if (sym && sym->kind == SymbolKind::Class) {
        node.className = sym->name;  // 规范化大小写
    }
    // 实参表达式照常分析 (嵌套 New/引用标记等都依赖)
    for (auto& a : node.args) analyzeExpr(*a);
    // ai/084c: 带参构造元数校验。判定源是预计算表 ctorParams_ (类名→Class_Initialize
    // 形参个数), 不依赖符号表 —— 跨模块类符号 stage 3.5 才注入, visit 期查不到。
    // 在表内 = 本工程类 (含包导出类); 不在表内还带实参 = COM/外部类 → v1 明确报错
    // (运行期没有可承接的构造通路, 不能静默丢参数)。
    std::string clsLower = Symbol::toLower(node.className);
    auto itCtor = ctorParams_.find(clsLower);
    if (itCtor != ctorParams_.end()) {
        int need = itCtor->second;
        int got = static_cast<int>(node.args.size());
        if (got != need) {
            if (need == 0) {
                diag_.error(DiagnosticID::SemCtorArityMismatch, node.loc,
                    "类 '" + node.className + "' 的 Class_Initialize 不带参数, New 不能提供实参");
            } else if (got == 0) {
                diag_.error(DiagnosticID::SemCtorArityMismatch, node.loc,
                    "类 '" + node.className + "' 的 Class_Initialize 带 " + std::to_string(need)
                    + " 个参数, 必须写成 New " + node.className + "(...)");
            } else {
                diag_.error(DiagnosticID::SemCtorArityMismatch, node.loc,
                    "New " + node.className + " 实参个数不符: 需要 " + std::to_string(need)
                    + " 个, 实给 " + std::to_string(got) + " 个");
            }
        }
    } else if (!node.args.empty()) {
        diag_.error(DiagnosticID::SemCtorArityMismatch, node.loc,
            "'" + node.className + "' 不是本工程类, New 不支持带参构造 (COM 对象经 Set x = New X 创建)");
    }
    // 即使未找到类符号，也不报错(可能是COM对象或外部类，运行时解析)
    lastExprType_ = Vb6Type::Object;
}

void SemanticAnalyzer::visit(TypeOfExpr& node) {
    analyzeExpr(*node.object);
    lastExprType_ = Vb6Type::Boolean;
}

void SemanticAnalyzer::visit(AddressOfExpr& node) {
    markReferenced(node.funcName);
    // 委托绑定的 AddressOf 产出委托值 (指针宽度); 裸 AddressOf 保持历史口径 Long.
    lastExprType_ = node.delegateTypeName.empty() ? Vb6Type::Long : Vb6Type::LongPtr;
}

// ============================================================
// Delegate 辅助 (tB 扩展)
// ============================================================

Symbol* SemanticAnalyzer::lookupDelegateSym(const std::string& typeName) {
    if (typeName.empty()) return nullptr;
    auto* s = symTab_.lookupModule(typeName);
    return (s && s->kind == SymbolKind::Delegate) ? s : nullptr;
}

bool SemanticAnalyzer::checkDelegateSignature(Symbol* del, Symbol* proc, SourceLocation loc) {
    auto fail = [&](const std::string& msg) {
        diag_.error(DiagnosticID::SemTypeMismatch, loc, msg);
        return false;
    };
    if (del->delegateProcKind == ProcKind::Function) {
        if (proc->kind != SymbolKind::Function)
            return fail("Delegate '" + del->name + "' 是 Function 签名, 目标 '" +
                        proc->name + "' 不是 Function");
        if (proc->type != del->delegateReturnType)
            return fail("Delegate '" + del->name + "' 返回类型不符: 委托为 " +
                        std::string(TypeSystem::typeToString(del->delegateReturnType)) +
                        ", 目标为 " + std::string(TypeSystem::typeToString(proc->type)));
    } else {
        if (proc->kind != SymbolKind::Sub)
            return fail("Delegate '" + del->name + "' 是 Sub 签名, 目标 '" +
                        proc->name + "' 不是 Sub");
    }
    if (del->params.size() != proc->params.size())
        return fail("Delegate '" + del->name + "' 参数个数不符: 委托 " +
                    std::to_string(del->params.size()) + " 个, 目标 " +
                    std::to_string(proc->params.size()) + " 个");
    for (size_t i = 0; i < del->params.size(); i++) {
        const auto& d = del->params[i];
        const auto& p = proc->params[i];
        if (d.type != p.type)
            return fail("Delegate '" + del->name + "' 第" + std::to_string(i + 1) +
                        "个参数类型不符: 委托 " + std::string(TypeSystem::typeToString(d.type)) +
                        ", 目标 '" + p.name + "' 为 " +
                        std::string(TypeSystem::typeToString(p.type)));
        if (d.isByVal != p.isByVal)
            return fail("Delegate '" + del->name + "' 第" + std::to_string(i + 1) +
                        "个参数 ByVal/ByRef 不符: 委托参数 '" + d.name + "'");
    }
    return true;
}

bool SemanticAnalyzer::matchesDelegateSignature(Symbol* del, Symbol* proc) {
    if (!del || !proc) return false;
    if (del->delegateProcKind == ProcKind::Function) {
        if (proc->kind != SymbolKind::Function) return false;
        if (proc->type != del->delegateReturnType) return false;
    } else {
        if (proc->kind != SymbolKind::Sub) return false;
    }
    if (del->params.size() != proc->params.size()) return false;
    for (size_t i = 0; i < del->params.size(); i++) {
        if (del->params[i].type != proc->params[i].type) return false;
        if (del->params[i].isByVal != proc->params[i].isByVal) return false;
    }
    return true;
}

void SemanticAnalyzer::bindDelegateAddressOf(const std::string& typeName, Expr& valueExpr,
                                             SourceLocation loc) {
    if (valueExpr.kind != ASTNodeKind::AddressOfExpr) return;
    Symbol* del = lookupDelegateSym(typeName);
    if (!del) return;
    if (currentModule_ && currentModule_->isClassModule) {
        diag_.error(DiagnosticID::SemTypeMismatch, loc,
            "Delegate 绑定暂不支持类模块 (v1): 类方法需要实例 (Me) 绑定");
        return;
    }
    auto& aof = static_cast<AddressOfExpr&>(valueExpr);

    std::string fn = aof.funcName;
    size_t dot = fn.find('.');
    if (dot != std::string::npos) {
        std::string mod = Symbol::toLower(fn.substr(0, dot));
        std::string curMod = currentModule_ ? Symbol::toLower(currentModule_->moduleName) : "";
        if (mod != curMod) {
            diag_.error(DiagnosticID::SemTypeMismatch, loc,
                "Delegate 绑定暂不支持跨模块目标: '" + fn + "'");
            return;
        }
        fn = fn.substr(dot + 1);
    }
    auto* proc = symTab_.lookupModule(fn);
    if (proc && proc->isExternal) {
        diag_.error(DiagnosticID::SemTypeMismatch, loc,
            "Delegate 绑定暂不支持跨模块目标: '" + aof.funcName + "'");
        return;
    }
    if (!proc || (proc->kind != SymbolKind::Sub && proc->kind != SymbolKind::Function)) {
        diag_.error(DiagnosticID::SemUndeclaredIdentifier, loc,
            "Delegate '" + del->name + "' 绑定目标未找到或不可绑定: '" + aof.funcName + "'");
        return;
    }
    // 重载组: 在候选集中找签名相符的变体 (无组时行为同旧: 唯一 proc 直接比对)
    if (proc->ovlCount > 0) {
        Symbol* match = nullptr;
        for (auto* c : symTab_.lookupModuleOverloads(fn)) {
            if ((c->kind == SymbolKind::Sub || c->kind == SymbolKind::Function) &&
                matchesDelegateSignature(del, c)) { match = c; break; }
        }
        if (!match) {
            checkDelegateSignature(del, proc, loc);  // 复用报告器产出一次具体不符原因
            return;
        }
        proc = match;
    } else if (!checkDelegateSignature(del, proc, loc)) {
        return;
    }

    aof.delegateTypeName = del->name;
    // 重载组内: 即使选中 head 也记录其 fp (桩名需区分具体变体); 无组时为 ""
    aof.funcOvlSuffix = proc->ovlCount > 0 ? ("$ov$" + proc->overloadFp) : "";
    proc->isReferenced = true;
    // 幂等登记桩生成需求 (同一目标变体可能被多处赋值引用).
    for (auto& t : del->delegateTargets) {
        if (Symbol::toLower(t.procName) == Symbol::toLower(proc->name) &&
            t.procFp == (proc->ovlCount > 0 ? proc->overloadFp : ""))
            return;
    }
    del->delegateTargets.push_back(
        {proc->name, "", proc->ovlCount > 0 ? proc->overloadFp : ""});
}

void SemanticAnalyzer::visit(MeExpr& node) {
    // Me代表当前对象实例 (类模块或窗体模块)
    if (currentModule_ && (currentModule_->isClassModule || currentModule_->isFormModule)) {
        lastExprType_ = Vb6Type::Object;
    } else {
        diag_.error(DiagnosticID::SemTypeMismatch, node.loc,
            "'Me' can only be used in class/form modules");
        lastExprType_ = Vb6Type::Object;
    }
}

void SemanticAnalyzer::visit(WithMemberExpr& node) {
    // .member → 访问With对象的成员
    // 简化: 推导为Variant
    lastExprType_ = Vb6Type::Variant;
}

// ============================================================
// 类型引用 Visitor
// ============================================================

void SemanticAnalyzer::visit(SimpleTypeRef& node) {
    // 类型引用不会产生表达式类型
}

void SemanticAnalyzer::visit(ArrayTypeRef& node) {
    // 分析数组维度表达式
    for (auto& dim : node.dimensions) {
        if (dim.lower) analyzeExpr(*dim.lower);
        if (dim.upper) analyzeExpr(*dim.upper);
    }
}

void SemanticAnalyzer::visit(FixedStringTypeRef& node) {
    if (node.length) analyzeExpr(*node.length);
}

// ---- 文件I/O语句 ----

void SemanticAnalyzer::visit(OpenStmt& node) {
    analyzeExpr(*node.pathName);
    if (node.recordLength) analyzeExpr(*node.recordLength);
    analyzeExpr(*node.fileNumber);
}

void SemanticAnalyzer::visit(GetStmt& node) {
    analyzeExpr(*node.fileNumber);
    if (node.recordNumber) analyzeExpr(*node.recordNumber);
    analyzeExpr(*node.varName);
}

void SemanticAnalyzer::visit(PutStmt& node) {
    analyzeExpr(*node.fileNumber);
    if (node.recordNumber) analyzeExpr(*node.recordNumber);
    analyzeExpr(*node.varName);
}
} // namespace vb6c3
