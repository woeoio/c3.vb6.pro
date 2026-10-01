#include "semantics/semantic_analyzer.hpp"
#include "semantics/interface_sig.hpp"  // tB Interface/继承线共用的小写键函数 (B07b)
#include <algorithm>
#include <cctype>
#include <cstdio>
#include <tuple>
#include <initializer_list>
#include <map>
#include <vector>
#include "semantics/semantic_analyzer_internal.h"

namespace vb6c3 {

// --- semantic_analyzer_util.cpp: 辅助方法 (赋值/实参检查、引用标记、未引用符号) + Optional 参数默认值求值 ---


// ============================================================
// 辅助方法
// ============================================================

void SemanticAnalyzer::checkAssignment(Vb6Type targetType, Vb6Type valueType,
                                        const SourceLocation& loc,
                                        const std::string& context) {
    if (!TypeSystem::canImplicitConvert(valueType, targetType)) {
        diag_.warn(DiagnosticID::SemTypeMismatch, loc,
            context + ": 无法将 " + TypeSystem::typeToString(valueType) +
            " 隐式转换为 " + TypeSystem::typeToString(targetType));
    }
}

std::string SemanticAnalyzer::computeOverloadFp(const Symbol& sym) const {
    if (sym.kind != SymbolKind::Sub && sym.kind != SymbolKind::Function) return "";
    // 类模块成员表 (memberParams/memberProcKinds/memberNames) 是单值槽,
    // 重载整组支持在 O3 改造这些表之后; 现在类模块同名保持旧"重复声明"行为。
    if (currentModule_ && currentModule_->isClassModule) return "";
    // 窗体/UserControl 生命周期与绘制回调 (Form_Load / UserControl_Resize 等)
    // 由运行期按固定名定形派发, 必须每名字唯一 → 禁入重载组 (O3 护栏)。
    // 控件事件处理器 (Command1_Click) 的名字含动态控件前缀, 无法在此静态识别,
    // 对其重载会在 MSVC 层报原型冲突 (确定性失败, 不会静默错编)。
    {
        std::string ln = Symbol::toLower(sym.name);
        if (currentModule_ && currentModule_->isFormModule &&
            (ln.compare(0, 5, "form_") == 0 || ln.compare(0, 12, "usercontrol_") == 0))
            return "";
    }
    std::string fp;
    fp += (sym.kind == SymbolKind::Function) ? "F" : "S";
    for (const auto& p : sym.params) {
        if (p.isParamArray) return "";  // fbc 规则: vararg 过程不进重载组
        fp += ':';
        fp += std::to_string(static_cast<int>(p.type));
        fp += p.isByVal ? 'b' : 'r';
        if (p.isOptional) fp += 'o';
    }
    fp += ":R";
    fp += std::to_string(static_cast<int>(sym.type));
    return fp;
}

int SemanticAnalyzer::ovlScoreParam(Vb6Type argT, const ParameterInfo& p) {
    if (argT == Vb6Type::Unknown) return 2;  // 未知实参宽松放行 (隐式转换档)
    if (p.type == argT) return 4;
    // 数值域内 widening (Integer→Long→Single→Double 等) 优先于跨域转换 (→String 等)
    if (TypeSystem::isNumeric(argT) && TypeSystem::isNumeric(p.type)) return 3;
    // Variant 实参给具体形参 = 运行期打包, 与 Variant 形参同档 (两者并存时报歧义,
    // 由用户显式 CLng/CStr 消歧 — VB 无静态类型可偏袒)
    if (argT == Vb6Type::Variant) return 1;
    if (p.type == Vb6Type::Variant) return 1;
    if (TypeSystem::canImplicitConvert(argT, p.type)) return 2;
    return 0;
}

Symbol* SemanticAnalyzer::resolveOverload(Symbol* head, const std::vector<Vb6Type>& argT,
                                          SourceLocation loc, std::string& suffixOut) {
    suffixOut.clear();
    if (!head) return nullptr;
    if (head->ovlCount == 0) return head;  // 非重载快路径, 与旧行为逐字节同
    auto cands = symTab_.lookupModuleOverloads(head->name);
    Symbol* best = nullptr;
    int bestScore = -1;
    int ties = 0;
    size_t n = argT.size();
    for (auto* s : cands) {
        if (s->kind != head->kind) continue;
        size_t required = 0;
        for (auto& p : s->params) if (!p.isOptional) required++;
        if (n < required || n > s->params.size()) continue;
        int score = 0;
        bool ok = true;
        for (size_t i = 0; i < n; i++) {
            int sc = ovlScoreParam(argT[i], s->params[i]);
            if (!sc) { ok = false; break; }
            score += sc;
        }
        if (!ok) continue;
        if (score > bestScore) { bestScore = score; best = s; ties = 1; }
        else if (score == bestScore) ties++;
    }
    if (!best) {
        diag_.error(DiagnosticID::SemTypeMismatch, loc,
            "没有与实参匹配的重载版本 '" + head->name + "' (候选 " +
            std::to_string(cands.size()) + " 个)");
        return head;
    }
    if (ties > 1) {
        diag_.error(DiagnosticID::SemTypeMismatch, loc,
            "对重载 '" + head->name + "' 的调用有歧义");
        return head;
    }
    if (!cands.empty() && best != cands.front())
        suffixOut = "$ov$" + best->overloadFp;
    return best;
}

bool SemanticAnalyzer::reportIfPackageBlocked(const std::string& name, SourceLocation loc) {
    auto it = blockedPkgNames_.find(Symbol::toLower(name));
    if (it == blockedPkgNames_.end()) return false;
    diag_.error(DiagnosticID::VbpPackageNotExported, loc,
                "'" + name + "' is not exported by package '" + it->second +
                "' (only Public members of exported modules are visible)");
    return true;
}

bool SemanticAnalyzer::reportIfPackageClassBlocked(const std::string& name,
                                                  SourceLocation loc) {
    if (blockedPkgClasses_.empty() || name.empty()) return false;
    std::string key = Symbol::toLower(name);
    auto it = blockedPkgClasses_.find(key);
    if (it == blockedPkgClasses_.end()) return false;
    if (!reportedBlockedClasses_.insert(key).second) return true;  // 已报过, 静默拦
    diag_.error(DiagnosticID::VbpPackageNotExported, loc,
                "'" + it->second + "' is a class of a package that does not export it"
                " (add 'Class=" + it->second + "; Public' to the package manifest)");
    return true;
}

void SemanticAnalyzer::resolveDeferredCrossModuleOverloads() {
    for (auto& site : deferredXmodCalls_) {
        // ai/023 S03: 包导出边界 —— 被屏蔽的名字在这里给出专用诊断 (VB7006),
        // 而不是漏成 stage-3 的 3001 警告 + cgen 期 C2xxx。
        {
            auto bit = blockedPkgNames_.find(Symbol::toLower(site.identName));
            if (bit != blockedPkgNames_.end()) {
                diag_.error(DiagnosticID::VbpPackageNotExported, site.loc,
                            "'" + site.identName + "' is not exported by package '" +
                            bit->second + "' (only Public members of exported modules)");
                continue;
            }
        }
        Symbol* head = symTab_.lookupModule(site.identName);
        // 普通 (非组) 跨模块过程在场: 调用点归它, 泛型模板不得抢占
        // (tB 优先级 非泛型 > 泛型; 模板注册符号, 与真实过程同名时只可能
        //  是模板未发符号的共存)
        bool plainHead = head && head->isExternal && head->ovlCount == 0 &&
            (head->kind == SymbolKind::Function || head->kind == SymbolKind::Sub);
        if (!head || !head->isExternal || head->ovlCount == 0 ||
            (head->kind != SymbolKind::Function && head->kind != SymbolKind::Sub)) {
            // 泛型 (tB, G3): 非跨模组 → 试模板推断绑定 (成功则改写 callee 扁名
            // 并登记实例化请求; 失败自带诊断, 站点一次性消费不重试)
            if (!plainHead) {
                GenInstRequest req;
                if (tryBindGenericCall(site, req)) genericRequests_.push_back(std::move(req));
            }
            continue;  // 未注入成组 (单模块裸调/内置/本地) → 旧行为
        }
        std::string suffix;
        Symbol* w = resolveOverload(head, site.argTypes, site.loc, suffix);
        site.node->calleeOvlSuffix = suffix;
        if (w) w->isReferenced = true;
    }
    deferredXmodCalls_.clear();
}

// ============================================================
// 泛型调用点推断 (tB 扩展, G3)
// ============================================================
// 绑定规则 (计划冻结版, 跟 tB 口径): 形参位为类型参数 P (或 P()) 时从对应
// 实参类型绑定 P; 具体形参只做宽松兼容; 全部 P 未绑齐 → 报"需显式 (Of …)".
// v1 边界 (文档化): 推断支持标量实参 (含标量数组的 T()); UDT/Object/多维数组
// 实参要求显式 (Of). 同一 P 多处绑定必须一致.
bool SemanticAnalyzer::tryBindGenericCall(DeferredXmodCallSite& site,
                                          GenInstRequest& reqOut) {
    if (!genReg_ || !site.node) return false;
    auto it = genReg_->find(Symbol::toLower(site.identName));
    if (it == genReg_->end()) return false;   // 非模板名: 静默放过 (旧行为)
    const GenTemplateView& tv = it->second;

    std::vector<std::unique_ptr<ParameterDecl>>* fparams = nullptr;
    switch (tv.decl->kind) {
    case ASTNodeKind::SubDecl:
        fparams = &static_cast<SubDecl*>(tv.decl)->params; break;
    case ASTNodeKind::FunctionDecl:
        fparams = &static_cast<FunctionDecl*>(tv.decl)->params; break;
    case ASTNodeKind::PropertyDecl:
        fparams = &static_cast<PropertyDecl*>(tv.decl)->params; break;
    default:
        return false;  // UDT 模板不经常规调用位实例化 (As 位置显式路径)
    }

    auto needExplicit = [&](const std::string& why) {
        diag_.error(DiagnosticID::SemTypeMismatch, site.loc,
            "无法推断泛型 '" + site.identName + "' 的类型实参: " + why +
            " (请用 (Of ...) 显式给出)");
        return false;
    };

    if (site.argTypes.size() != fparams->size())
        return needExplicit("实参个数与模板形参不符");

    std::map<std::string, Vb6Type> bound;  // lower(P) → 绑定类型 (已剥数组位)
    // 数组实参判定读记录期捕获的标记 (见 DeferredXmodCallSite::argIsArray):
    // 延后绑定期局部作用域已弹出, 现场回查符号不可靠.
    auto argIsArray = [&](size_t i) -> bool {
        return i < site.argIsArray.size() && site.argIsArray[i];
    };
    for (size_t i = 0; i < site.argTypes.size(); i++) {
        ParameterDecl* fp = (*fparams)[i].get();
        const ASTNode* el = fp->asType.get();
        bool formalArr = false;
        if (el && el->kind == ASTNodeKind::ArrayTypeRef) {
            auto& ar = static_cast<const ArrayTypeRef&>(*el);
            // 0 维 = 形参动态数组 arr() (rank 未定, v1 视作标量元素 T 的数组);
            // 1 维 = 定长 T(n). 两者都按 T() 位绑定; >1 维才拒 (v1 不支持).
            if (ar.dimensions.size() > 1)
                return needExplicit("多维数组形参 (v1 不支持)");
            el = ar.elementType.get();
            formalArr = true;
        }
        std::string tvar;
        if (el && el->kind == ASTNodeKind::SimpleTypeRef) {
            auto& sr = static_cast<const SimpleTypeRef&>(*el);
            for (auto& tp : tv.typeParams) {
                if (genLower(tp) == genLower(sr.name)) { tvar = genLower(tp); break; }
            }
        }
        Vb6Type at = site.argTypes[i];
        if (!tvar.empty()) {
            if (formalArr) {
                if (!argIsArray(i))
                    return needExplicit("形参为 T() 而实参不是数组");
                uint16_t bits = static_cast<uint16_t>(at);
                at = static_cast<Vb6Type>(bits & ~static_cast<uint16_t>(Vb6Type::Array));
            }
            if (at == Vb6Type::UserDefinedType || at == Vb6Type::Object ||
                at == Vb6Type::Variant || at == Vb6Type::Unknown ||
                (static_cast<uint16_t>(at) & static_cast<uint16_t>(Vb6Type::Array)))
                return needExplicit("实参类型 v1 不可推断 (UDT/Object/Variant/嵌套数组)");
            auto prev = bound.find(tvar);
            if (prev == bound.end()) bound[tvar] = at;
            else if (prev->second != at)
                return needExplicit("同一类型参数多处绑定不一致");
        } else {
            // 具体形参: 宽松兼容判定 (拒显式不可转; 其余交给物化后的既有校验)
            Vb6Type ft = resolveTypeRef(fp->asType.get());
            if (ft != Vb6Type::Variant && at != Vb6Type::Variant &&
                at != Vb6Type::Unknown && ft != at &&
                !TypeSystem::canImplicitConvert(at, ft))
                return needExplicit("实参与具体形参类型不符");
        }
    }

    std::vector<std::string> argNames;
    for (auto& tp : tv.typeParams) {
        auto b = bound.find(genLower(tp));
        if (b == bound.end())
            return needExplicit("类型参数只出现在无法推断的位置");
        argNames.push_back(TypeSystem::typeToString(b->second));
    }

    auto* id = dynamic_cast<IdentifierExpr*>(site.node->callee.get());
    if (!id) return needExplicit("调用目标不是标识符形态");
    reqOut.flat = genMakeFlatLower(site.identName, argNames);
    reqOut.base = site.identName;
    reqOut.args = argNames;
    id->name = reqOut.flat;   // 改写调用点: 下游按普通过程名解析
    return true;
}

void SemanticAnalyzer::checkCallArgs(Symbol* procSym, IndexOrCallExpr& callNode) {
    if (!procSym) return;

    // 内置函数未注册参数信息, 跳过参数校验
    if (procSym->isBuiltin) return;

    size_t expectedParams = 0;
    size_t optionalParams = 0;
    bool hasParamArray = false;

    for (auto& p : procSym->params) {
        if (p.isParamArray) {
            hasParamArray = true;
            continue;
        }
        if (p.isOptional) {
            optionalParams++;
        }
        expectedParams++;
    }

    size_t providedArgs = callNode.positional.size() + callNode.named.size();

    if (hasParamArray) {
        // ParamArray: 至少需要 required 个参数
        size_t required = expectedParams - optionalParams;
        if (providedArgs < required) {
            diag_.error(DiagnosticID::SemWrongNumberOfArguments, callNode.loc,
                "参数数量错误: '" + procSym->name + "' 至少需要 " +
                std::to_string(required) + " 个参数, 实际提供 " +
                std::to_string(providedArgs));
        }
    } else {
        size_t required = expectedParams - optionalParams;
        if (providedArgs < required || providedArgs > expectedParams) {
            diag_.error(DiagnosticID::SemWrongNumberOfArguments, callNode.loc,
                "参数数量错误: '" + procSym->name + "' 需要 " +
                std::to_string(required) + "-" + std::to_string(expectedParams) +
                " 个参数, 实际提供 " + std::to_string(providedArgs));
        }
    }

    // 命名参数检查
    for (auto& namedArg : callNode.named) {
        bool found = false;
        std::string lowerName = Symbol::toLower(namedArg.name);
        for (auto& p : procSym->params) {
            if (Symbol::toLower(p.name) == lowerName) {
                found = true;
                break;
            }
        }
        if (!found) {
            diag_.error(DiagnosticID::SemNamedArgNotFound, callNode.loc,
                "命名参数未找到: '" + namedArg.name + "' (过程 '" + procSym->name + "')");
        }
    }

    // 重复命名参数检查
    std::vector<std::string> seenNamedArgs;
    for (auto& namedArg : callNode.named) {
        std::string lowerName = Symbol::toLower(namedArg.name);
        for (auto& seen : seenNamedArgs) {
            if (Symbol::toLower(seen) == lowerName) {
                diag_.error(DiagnosticID::SemDuplicateNamedArg, callNode.loc,
                    "重复的命名参数: '" + namedArg.name + "'");
                break;
            }
        }
        seenNamedArgs.push_back(namedArg.name);
    }
}

void SemanticAnalyzer::markReferenced(const std::string& name) {
    auto* sym = symTab_.lookup(name);
    if (sym) {
        sym->isReferenced = true;
    }
}

std::string SemanticAnalyzer::makeInternalName(const std::string& prefix,
                                                const std::string& name) {
    return prefix + "_" + name;
}

void SemanticAnalyzer::checkUnreferencedSymbols() {
    // 遍历模块级符号, 检查未引用的变量/常量
    auto* modScope = symTab_.moduleScope();
    for (auto& [key, sym] : modScope->symbols()) {
        if (!sym->isReferenced &&
            (sym->kind == SymbolKind::Variable || sym->kind == SymbolKind::Constant)) {
            // 仅在 verbose 模式下警告, VB6默认允许未使用变量
            diag_.warn(DiagnosticID::SemUndeclaredIdentifier, sym->location,
                "未使用的" + std::string(sym->kindName()) + ": '" + sym->name + "'");
        }
    }
}


// ============================================================
// P14.1.4: Optional参数默认值求值
// ============================================================
std::string SemanticAnalyzer::evalOptionalDefault(ASTNode* defaultValue, Vb6Type paramType) {
    // 无显式默认值 -> 返回空字符串, cgen将使用类型零值
    if (!defaultValue) return "";
    
    // 将AST字面量表达式转换为C表达式字符串
    auto* lit = dynamic_cast<LiteralExpr*>(defaultValue);
    if (lit) {
        switch (lit->literalKind) {
            case LiteralKind::Integer:
            case LiteralKind::Long:
                return lit->rawText;  // "10", "-1" 等
            case LiteralKind::LongPtr: {
                // Fix 082: rawText 带 VB 后缀 ("&H80000000^"), 原样返回会写进生成 C →
                // C2059. 与 cgen_expr.cpp 的 LongPtr 分支同规则: LongPtr 是平台相关宽度
                // (32/64 位机分别 4/8 字节) → intptr_t; 十六进制无符号形式避免十进制
                // INT64_MIN 字面量溢出.
                std::string t = lit->rawText;
                if (!t.empty() && t.back() == '^') t.pop_back();
                if (t.size() >= 2 && t[0] == '&' && (t[1] == 'H' || t[1] == 'h')) {
                    return "((intptr_t)0x" + t.substr(2) + "ULL)";
                }
                if (t.size() >= 2 && t[0] == '&' && (t[1] == 'O' || t[1] == 'o')) {
                    return "((intptr_t)0" + t.substr(2) + "ULL)";
                }
                if (t.size() >= 2 && t[0] == '&' && (t[1] == 'B' || t[1] == 'b')) {
                    return "((intptr_t)0b" + t.substr(2) + "ULL)";
                }
                return "((intptr_t)" + t + "ULL)";
            }
            case LiteralKind::Single: {
                // Fix 133z: 单精度默认值 `1!` → C 浮点字面量 `1.0000000f`.
                // 原样返回 rawText ("1!") 会写进生成 C → 语法错误
                // (czUI.ctl: `Optional ByVal penWidth As Single = 1!` →
                // `if (!_has_penWidth) penWidth = 1!;` → C2059).
                float fv = lit->floatValue;
                std::ostringstream oss133z;
                char buf133z[64];
                snprintf(buf133z, sizeof(buf133z), "%.9g", (double)fv);
                oss133z << buf133z;
                if (strchr(buf133z, '.') == nullptr && strchr(buf133z, 'e') == nullptr
                    && strchr(buf133z, 'E') == nullptr)
                    oss133z << ".0";
                oss133z << "f";
                return oss133z.str();
            }
            case LiteralKind::Double: {
                // Fix 133z: parser 无 Single 字面量kind, `1!`/`0!` 以 Double 存储,
                // rawText 带 VB 后缀 ("1!"). 原样返回会写进生成 C → C2059
                // (czUI.ctl: `Optional ByVal penWidth As Single = 1!` →
                // `if (!_has_penWidth) penWidth = 1!;`). 剥掉 `!`/`#` 尾缀并
                // 保证是合法 C 浮点字面量 (整数形态补 `.0`).
                std::string dt = lit->rawText;
                while (!dt.empty() && (dt.back() == '!' || dt.back() == '#'))
                    dt.pop_back();
                if (dt.find('.') == std::string::npos
                    && dt.find('e') == std::string::npos
                    && dt.find('E') == std::string::npos)
                    dt += ".0";
                return dt;
            }
            case LiteralKind::String:
                // VB6 "hello" -> C vb6_BSTR_FromStr(L"hello")
                {
                    std::string sInner = lit->rawText;
                    if (sInner.size() >= 2 && sInner.front() == '"' && sInner.back() == '"')
                        sInner = sInner.substr(1, sInner.size() - 2);
                    return "vb6_BSTR_FromStr(L\"" + sInner + "\")";
                }
            case LiteralKind::Boolean:
                // VB6 True = -1, False = 0
                return (lit->rawText == "True" || lit->rawText == "-1") ? "-1" : "0";
            case LiteralKind::Nothing:
                return "NULL";
            case LiteralKind::Empty:
                return "vb6_VariantEmpty()";
            case LiteralKind::Null:
                return "vb6_VariantNull()";
            default:
                break;
        }
    }
    
    // UnaryExpr: 递归求值操作数, 加前缀 (P20-20)
    auto* unary = dynamic_cast<UnaryExpr*>(defaultValue);
    if (unary) {
        std::string inner = evalOptionalDefault(unary->operand.get(), paramType);
        if (inner.empty()) return "";
        switch (unary->op) {
            case UnaryOp::Negate:
                // 数值型: "(-1)" "(-3.14)"
                if (inner.find("vb6_") == 0) return "";  // 非数值C表达式, 暂不处理
                return "(-" + inner + ")";
            case UnaryOp::Not:
                // Not表达式: 暂不常见做默认值, 返回空
                return "";
        }
    }

    // IdentifierExpr: 解析VB6内建常量 (P20-20)
    auto* ident = dynamic_cast<IdentifierExpr*>(defaultValue);
    if (ident) {
        const std::string& n = ident->name;
        // 转小写比较
        std::string nLower = n;
        for (auto& c : nLower) c = (char)tolower((unsigned char)c);

        // 字符串常量
        if (nLower == "vbcrlf" || nLower == "vbnewline")
            return "vb6_BSTR_FromStr(L\"\\r\\n\")";
        if (nLower == "vbcr")
            return "vb6_BSTR_FromStr(L\"\\r\")";
        if (nLower == "vblf")
            return "vb6_BSTR_FromStr(L\"\\n\")";
        if (nLower == "vbtab")
            return "vb6_BSTR_FromStr(L\"\\t\")";
        if (nLower == "vbnullstring")
            return "vb6_BSTR_FromStr(L\"\")";
        if (nLower == "vbback")
            return "vb6_BSTR_FromStr(L\"\\b\")";
        if (nLower == "vbformfeed")
            return "vb6_BSTR_FromStr(L\"\\f\")";
        if (nLower == "vbverticaltab")
            return "vb6_BSTR_FromStr(L\"\\v\")";

        // 数值/枚举常量
        if (nLower == "vbtrue")  return "-1";
        if (nLower == "vbfalse") return "0";
        if (nLower == "vbyes")   return "6";
        if (nLower == "vbno")    return "7";
        if (nLower == "vbok")    return "1";
        if (nLower == "vbcancel") return "2";
        if (nLower == "vbabort")  return "3";
        if (nLower == "vbretry")  return "4";
        if (nLower == "vbignore") return "5";

        // 特殊值
        if (nLower == "vbempty")    return "vb6_VariantEmpty()";
        if (nLower == "vbnull")     return "vb6_VariantNull()";
        if (nLower == "vbnothing")  return "NULL";

        // 项目级Const/EnumMember: 通过符号表查找Constant或EnumMember符号 (P20-20)
        // Fix 081a: 也查找 EnumMember，使 Optional ByVal Ecl As QRCodegenEcc = QRCodegenEcc_LOW
        // 等枚举常量默认值能正确解析为整数值
        {
            Symbol* sym = symTab_.lookup(n);
            if (sym && sym->hasConstValue) {
                if (sym->kind == SymbolKind::Constant || sym->kind == SymbolKind::EnumMember) {
                    switch (sym->constType) {
                    case Vb6Type::Long:
                    case Vb6Type::Integer: { return std::to_string(sym->constIntValue); }
                    case Vb6Type::Single:
                    case Vb6Type::Double: { return std::to_string(sym->constFloatValue); }
                    case Vb6Type::String: {
                        // C-escape constStringValue before embedding in C string literal
                        std::string cEsc;
                        cEsc.reserve(sym->constStringValue.size() + 16);
                        for (char ec : sym->constStringValue) {
                            switch (ec) {
                                case '\\': cEsc += "\\\\"; break;
                                case '"':  cEsc += "\\\""; break;
                                case '\n': cEsc += "\\n"; break;
                                case '\r': cEsc += "\\r"; break;
                                case '\t': cEsc += "\\t"; break;
                                default:   cEsc += ec; break;
                            }
                        }
                        return std::string("vb6_BSTR_FromStr(L\"") + cEsc + "\")";
                    }
                    case Vb6Type::Boolean: { return sym->constBoolValue ? "-1" : "0"; }
                    default: break;
                }
            }
        }
        }

        // 未知标识符, 暂不处理
        return "";
    }
    
    // 其他非字面量表达式 -> 暂不支持, 返回空让cgen用类型零值
    return "";
}


// 虚方法 (tB, B08b): 本类体内调这个名字**必须**走虚槽 (链上有更深的类 Overrides 了它)。
// 表由 2.8 的 runVirtualContractChecks 算好; 空表 = 全工程没有 Overrides → 本函数对所有名字
// 都返回 false, 也就是对存量代码零影响。
bool SemanticAnalyzer::virtualCallNeedsDispatch(const std::string& name) const {
    if (!clsreg_ || clsreg_->empty() || !currentModule_) return false;
    if (!currentModule_->isClassModule) return false;
    auto self = clsreg_->find(ifaceLower(currentModule_->moduleName));
    if (self == clsreg_->end() || self->second.mod != currentModule_) return false;
    const ClassChainView& v = self->second;
    if (v.dynamicKeys.empty()) return false;
    const std::string lk = ifaceLower(name);
    if (lk.empty()) return false;
    for (const auto& k : v.dynamicKeys) {
        if (k == lk) return true;
    }
    return false;
}

// 类继承 (tB, B07b): 祖先自己声明的成员名判定。刻意不复用 stage 3.4 回填的 inhProcs/inhFields
// —— 本判定发生在语义分析途中, 那时合并还没跑; 而且"基类写了什么"这件事 2.8 就已经定死了。
bool SemanticAnalyzer::declaredByAncestor(const std::string& name) const {
    if (!clsreg_ || clsreg_->empty() || !currentModule_) return false;
    if (!currentModule_->isClassModule) return false;
    auto self = clsreg_->find(ifaceLower(currentModule_->moduleName));
    if (self == clsreg_->end() || self->second.mod != currentModule_) return false;
    const ClassChainView& v = self->second;
    if (v.chainBroken || v.chain.size() < 2) return false;
    const std::string lk = ifaceLower(name);
    if (lk.empty()) return false;
    for (size_t i = 0; i + 1 < v.chain.size(); i++) {  // 根 → 父 (末位是自身, 不含)
        auto it = clsreg_->find(v.chain[i]);
        if (it == clsreg_->end() || !it->second.mod) continue;
        for (const auto& d : it->second.mod->declarations) {
            if (!d) continue;
            std::string dn;
            switch (d->kind) {
                case ASTNodeKind::VariableDecl: dn = static_cast<const VariableDecl&>(*d).name; break;
                case ASTNodeKind::SubDecl:      dn = static_cast<const SubDecl&>(*d).name; break;
                case ASTNodeKind::FunctionDecl: dn = static_cast<const FunctionDecl&>(*d).name; break;
                case ASTNodeKind::PropertyDecl: dn = static_cast<const PropertyDecl&>(*d).name; break;
                default: continue;
            }
            if (ifaceLower(dn) == lk) return true;
        }
    }
    return false;
}

// ============================================================
// Protected 可见性 (tB, B08c)
// ============================================================

namespace {

// 一个类**自己声明**的、名字小写键为 lk 的成员里是否含 Protected。
// 返回值: -1 = 本类没声明过这个名字; 0 = 声明了但都不是 Protected; 1 = 有任一 Protected。
// 四类成员 (字段/Sub/Function/Property) —— 刻意与 driver 侧 `memberAccess()` 的口径一致
// (2.8 的"这工程有没有 Protected"早退判据用的就是它)，两边认同样的成员才不会出现
// "表没建 → 判定静默失效"的缝。Event 不在其中：带 Event 的基类在 2.8 已被判死。
// 属性也按**名字**取严而不是按方向: MemberAccessExpr 上拿不到 Get/Let/Set 方向 (与 B08b 的
// dynamicKeys 同一个限制), 同名多方向访问级别不同这种形状在 v1 里按 Protected 论 ——
// 手册页已写明这条边界。
// 为什么读 AST 而不读符号表的 memberAccessLevels: 那张表只在本模块自己的符号表里, 而
// 跨模块的 Class 符号要到 stage 3.5 才注入 (本判定在 stage 3); 登记表里的 mod 指针是
// stage 2.8 就定死的全工程视图, 与 stage 3.4 的合并顺序无关。
int ownProtectedAccess(const Module& m, const std::string& lk) {
    int found = -1;
    for (const auto& d : m.declarations) {
        if (!d) continue;
        std::string dn;
        AccessLevel acc = AccessLevel::Public;
        switch (d->kind) {
            case ASTNodeKind::VariableDecl: {
                const auto& x = static_cast<const VariableDecl&>(*d); dn = x.name; acc = x.access; break; }
            case ASTNodeKind::SubDecl: {
                const auto& x = static_cast<const SubDecl&>(*d); dn = x.name; acc = x.access; break; }
            case ASTNodeKind::FunctionDecl: {
                const auto& x = static_cast<const FunctionDecl&>(*d); dn = x.name; acc = x.access; break; }
            case ASTNodeKind::PropertyDecl: {
                const auto& x = static_cast<const PropertyDecl&>(*d); dn = x.name; acc = x.access; break; }
            default: continue;
        }
        if (dn.empty() || ifaceLower(dn) != lk) continue;
        if (found < 0) found = 0;
        if (acc == AccessLevel::Protected) return 1;
    }
    return found;
}

} // namespace

const ClassChainView* SemanticAnalyzer::selfClassView() const {
    if (!clsreg_ || clsreg_->empty() || !currentModule_) return nullptr;
    if (!currentModule_->isClassModule) return nullptr;
    auto it = clsreg_->find(ifaceLower(currentModule_->moduleName));
    if (it == clsreg_->end() || it->second.mod != currentModule_) return nullptr;
    return &it->second;
}

bool SemanticAnalyzer::checkProtectedVisibility(const Expr& obj, const std::string& member,
                                                const SourceLocation& loc) {
    if (pass_ != 2 || !clsreg_ || clsreg_->empty() || member.empty()) return false;
    const std::string lk = ifaceLower(member);

    // --- 1) 接收者 -> 工程类 (读 srcTypeName: 模块级字段/局部变量/参数三处都留了原文类型名) ---
    std::string recvKey;
    if (obj.kind == ASTNodeKind::IdentifierExpr) {
        const auto& id = static_cast<const IdentifierExpr&>(obj);
        Symbol* s = symTab_.lookup(id.name);
        if (!s || (s->kind != SymbolKind::Variable && s->kind != SymbolKind::Parameter)) return false;
        if (s->srcTypeName.empty()) return false;
        recvKey = ifaceLower(s->srcTypeName);
    } else if (obj.kind == ASTNodeKind::MeExpr) {
        const ClassChainView* self = selfClassView();
        if (!self) return false;
        recvKey = ifaceLower(self->name);
    } else {
        // v1 边界: 属性返回对象 (pvSocket.Pick())、With 块内的 .X、函数实参位置的
        // 调用链等形状都不判定 —— 认不出接收者就放过, 不在这里猜类型。
        return false;
    }
    auto rv = clsreg_->find(recvKey);
    if (rv == clsreg_->end() || !rv->second.mod || rv->second.chainBroken) return false;
    const ClassChainView& recv = rv->second;

    // --- 2) 沿链找最近声明者 (叶优先 = 与 stage 3.4 的遮蔽裁决同向) ---
    std::string declKey;
    std::string declName;
    for (size_t i = recv.chain.size(); i-- > 0;) {
        auto it = clsreg_->find(recv.chain[i]);
        if (it == clsreg_->end() || !it->second.mod) continue;
        const int p = ownProtectedAccess(*it->second.mod, lk);
        if (p < 0) continue;             // 这一层没声明, 继续往根走
        if (p == 0) return false;        // 最近者抢到了键, 且它不是 Protected
        declKey = recv.chain[i];
        declName = it->second.name;
        break;
    }
    if (declKey.empty()) return false;   // 链上无人声明这个名字 (交别的诊断管)

    // --- 3) 当前模块在不在这条家族链上 ---
    const ClassChainView* cur = selfClassView();
    if (!cur) {
        // 类模块但没登记 (泛型模板 / 接口宿主): 证明不了它不在家族里, 放过。
        // 标准模块 / 窗体 / 用户控件不可能是任何类的家族成员, 直接落到下面的越权分支。
        if (currentModule_ && currentModule_->isClassModule) return false;
    } else {
        for (const auto& k : cur->chain) {
            if (k == declKey) return false;   // 声明者就是本类或本类的祖先
        }
    }
    diag_.error(DiagnosticID::SemProtectedOutsideFamily, loc,
        "Member '" + member + "' of class '" + declName + "' is Protected: reachable only from"
        " inside that class family (the declaring class and its derivatives), current context '" +
        (currentModule_ ? currentModule_->moduleName : std::string()) + "' (ai/022 B08c)");
    return true;
}

bool SemanticAnalyzer::namesProjectLevel(const std::string& name) const {
    if (name.empty()) return false;
    const std::string lk = ifaceLower(name);
    if (projPublicProcs_.count(lk)) return true;
    return memberObjCtx_ && projModuleNames_.count(lk);
}

} // namespace vb6c3
