#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include <cstdio>
#include <cstdlib>

namespace vb6c3 {

// --- cgen_util_classcall.cpp: 类成员调用解析 (调用/写入参数匹配 + 返回类型) ---


// ============================================================
// Fix 011r-1: 类实例成员调用解析 (跨模块符号表的精确类查找)
// ============================================================
std::string CCodeGen::resolveClassMemberCall(const std::string& className,
                                               const std::string& memberName) const {
    if (className.empty()) return "";
    if (!symTab_.moduleScope()) return "";

    const std::string memberLower = Symbol::toLower(memberName);
    // 调试钩子 (任务 #7): `Form.hWnd` / `CurrControl.hWnd` 这类 Extender 成员被解析成
    // vb6_VBFlexGrid_hWnd 全局 (C2065×17) —— 打印进入时的类名与最终命中的返回路径,
    // 定位 className 是在哪一步变成 VBFlexGrid 的。用法: C3_DBG159=1
    const bool dbg159 = std::getenv("C3_DBG159") != nullptr
        && (memberLower == "hwnd" || memberLower == "name"
            || memberLower == "enabled" || memberLower == "borderstyle");
    if (dbg159) {
        std::fprintf(stderr, "[DBG159] enter cls=%s member=%s mod=%s isClassMod=%d\n",
                     className.c_str(), memberName.c_str(), moduleName_.c_str(),
                     isClassModule_ ? 1 : 0);
    }
    if (std::getenv("C3_DBG110") && memberLower == "enabled") {
        std::fprintf(stderr, "[DBG110] resolveClassMemberCall cls=%s member=%s mod=%s\n",
                     className.c_str(), memberName.c_str(), moduleName_.c_str());
    }
    // Fix 011r-1b: VB6 case-insensitive — compare class name ignoring case
    // (用户源代码可能写 "cWinsock", 但struct定义用的是文件名大小写 "cWinSock")
    const std::string classNameLower = Symbol::toLower(className);
    // 跟踪规范类名 (取自符号表, 用于C输出大小写一致)
    std::string canonicalClassName;

    // 遍历模块级符号, 收集属于指定类且名字匹配的所有方法/属性符号
    // 优先 Get (读上下文最常见), 其次 Sub/Function, 再 Property Let, 最后 Set
    const Symbol* foundGet   = nullptr;
    const Symbol* foundLet   = nullptr;
    const Symbol* foundSet   = nullptr;
    const Symbol* foundSubFn = nullptr;

    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
        if (sym->lowerName != memberLower) continue;

        // 类匹配判定 (case-insensitive: VB6 case-insensitive)
        bool matches = false;
        if (sym->isExternal) {
            if (Symbol::toLower(sym->sourceModule) == classNameLower) {
                matches = true;
                if (canonicalClassName.empty()) canonicalClassName = sym->sourceModule;
            }
        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classNameLower) {  // Fix 013: moduleName_ = VB_Name
            // 当类模块编译自身时, 同模块类的方法符号 isExternal=false
            matches = true;
            if (canonicalClassName.empty()) canonicalClassName = moduleName_;
        }
        if (!matches) continue;

        switch (sym->kind) {
            case SymbolKind::PropertyGet:  foundGet   = sym.get(); break;
            case SymbolKind::PropertyLet:  foundLet   = sym.get(); break;
            case SymbolKind::PropertySet:  foundSet   = sym.get(); break;
            case SymbolKind::Sub:
            case SymbolKind::Function:     foundSubFn = sym.get(); break;
            default: break;  // Variable / Constant / Class / EnumType... 跳过
        }
    }

    // Fix 093a: 成员名大小写规范化 — 调用点拼写可能与类内声明不同 (VB6 大小写
    // 不敏感), 直接用之会生成大小写不匹配的 C 符号 → LNK2019. 以声明拼写为准.
    const std::string memberCanon = canonicalClassMemberName(
        canonicalClassName.empty() ? className : canonicalClassName, memberName);

    // Fix 089g: 跨模块同名成员 storageKey 抢占 — 多个类有同名 Property Get 时
    // (如 cWebSocketClient / cWebSocketServerClient 都有 State), driver.cpp 的
    // globalPublicSyms 按 storageKey (state$pg) 只保留先分析模块的符号. 目标类的
    // Get 外部符号可能被别类抢占而 Let ($pl) 正常注入 → 上面的循环只找到 Let,
    // 读上下文 (oClient.State = 1) 误发 prop_let_State → void 返回/C2198.
    // 修正: Get 缺失但读到 Let/Set 时, 查 Class 符号的 memberProcKinds — semantic
    // 按读上下文优先级 (Get > Function > Sub > Let > Set) 写入, 可确定该成员在读
    // 上下文的真实形式. 若读形式是 PropertyGet/Function/Sub, 按读形式重定向,
    // 避免生成 prop_let_/prop_set_ 调用.
    if (!foundGet) {
        const Symbol* classSymFix = nullptr;
        std::string classCanonFix;
        for (const auto& [ckey2, csym2] : symTab_.moduleScope()->symbols()) {
            if (csym2->kind != SymbolKind::Class) continue;
            if (csym2->isExternal) {
                if (Symbol::toLower(csym2->sourceModule) == classNameLower) {
                    classSymFix = csym2.get();
                    classCanonFix = csym2->sourceModule;
                    break;
                }
            } else if (isClassModule_ && Symbol::toLower(moduleName_) == classNameLower) {
                classSymFix = csym2.get();
                classCanonFix = moduleName_;
                break;
            }
        }
        if (classSymFix) {
            auto itKindFix = classSymFix->memberProcKinds.find(memberLower);
            if (itKindFix != classSymFix->memberProcKinds.end()) {
                std::string canonFix = classCanonFix.empty() ? canonicalClassName : classCanonFix;
                switch (itKindFix->second) {
                    case ProcKind::PropertyGet:
                        return "vb6_" + cIdent(canonFix) + "_prop_get_" + cIdent(memberCanon);
                    case ProcKind::Function:
                    case ProcKind::Sub:
                        return "vb6_" + cIdent(canonFix) + "_" + cIdent(memberCanon);
                    default:
                        break;  // PropertyLet/PropertySet 为主 — 维持 scope 选择
                }
            }
        }
    }

    // 选择优先级: Get > Sub/Function > Let > Set
    // (本helper用于读上下文; 写上下文由tryRewriteCOMLvalue把Get改写成Let/Set)
    const Symbol* chosen = foundGet;
    if (!chosen) chosen = foundSubFn;
    if (!chosen) chosen = foundLet;
    if (!chosen) chosen = foundSet;

    if (!chosen) {
        // Fix 014: 跨模块类成员符号缺失回退 (Class symbol fallback)
        // 当 cTlsSocket.Create 与其他类的 Public 方法同名时 (cPassword.Create / cAsyncSocket.Create),
        // driver.cpp 的 globalPublicSyms 以 storageKey (lowerName) 去重, 只保留首个注入,
        // 导致消费模块 (cTlsReMaster) 的作用域中外部 "create" 符号 sourceModule 不匹配 className.
        // scope iteration 找不到 sourceModule==className 的方法符号 → chosen 为 null.
        // 回退: 直接查找目标类的 Class 符号 (每个类的 Class 符号 storageKey=lowerName 唯一,
        // 不会被同名方法冲突覆盖), 验证 memberName 是否在该类 memberNames 中,
        // 命中则按 Sub/Function 风格发出 vb6_<className>_<memberName> (无前缀).
        // 限制: Property 变体需要 prop_get_/let_/set_ 前缀, 此回退假设 Sub/Function;
        //       若目标成员确实是 Property, 链接器会失败 — 届时再加 Property 分支.
        const Symbol* classSym = nullptr;
        std::string classCanonName;
        for (const auto& [ckey, csym] : symTab_.moduleScope()->symbols()) {
            if (csym->kind != SymbolKind::Class) continue;
            if (csym->isExternal) {
                if (Symbol::toLower(csym->sourceModule) == classNameLower) {
                    classSym = csym.get();
                    classCanonName = csym->sourceModule;
                    break;
                }
            } else if (isClassModule_ && Symbol::toLower(moduleName_) == classNameLower) {
                // 同模块 (本消费模块本身就是该类)
                classSym = csym.get();
                classCanonName = moduleName_;
                break;
            }
        }
        if (classSym) {
            for (const auto& mn : classSym->memberNames) {
                if (Symbol::toLower(mn) == memberLower) {
                    // Fix 016: 用 Class 符号的 memberProcKinds 直接判断前缀, 替代
                    // Fix 014b 的 "$pg 外部符号启发式" (后者在跨类同名混合场景下误判:
                    // 如 cCryptoHMAC.Mode 是 Function 而 cDelay.Mode 是 PropertyGet,
                    // 6 个类的 Mode 共同以 storageKey="mode" 去重, 消费模块作用域中
                    // "mode$pg" 外部符号被某个类率先占据, 启发式将所有类的 Mode 都
                    // 发出 prop_get_Mode — 对 Function 类生成不存在的函数名 → LNK2019).
                    // memberProcKinds 按 className 索引, 互不影响; 其值由
                    // semantic_analyzer.cpp 在创建 Class 符号时按读上下文优先级
                    // (Get > Function > Sub > Let > Set) 填充.
                    std::string prefix;
                    auto itKind = classSym->memberProcKinds.find(memberLower);
                    if (itKind != classSym->memberProcKinds.end()) {
                        switch (itKind->second) {
                            case ProcKind::PropertyGet: prefix = "prop_get_"; break;
                            case ProcKind::PropertyLet: prefix = "prop_let_"; break;
                            case ProcKind::PropertySet: prefix = "prop_set_"; break;
                            default: break;  // Function / Sub: 无前缀
                        }
                    }
                    // 若 memberProcKinds 缺失 (理论上不应发生, 旧 Class 符号兼容),
                    // 回退到原启发式: 作用域中存在 "<memberLower>$pg" 即判为 PropertyGet.
                    if (prefix.empty() && itKind == classSym->memberProcKinds.end()) {
                        std::string pgKey = memberLower + "$pg";
                        if (symTab_.moduleScope()->symbols().find(pgKey)
                            != symTab_.moduleScope()->symbols().end()) {
                            prefix = "prop_get_";
                        }
                    }
                    if (dbg159) {
                        std::fprintf(stderr,
                            "[DBG159] 014fallback -> vb6_%s_%s%s (memberNames hit)\n",
                            cIdent(classCanonName).c_str(), prefix.c_str(),
                            cIdent(mn).c_str());
                    }
                    return "vb6_" + cIdent(classCanonName) + "_" + prefix + cIdent(mn);                }
            }
        }
        if (dbg159) {
            std::fprintf(stderr, "[DBG159] nomatch cls=%s member=%s -> \"\"\n",
                         className.c_str(), memberName.c_str());
        }
        return "";  // 非方法/属性 → 视为数据字段访问
    }

    std::string prefix;
    switch (chosen->kind) {
        case SymbolKind::PropertyGet: prefix = "prop_get_"; break;
        case SymbolKind::PropertyLet: prefix = "prop_let_"; break;
        case SymbolKind::PropertySet: prefix = "prop_set_"; break;
        default: break;  // Sub/Function: 无前缀
    }

    // 强制使用规范类名 (来自符号表, 与类定义struct名一致),
    // 避免用户源代码大小写差异导致生成的函数名与定义不匹配
    if (dbg159) {
        std::fprintf(stderr, "[DBG159] scopeSym -> vb6_%s_%s%s (kind=%d)\n",
                     cIdent(canonicalClassName).c_str(), prefix.c_str(),
                     cIdent(memberCanon).c_str(), (int)chosen->kind);
    }
    return "vb6_" + cIdent(canonicalClassName) + "_" + prefix + cIdent(memberCanon);
}


// ============================================================
// Fix 033: 类感知方法/属性参数查找 (Phase A + Phase B 回退)
// ============================================================
// IndexOrCallExpr calleeParams 解析专用 — 与 resolveClassMemberCall 相同的
// 类匹配规则, 但返回 ParameterInfo 列表本身 (而非函数名).
// Phase A 失败的根本原因: driver.cpp globalPublicSyms 按 storageKey 去重,
// 跨模块同名方法 (如 N 个类的 Create) 在消费模块中仅保留首个注册者的 external
// 符号 → sourceModule 不匹配 className → Phase A 遍历不到目标类符号.
// 此时 Phase B 通过 Class 符号自身的 memberParams 表 (semantic_analyzer 在
// 类扫描阶段已填好, 按 Get > Function > Sub > Let > Set 优先级存参数) 取回.
bool CCodeGen::findClassMemberCallParams(const std::string& className,
                                         const std::string& memberName,
                                         std::vector<ParameterInfo>& outParams,
                                         bool& outIsBuiltin) const {
    outParams.clear();
    outIsBuiltin = false;
    if (className.empty() || !symTab_.moduleScope()) return false;

    const std::string memberLower = Symbol::toLower(memberName);
    const std::string classLower  = Symbol::toLower(className);

    // ---- Phase A: 与 resolveClassMemberCall 同迭代规则 ----
    // 找出模块作用域中所有与目标 className / memberName 匹配的 Sub/Function/Property
    // 符号, 按 Get > Function > Sub > Let > Set 优先级选择最优先者.
    // Note: 同模块类 (!isExternal) 由 isClassModule_ && moduleName_ 匹配; 跨模块类
    // 由 isExternal && sourceModule 匹配.
    const Symbol* foundGet   = nullptr;
    const Symbol* foundLet   = nullptr;
    const Symbol* foundSet   = nullptr;
    const Symbol* foundSubFn = nullptr;
    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
        if (sym->lowerName != memberLower) continue;

        bool matches = false;
        if (sym->isExternal) {
            if (Symbol::toLower(sym->sourceModule) == classLower) matches = true;
        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classLower) {
            matches = true;
        }
        if (!matches) continue;

        switch (sym->kind) {
            case SymbolKind::PropertyGet:  foundGet   = sym.get(); break;
            case SymbolKind::PropertyLet:  foundLet   = sym.get(); break;
            case SymbolKind::PropertySet:  foundSet   = sym.get(); break;
            case SymbolKind::Sub:
            case SymbolKind::Function:     foundSubFn = sym.get(); break;
            default: break;  // Variable / Constant / Class / EnumType / DeclareSub/DeclareFunc 跳过
        }
    }

    const Symbol* chosen = foundGet;
    if (!chosen) chosen = foundSubFn;
    if (!chosen) chosen = foundLet;
    if (!chosen) chosen = foundSet;

    if (chosen
        && (chosen->kind == SymbolKind::Sub || chosen->kind == SymbolKind::Function
            || chosen->kind == SymbolKind::PropertyGet
            || chosen->kind == SymbolKind::PropertyLet
            || chosen->kind == SymbolKind::PropertySet)) {
        outParams = chosen->params;
        outIsBuiltin = chosen->isBuiltin;
        return true;
    }

    // ---- Phase B: 类符号自身 memberParams 表回退 ----
    // Note: 按 className 找到 Class 符号 (跨模块 external Class 符号已由 driver.cpp
    // 从 producing 模块拷贝 memberParams 表); 在该表上 lookup memberLower.
    // memberProcKinds 表 (优先级一致) 同步检查: 若有则进一步确认该成员在该类存在;
    // memberParams 表可能因 declaration 缺失 params 而存在性小于 memberProcKinds,
    // 保险起见以 memberParams 命中作为唯一成功判据.
    for (const auto& [ckey, csym] : symTab_.moduleScope()->symbols()) {
        if (csym->kind != SymbolKind::Class) continue;
        bool classMatches = false;
        if (csym->isExternal) {
            if (Symbol::toLower(csym->sourceModule) == classLower) classMatches = true;
        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classLower) {
            classMatches = true;
        }
        if (!classMatches) continue;

        auto it = csym->memberParams.find(memberLower);
        if (it != csym->memberParams.end()) {
            outParams = it->second;
            outIsBuiltin = false;  // memberParams 仅记录用户类方法, 不是 RTL builtin
            return true;
        }
        // 同 className 的 Class 符号在消费模块中可能注册了多个外部副本, 但都来自同一
        // producing 模块的 Class 符号 memberParams, 命中任一即可. 未命中时继续遍历
        // 后续同名 Class 符号 (理论上不应出现, 留作防御性).
    }

    return false;
}


// ============================================================
// Fix 091a: 属性写方向 (Property Let/Set) 参数查找 (声明见 cgen.hpp)
// ============================================================
bool CCodeGen::findClassMemberWriteParams(const std::string& className,
                                          const std::string& memberName,
                                          bool isSet,
                                          std::vector<ParameterInfo>& outParams) const {
    outParams.clear();
    if (className.empty() || memberName.empty() || !symTab_.moduleScope()) return false;

    const std::string memberLower = Symbol::toLower(memberName);
    const std::string classLower  = Symbol::toLower(className);
    const SymbolKind wantKind = isSet ? SymbolKind::PropertySet : SymbolKind::PropertyLet;

    // ---- Phase A: 模块作用域中的 PropertyLet/PropertySet 符号 ----
    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
        if (sym->kind != wantKind) continue;
        if (sym->lowerName != memberLower) continue;
        bool matches = false;
        if (sym->isExternal) {
            if (Symbol::toLower(sym->sourceModule) == classLower) matches = true;
        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classLower) {
            matches = true;
        }
        if (!matches) continue;
        outParams = sym->params;
        return true;
    }

    // ---- Phase B: Class 符号自身 memberLetParams/memberSetParams ----
    // Note: 跨模块同名属性 ($pl storageKey) 冲突时 Phase A 找不到本类符号, 此时
    // 用 driver.cpp 从 producing 模块拷贝来的写方向表 (同类内 Let/Set 各自唯一,
    // 不会被 Get 优先级遮蔽).
    for (const auto& [ckey, csym] : symTab_.moduleScope()->symbols()) {
        if (csym->kind != SymbolKind::Class) continue;
        bool classMatches = false;
        if (csym->isExternal) {
            if (Symbol::toLower(csym->sourceModule) == classLower) classMatches = true;
        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classLower) {
            classMatches = true;
        }
        if (!classMatches) continue;
        const auto& tbl = isSet ? csym->memberSetParams : csym->memberLetParams;
        auto it = tbl.find(memberLower);
        if (it != tbl.end()) {
            outParams = it->second;
            return true;
        }
    }

    return false;
}


// ============================================================
// Fix 015: Method chaining 解析辅助
// ============================================================

std::string CCodeGen::canonicalClassName(const std::string& typeName) const {
    if (typeName.empty()) return "";
    if (!symTab_.moduleScope()) return typeName;
    std::string lower = Symbol::toLower(typeName);
    // 在模块作用域符号表中查找 Class 符号 (含本模块与 extern 注入),
    // 返回符号记录的规范名 (clsSym->name 或 sourceModule) —
    // 这与 struct 定义 vb6_cls_<name> 中使用的大小写一致.
    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
        if (sym->kind != SymbolKind::Class) continue;
        if (sym->isExternal) {
            if (Symbol::toLower(sym->sourceModule) == lower) return sym->sourceModule;
        }
        if (Symbol::toLower(sym->name) == lower) return sym->name;
    }
    return typeName;  // 未找到 Class 符号 — 用源码大小写返回 (caller 自行承担)
}


std::string CCodeGen::getClassMethodReturnType(const std::string& className,
                                               const std::string& memberName) const {
    if (!symTab_.moduleScope()) return "";
    const std::string memberLower = Symbol::toLower(memberName);
    const std::string classLower  = Symbol::toLower(className);

    // 遍历模块级符号查找属于 className 的 Function/PropertyGet
    // (与 resolveClassMemberCall 相同的匹配规则)
    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
        if (sym->lowerName != memberLower) continue;
        bool matches = false;
        if (sym->isExternal) {
            if (Symbol::toLower(sym->sourceModule) == classLower) matches = true;
        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classLower) {
            matches = true;
        }
        if (!matches) continue;

        if (sym->kind == SymbolKind::Function || sym->kind == SymbolKind::PropertyGet) {
            // 仅 Function/PropertyGet 有返回值
            // Fix <vbeclipse>: 返回类名的权威字段是 variableTypeName (见 Fix 015:
            // semantic_analyzer 在 Function/PropertyGet 符号上记录返回类名), 判据不应
            // 绑死在 type==Object 上。driver 跨模块注入的跨模块 PropertyGet 符号 type
            // 常记为 Variant (type=12) 而非 Object — 如 Folder.Views As List
            // (views$pg kind=PropertyGet type=Variant variableTypeName="List"),
            // 旧判据整条链断 → `With Folder.Views` 拿不到类, 其内 .IsEmpty 落到
            // cgen_expr_with.cpp 的空 className 兜底并套当前模块前缀 →
            // vb6_ucPerspective_IsEmpty (LNK2019).
            // 收紧口径: variableTypeName 非空, 且 (a) type==Object (原行为) 或
            // (b) 该名在作用域内确实解析为项目 Class 符号 (排除 String/Long/UDT/Enum
            // 等非类命名类型, 链应在此终止) — 与下方 Phase B 同一校验口径.
            if (!sym->variableTypeName.empty()) {
                bool retIsClass = false;
                for (const auto& [rk, rs] : symTab_.moduleScope()->symbols()) {
                    if (rs->kind != SymbolKind::Class) continue;
                    if (Symbol::toLower(rs->name) == Symbol::toLower(sym->variableTypeName)
                        || (rs->isExternal
                            && Symbol::toLower(rs->sourceModule)
                                   == Symbol::toLower(sym->variableTypeName))) {
                        retIsClass = true;
                        break;
                    }
                }
                // Fix <vbeclipse> rev7: 判据**收紧为 retIsClass**, 不再接受
                // `type == Vb6Type::Object` 单独成立。
                //
                // 原因: `Property Get View() As Object` (View.cls:66) 的
                // variableTypeName 就是 "Object" 且 type==Object ⇒ 旧条件成立 ⇒
                // canonicalClassName("Object") 在作用域里查不到 Class 符号, 按源码
                // 大小写**返回伪类名 "Object"**。下游 (inferClassTypeOfExpr
                // MemberAccessExpr 分支) 拿到非空串就当类实例继续链式推 ⇒
                // ucFolder.c 生成
                //     vb6_ComPackBSTR(vb6_View_prop_get_View((*View))->Caption)
                // 而 vb6_View_prop_get_View() 实际返回 void* (As Object) ⇒
                // C2223/C2039 ×62 (ucFolder / ucSplitBar / modSubClass 全线)。
                //
                // 口径与下方 Phase B 的校验**完全一致**: 返回名必须在作用域内真的
                // 解析为 Class 符号。`As Object` 是后期绑定, 链必须终止 → 交回 COM
                // 路径 (下游 vb6_ComGetProp 拿 IDispatch 问 Caption/hWnd), 那才是
                // VB6 的真实语义。type==Object 但 vtn 为空的注入残缺情形本来就
                // 落到 Phase B, 行为不变。
                if (retIsClass) {
                    // Fix 015 semantic_analyzer 已在该 Function 的 variableTypeName 记录返回类名
                    return canonicalClassName(sym->variableTypeName);
                }
            }
            // Fix 088e: 符号匹配但注入的 type 记录不完整 (PropertyGet 返回类实例
            // 如 RecvBuffer As cByteBuffer, 注入符号 type 未标 Object/变量类型名空)
            // 时, 不再提前 return "" — 落入 Phase B 用 Class 符号的 memberReturnTypes
            // (语义层原始记录, 不受注入影响) 兜底. 否则 Client.RecvBuffer.Size 等
            // 链推断断 → 生成 prop_get(...).Size (C2039: Size 不是 vb6_cls_cByteBuffer
            // 的成员).
            continue;
        }
        // Property Let/Set / Sub 无返回值, 跳过
    }

    // Fix 015: storageKey 冲突回退. 当多名 Public 方法在不同类中重名 (如 cCryptoHMAC.DataString
    // 与 cCryptoHash.DataString, cCryptoHMAC.Mode/cCryptoHash.Mode/cDelay.Mode/...), driver.cpp 的
    // globalPublicSyms 按 storageKey(lowerName) 去重, 消费模块作用域中该名字的外部符号
    // 只保留首个注册者的 sourceModule — 与 className 不匹配 → 上面的循环找不到 → 返回 "".
    // 类似 resolveClassMemberCall 的 Fix 014b 回退, 这里从 Class 符号自身的 memberReturnTypes
    // 表(以 className 索引, 不受 storageKey 冲突影响)读取成员返回类型名. 再校验该返回类型
    // 在当前作用域确实存在对应 Class 符号 (排除 UDT/Enum/String 等非类命名类型 — 链应终止).
    const Symbol* classSym = nullptr;
    for (const auto& [ckey, csym] : symTab_.moduleScope()->symbols()) {
        if (csym->kind != SymbolKind::Class) continue;
        if (csym->isExternal) {
            if (Symbol::toLower(csym->sourceModule) == classLower) {
                classSym = csym.get();
                break;
            }
        } else if (isClassModule_ && Symbol::toLower(moduleName_) == classLower) {
            // 类模块编译自身: 同模块的 Class 符号 isExternal=false
            classSym = csym.get();
            break;
        }
    }
    if (classSym) {
        auto it = classSym->memberReturnTypes.find(memberLower);
        if (it != classSym->memberReturnTypes.end()) {
            const std::string& rawRetName = it->second;
            const std::string retLower = Symbol::toLower(rawRetName);
            // 校验 rawRetName 对应当前作用域中真实存在的 Class 符号 (canonical 同时取大小写)
            for (const auto& [k2, s2] : symTab_.moduleScope()->symbols()) {
                if (s2->kind != SymbolKind::Class) continue;
                if (s2->isExternal) {
                    if (Symbol::toLower(s2->sourceModule) == retLower) {
                        return s2->sourceModule;
                    }
                } else if (Symbol::toLower(s2->name) == retLower) {
                    return s2->name;
                }
            }
            // 返回类型不是类(可能是 UDT/Enum/String/接口等) → 链终止, 不应当继续链式调用
            return "";
        }
    }
    return "";
}


// Fix <vbeclipse> rev8: classMemberReturnsAsObject —— className 的 memberName 成员
// 返回类型是不是 `As Object` (C 侧 void*, 语义 = IDispatch)。
//
// 判据两层, 与 getClassMethodReturnType 的收紧口径 (rev7) 同源:
//  ① 类符号的 memberReturnTypes 里有该成员, 且归一化后是 Object/Variant/Empty/Null
//  ② 该成员的 PropertyGet/Function 符号 type==Object 且 variableTypeName 解析不出
//     Class 符号 (As Object 的 variableTypeName 就是 "Object")
// 两层都要求"**不是**工程类" —— 真的返回工程类时走 vb6_cls_X* 原生路径。
bool CCodeGen::classMemberReturnsAsObject(const std::string& className,
                                          const std::string& memberName) const {
    if (className.empty() || memberName.empty()) return false;
    if (!symTab_.moduleScope()) return false;
    const std::string memberLower = Symbol::toLower(memberName);
    const std::string classLower  = Symbol::toLower(className);
    auto isClassName = [this](const std::string& n) {
        for (const auto& [k, s] : symTab_.moduleScope()->symbols()) {
            if (s->kind != SymbolKind::Class) continue;
            if (Symbol::toLower(s->name) == Symbol::toLower(n)
                || (s->isExternal && Symbol::toLower(s->sourceModule) == Symbol::toLower(n)))
                return true;
        }
        return false;
    };
    auto isObjectish = [](const std::string& n) {
        const std::string l = Symbol::toLower(n);
        return l == "object" || l == "variant" || l == "empty" || l == "null";
    };
    // ① 类符号的 memberReturnTypes
    for (const auto& [ck, cs] : symTab_.moduleScope()->symbols()) {
        if (cs->kind != SymbolKind::Class) continue;
        const bool clsHit = cs->isExternal
            ? (Symbol::toLower(cs->sourceModule) == classLower)
            : (Symbol::toLower(cs->name) == classLower);
        if (!clsHit) continue;
        auto it = cs->memberReturnTypes.find(memberLower);
        if (it != cs->memberReturnTypes.end() && isObjectish(it->second)) return true;
        break;
    }
    // ② 成员符号本体
    for (const auto& [k, s] : symTab_.moduleScope()->symbols()) {
        if (s->lowerName != memberLower) continue;
        const bool clsHit = s->isExternal
            ? (Symbol::toLower(s->sourceModule) == classLower)
            : (isClassModule_ && Symbol::toLower(moduleName_) == classLower);
        if (!clsHit) continue;
        if (s->kind != SymbolKind::PropertyGet && s->kind != SymbolKind::Function) continue;
        if (s->type != Vb6Type::Object) continue;
        if (!s->variableTypeName.empty() && !isObjectish(s->variableTypeName)
            && isClassName(s->variableTypeName))
            return false;               // 真的返回工程类 → 走原生路径
        return true;
    }
    return false;
}
// 账 #191: 工程内 UserControl 的**实例成员访问** —— Fix 112 那段从发码点抽到这里，
// 让「单枚控件」与「控件数组的元素」两条路共用一个出口（以前只有单枚那条，
// `ucChartBar1(i).AddSerie ...` 就整条掉进 COM 兜底 ⇒ 空值）。
// recvExpr = 这一枚实例的宿主句柄表达式（单枚是 `vb6_hwnd_<名>`，元素是
// `vb6_CtrlArr_GetAt(&vb6_arr_<名>, <idx>)`）—— 两条路只差这一个串。
// 命中并把 lastExpr_ / pendingChainObj_ 填好时返回 true；成员不在这个 .ctl 上返回 false
// （调用方据此继续走它自己的兜底）。
bool CCodeGen::emitUcInstanceMemberExpr(const std::string& ucClass,
                                        const std::string& member,
                                        const std::string& recvExpr) {
    if (ucClass.empty()) return false;
    std::string resolvedFn = resolveClassMemberCall(ucClass, member);
    if (resolvedFn.empty()) return false;

    std::string thisArg = "(vb6_cls_" + cIdent(ucClass) + "*)vb6_UC_InstanceOf(" + recvExpr + ")";
    if (asCallCallee_) {
        pendingChainObj_ = thisArg;
        lastExpr_ = resolvedFn;
        return true;
    }
    // 值上下文无括号调用: 按形参表补 Optional 默认值 (同 Fix 089h)
    std::vector<ParameterInfo> paramsUC;
    bool isBuiltinUC = false;
    if (resolvedFn.find("_prop_") == std::string::npos
        && findClassMemberCallParams(ucClass, member, paramsUC, isBuiltinUC)
        && !paramsUC.empty() && !isBuiltinUC) {
        std::string argListUC = thisArg;
        for (size_t i = 0; i < paramsUC.size(); i++) {
            const auto& pm = paramsUC[i];
            argListUC += ", ";
            std::string defV = (pm.hasDefaultValue && !pm.defaultValueExpr.empty())
                             ? pm.defaultValueExpr : defaultValue(pm.type);
            if (pm.isByVal) argListUC += defV;
            else {
                std::string ct = mapType(pm.type);
                if (pm.type == Vb6Type::Variant || pm.type == Vb6Type::Empty
                    || pm.type == Vb6Type::Null || pm.type == Vb6Type::Object)
                    argListUC += "&(" + ct + "){0}";
                else argListUC += "&(" + ct + "){" + defV + "}";
            }
        }
        for (size_t i = 0; i < paramsUC.size(); i++) {
            if (paramsUC[i].isOptional && !paramsUC[i].isParamArray) argListUC += ", 0";
        }
        lastExpr_ = resolvedFn + "(" + argListUC + ")";
    } else {
        lastExpr_ = resolvedFn + "(" + thisArg + ")";
    }
    return true;
}

} // namespace vb6c3
