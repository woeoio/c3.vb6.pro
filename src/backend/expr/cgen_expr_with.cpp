#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>
#include <cstdio>

namespace vb6c3 {

// --- cgen_expr_with.cpp: WithMemberExpr 求值（With 块成员访问） ---

void CCodeGen::visit(WithMemberExpr& node) {
    // P17.1: With块内 .Member — 根据对象类型分发
    if (withObjectVars_.empty() || withObjectInfoStack_.empty()) {
        lastExpr_ = "/* .Member outside With */";
        return;
    }

    const auto& info = withObjectInfoStack_.back();
    const std::string& tempVar = withObjectVars_.back();
    std::string memLower = node.memberName;
    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);

    switch (info.kind) {
    case WithObjKind::FormControl: {
        // Fix 110a: Unknown 类型控件 (工程内 UserControl 实例 / 第三方 ActiveX)
        // 没有 Win32 属性表 → 走 COM 晚绑定 (与 cgen_expr_member 的
        // `itCtrl->second != Menu` 分支一致). 否则会发射 tempVar.member →
        // C2039 ("BackColorOpacity": 不是 "HWND__" 的成员).
        if (info.ctrlType == FrmControlType::Unknown) {
            comObjExpr_ = tempVar;
            comMemberName_ = node.memberName;
            isComMarker_ = true;
            isEarlyBoundCom_ = false;
            earlyBoundSym_ = nullptr;
            lastExpr_ = tempVar + "  /* With COM ." + node.memberName + " */";
            return;
        }
        // Fix <vbeclipse>: With <CommonDialog> 块内的 .ShowOpen / .ShowSave / …
        // 与成员侧 cgen_expr_member 的 CD 那一刀同构：打 com 标记（这里存的是
        // With 临时变量名，消费侧见 cgen_expr_call_callee_withm.inc）。不接的
        // 话落到 warn 兜底 → `tempVar.ShowOpen()` → C2039 "不是 HWND__ 的成员"。
        if (info.ctrlType == FrmControlType::CommonDialog) {
            if (memLower == "showopen" || memLower == "showsave" || memLower == "showcolor"
                || memLower == "showfont" || memLower == "showprinter" || memLower == "showabout") {
                comObjExpr_ = tempVar;
                comMemberName_ = node.memberName;
                isComMarker_ = true;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                lastExpr_ = tempVar + "  /* With CD ." + node.memberName + " */";
                return;
            }
        }
        // .Property → vb6_GetControlXxx(tempVar) or Menu prop
        std::string readFn = getControlPropReadFn(info.ctrlType, node.memberName);
        if (!readFn.empty()) {
            if (info.ctrlType == FrmControlType::Menu) {  // P20-36: Menu uses (hmenu, menuId) args
                std::string mnuLower = info.ctrlOrigName;
                std::transform(mnuLower.begin(), mnuLower.end(), mnuLower.begin(), ::tolower);
                lastExpr_ = readFn + "(" + makeCtrlHwndArg(mnuLower, info.ctrlType) + ")  /* With menu .Property */";
            } else {
                lastExpr_ = readFn + "(" + tempVar + ")  /* With ctrl .Property */";
            }
            return;
        }
        // C29-SL-m（账 #150）: With 块里的控件**零实参方法** —— `With sldA : .ClearSel : End With`。
        // 以前成员名不分属性还是方法，方法也落到下面那条"未知属性"兜底，发成
        // `tempVar.方法名()` —— HWND 是 struct 指针，那是**编译错**（实测构建失败 exit 2，
        // 伴 VB4001），不是运行期静默空转。表与两条码头共用（controlZeroArgMethod），
        // 控件类型直接用 With 入口那一帧记着的 ctrlType，不再回头查名字。
        std::string zaFnW = controlZeroArgMethod(info.ctrlType, memLower);
        if (!zaFnW.empty()) {
            // 两形都要照顾：带括号的 `.SetFocus()` 走到这里时 asCallCallee_=true，
            // 直接交付完整调用文本会被调用点再补一对括号 ⇒ `f(hwnd)()`（实测过，非法 C）。
            // 所以照 ClassInstance 那条 Fix 090s 的协议：交付裸函数名 + pendingChainObj_，
            // 由调用点把 this（就是 With 入口那枚 HWND）拼进实参表。
            if (asCallCallee_) {
                pendingChainObj_ = "(void*)" + tempVar;
                lastExpr_ = zaFnW;
            } else {
                lastExpr_ = zaFnW + "(" + tempVar + ")  /* With ctrl .Method */";
            }
            return;
        }
        diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
            std::string("P17.1: Unknown control property '.'") + node.memberName + "' in With block");
        lastExpr_ = tempVar + "." + cIdent(node.memberName);
        return;
    }
    case WithObjKind::WithEventsCtrl: {
        // .Property → vb6_GetControlXxx(ctrlOrigName)
        std::string readFn = getControlPropReadFn(info.ctrlType, node.memberName);
        if (!readFn.empty()) {
            lastExpr_ = readFn + "(" + info.ctrlOrigName + ")  /* With WE ctrl .Property */";
            return;
        }
        lastExpr_ = info.ctrlOrigName + "." + cIdent(node.memberName);
        return;
    }
    case WithObjKind::COMObject: {
        // Fix 161f-extlist: With 目标是 ListView 槽 (形参/局部) —— `.View` / `.hWnd`
        // / `.Sorted` 等标量属性走控件读表, 不让它落 COM 晚绑定
        // (vb6_ComGetStringProp(_vb6_with_N, L"hWnd") = 拿 HWND 当 IDispatch →
        //  hWnd 读回 Empty ⇒ 下游 SendMessage 全打 0; extlist MListViewEx 实测)。
        // 集合成员 (.ListItems/.ColumnHeaders) 不在这张表, 保持 COM 标记链
        // (下游 cgen_expr_call_com_bind.inc 的 listViewHwndExprOf 会认这枚 temp)。
        if (info.ctrlType == FrmControlType::ListView) {
            std::string readFnLv = getControlPropReadFn(FrmControlType::ListView, node.memberName);
            if (!readFnLv.empty()) {
                lastExpr_ = readFnLv + "((void*)" + tempVar + ")  /* With ListView slot .Property */";
                return;
            }
        }
        // .Property → 设置COM标记，让下游(IndexOrCallExpr/AssignmentStmt)处理
        comObjExpr_ = tempVar;
        comMemberName_ = node.memberName;
        isComMarker_ = true;
        isEarlyBoundCom_ = false;
        earlyBoundSym_ = nullptr;
        lastExpr_ = tempVar + "  /* With COM ." + node.memberName + " */";
        return;
    }
    case WithObjKind::ClassInstance: {
        // .Method/Property → 类方法调用 funcName(tempVar)
        // .DataMember → 尝试查找属性/方法, 找不到则用COM后期绑定
        // Fix 011r-1: 若className已知, 优先用resolveClassMemberCall精确解析该类成员
        // 避免symTab_.lookupModule捡错模块(Pattern A2)
        if (!info.className.empty()) {
            const std::string directFnW = resolveClassMemberCall(info.className, node.memberName);
            // 类虚表 (tB, ai/022 B08e): `With b : b.M() 写成 .M()` 与 `b.M()` 是同一个调用,
            // 接收者就是 With 入口那个类实例 temp (纯读变量名) → 必须按 __cvtbl 派发, 否则
            // `With up As InhBase`（up 持派生实例）会静默绑回基类实现 = B08d 修掉的切片。
            // 属性写方向不改写: 3.4b 不给 Let/Set 建槽 (D33-7), 改派到会拿到 Get 的槽。
            std::string funcName = directFnW;
            if (!directFnW.empty()
                && !suppressVirtDispatch_
                && directFnW.find("_prop_let_") == std::string::npos
                && directFnW.find("_prop_set_") == std::string::npos) {
                const std::string dispatchFnW = virtDispatchCallee(
                    info.className, node.memberName, tempVar, /*mustDispatch=*/true, node.loc);
                if (!dispatchFnW.empty()) funcName = dispatchFnW;
            }
            if (!funcName.empty()) {
                // Fix 044a: When used as standalone expression (!asCallCallee_),
                // pad Optional params (value + _has_ flags). When used as callee
                // in IndexOrCallExpr (asCallCallee_=true), the IndexOrCallExpr
                // will handle padding via Fix 044b.
                if (!asCallCallee_) {
                    std::vector<ParameterInfo> params;
                    bool isBuiltin = false;
                    if (findClassMemberCallParams(info.className, node.memberName,
                                                   params, isBuiltin)
                        && !params.empty() && !isBuiltin) {
                        std::string argList = tempVar;
                        for (size_t i = 0; i < params.size(); i++) {
                            const auto& param = params[i];
                            argList += ", ";
                            std::string defVal;
                            if (param.hasDefaultValue && !param.defaultValueExpr.empty()) {
                                defVal = param.defaultValueExpr;
                            } else {
                                defVal = defaultValue(param.type);
                            }
                            if (param.isByVal) {
                                argList += defVal;
                            } else {
                                std::string cType = mapType(param.type);
                                if (param.type == Vb6Type::Variant
                                    || param.type == Vb6Type::Empty
                                    || param.type == Vb6Type::Null
                                    || param.type == Vb6Type::Object) {
                                    argList += "&(" + cType + "){0}";
                                } else {
                                    argList += "&(" + cType + "){" + defVal + "}";
                                }
                            }
                        }
                        for (size_t i = 0; i < params.size(); i++) {
                            const auto& param = params[i];
                            if (param.isOptional && !param.isParamArray) {
                                argList += ", 0";
                            }
                        }
                        lastExpr_ = funcName + "(" + argList + ")";
                    } else {
                        lastExpr_ = funcName + "(" + tempVar + ")";
                    }
                } else {
                    // Fix 090s: asCallCallee_ (CallStmt 无括号调用 / IndexOrCallExpr
                    // 带括号调用) — 交付 this 给调用点补全用户实参与 Optional padding
                    // (同 MAE Fix 015/088b 的 pendingChainObj_ 协议). 此前生成
                    // funcName(tempVar) 完整调用文本 → CallStmt bare-call 分支因
                    // callExpr 含 '(' 跳过参数补齐 → With 内无括号调用 .Start
                    // (cHttpServer.Start 声明带 4 个 Optional 参) 只传 this → C2198.
                    pendingChainObj_ = "(void*)" + tempVar;
                    lastExpr_ = funcName;
                }
                return;
            }
            // 未找到方法/属性 → 假设是数据字段: tempVar->member
            // (此时tempVar类型为 vb6_cls_<className>*, ->访问正确编译)
            // Fix 092p: 字段名规范化回声明名 (源码大小写变体 → C 结构体实际成员名)
            lastExpr_ = tempVar + "->"
                      + cIdent(canonicalClassFieldName(info.className, node.memberName))
                      + "  /* With class ." + node.memberName + " field */";
            return;
        }

        // className未知 — 使用原symTab查找(可能捡错模块, 但无法避免)
        Symbol* memSym = symTab_.lookupModule(node.memberName);
        if (!memSym) {
            memSym = symTab_.lookup(node.memberName);
        }
        if (memSym) {
            if (memSym->kind == SymbolKind::PropertyGet || memSym->kind == SymbolKind::PropertyLet
                || memSym->kind == SymbolKind::PropertySet || memSym->kind == SymbolKind::Sub
                || memSym->kind == SymbolKind::Function) {
                std::string memberCName = node.memberName;
                if (memSym->kind == SymbolKind::PropertyGet)
                    memberCName = "prop_get_" + node.memberName;
                else if (memSym->kind == SymbolKind::PropertyLet)
                    memberCName = "prop_let_" + node.memberName;
                else if (memSym->kind == SymbolKind::PropertySet)
                    memberCName = "prop_set_" + node.memberName;
                std::string funcName = cProcName(memberCName, memSym->access,
                    memSym->isExternal ? memSym->sourceModule : "");
                lastExpr_ = funcName + "(" + tempVar + ")";
                return;
            }
            // 变量/常量/枚举成员 → 尝试属性查找
        }
        // Fix 010n/023d: 未知成员 → COM后期绑定. 改为设置 COM marker
        // 让下游 (IndexOrCallExpr / AssignmentStmt / BinaryExpr / 外层
        // MemberAccessExpr) 在 resolveComValue 时根据上下文 unpackType 选择
        // 正确的 vb6_ComGet*Prop / vb6_ComCall 函数.
        // 此前直接 emit `vb6_ComGetStringProp(tempVar, L"Member")` 会让外层
        // MemberAccessExpr 的链式 COM 检测2 (visit(MemberAccessExpr) ~line 1479)
        // 无法识别为 COM 对象表达式 — 因为该检测只匹配 vb6_ComCallObject /
        // ComGetObjectProp / ComCall / ComGetProp, 不匹配 ComGetStringProp,
        // 导致 `.X.Y` 被错生成成 `vb6_ComGetStringProp(obj, L"X").Y` (C2224:
        // .Y of BSTR-结构体类型). 与 WithObjKind::COMObject case (line ~3362)
        // 行为一致 — 走 COM marker 通道让 resolveComValue("Object") 在链式
        // 外层自然展开为 vb6_ComGetObjectProp.
        comObjExpr_ = tempVar;
        comMemberName_ = node.memberName;
        isComMarker_ = true;
        isEarlyBoundCom_ = false;
        earlyBoundSym_ = nullptr;
        lastExpr_ = tempVar + "  /* With class ." + node.memberName + " COM dispatch */";
        return;
    }
    case WithObjKind::BuiltinObject: {
        // Fix 010l: .Property on builtin object (Err/App/etc.)
        const std::string& bn = info.ctrlOrigName;  // builtin name (lowercase)
        if (bn == "err") {
            if (memLower == "number")      { lastExpr_ = "vb6_ErrNumber()";      return; }
            if (memLower == "description") { lastExpr_ = "vb6_ErrDescription()"; return; }
            if (memLower == "source")      { lastExpr_ = "vb6_ErrSource()";      return; }
            if (memLower == "lastdllerror") { lastExpr_ = "GetLastError()";       return; }
            if (memLower == "helpfile")    { lastExpr_ = "(BSTR)0";              return; }
            if (memLower == "helpcontext") { lastExpr_ = "0";                    return; }
            if (memLower == "clear")       { lastExpr_ = "vb6_ErrClear()";       return; }
            if (memLower == "raise")       { lastExpr_ = "vb6_ErrRaise";         return; }
        }
        // Fallback: unknown builtin member
        lastExpr_ = "/* With " + bn + "." + node.memberName + " */ 0";
        return;
    }
    case WithObjKind::Unknown:
    default:
        // UDT/fallback: struct.field访问
        // Fix 081j: With块临时变量改为指针，用 -> 访问成员
        // Fix 085: UDT With 块中字段为对象 (Collection/COM/项目类) 时追加标记,
        // 供外层 MemberAccessExpr 消费转 COM/类方法路径.
        // 例: With uCtx: .LocalCertificates.Item(lIdx) → _vb6_with_15->LocalCertificates
        {
            std::string tvLower = Symbol::toLower(tempVar);
            std::string udtCType =
                knownUdtVars_.count(tvLower) ? knownUdtVars_[tvLower] : "";
            // Fix 110t: With 块 UDT 对象的**对象字段** (Collection/COM → void*) 作为
            // 调用 callee 时 (With m_Serie(i) 内 `.CustomColors(j + 1)`), 只追加
            // "/* udt objfield void* */" 注释标记不够 — 外层节点是 IndexOrCallExpr
            // 而非 MemberAccessExpr, 标记无人消费, 直接拼接实参 → 畸形调用
            // `_vb6_with_19->CustomColors((j + 1))` C2064 (ucTreeMaps.c 1864/1911/...).
            // 此处改设 COM marker, 由 IndexOrCallExpr 的后期绑定通道生成
            // vb6_ComCall(field, L"Item", args, argc).
            if (asCallCallee_ && !udtCType.empty()
                && udtFieldObjCType(udtCType, Symbol::toLower(node.memberName)) == "void*") {
                comObjExpr_ = tempVar + "->" + cIdent(node.memberName);
                comMemberName_ = node.memberName;
                isComMarker_ = true;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                lastExpr_ = comObjExpr_;
                return;
            }
            lastExpr_ = appendUdtObjFieldMarker(tempVar, udtCType, node.memberName, "->");
        }
        return;
    }
}

} // namespace vb6c3
