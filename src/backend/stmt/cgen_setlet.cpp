#include "backend/cgen.hpp"
#include <algorithm>
#include <cctype>
#include <iostream>
#include <functional>

namespace vb6c3 {

// --- cgen_setlet.cpp: Set/Let 语句与 Mid$ 语句生成 ---
// 2026-09-17 拆分：原 809 行（visit(SetStmt&) 556 行 + visit(LetStmt&) 213 行 + visit(MidStmt&) 24 行）。
//   visit(SetStmt&) 的函数体（554 行）拆为两个「函数体片段」，在本文件函数体内 #include：
//     detail/cgen_setlet_set_prop.inc  —— 目标侧分派：Nothing / 属性 / COM / 接口引用（原 12~321 行）
//     detail/cgen_setlet_set_rhs.inc   —— 值侧语义：Variant 容器与包装 / WithEvents 事件连接（原 322~565 行）
//   visit(LetStmt&) 与 visit(MidStmt&) 函数体不足 500 行，原样留在本文件，未搬移。
// .inc 是「函数体片段」，在 visit(SetStmt&) 函数体内被 #include（C++ 允许），故不用 .cpp/.hpp 后缀 ——
// 它们不是独立编译单元，单独 include 会编译不过。片段局部变量原样不动，逐行未改 → 零行为改动。

// 账 #258 (§B87): Set 的右值永远是「对象引用上下文」—— VB6 在这里交出**控件本身**，
// 默认属性只在值上下文里展开（`s = Text1` 取 .Text 才对，`Set o = Text1` 要的是控件）。
// 改前这一问有两份答案：环境闸 suppressDefaultProp_（With 块 / As Object 形参那两档设它）
// 与 P16 事后在发好的文本里截 "vb6_hwnd_"（只认 knownWithEventsCtrlVars_ 的目标）——
// 于是泛对象槽 (`Dim o As Object`) 与 Variant 槽拿到的是一枚 BSTR / 一个 int。
// 闸只在这一枚标识符上开：右值是表达式（`Set o = f(Text1)`）时子树照旧按值上下文，不顺着漏下去。
void CCodeGen::emitSetObjectRhs(const ExprPtr& value) {
    const bool bareIdent = value && value->kind == ASTNodeKind::IdentifierExpr;
    const bool savedSuppress = suppressDefaultProp_;
    if (bareIdent) suppressDefaultProp_ = true;
    emitExpr(*value);
    suppressDefaultProp_ = savedSuppress;
}

void CCodeGen::visit(SetStmt& node) {
#include "backend/detail/stmt/cgen_setlet_set_prop.inc"
#include "backend/detail/stmt/cgen_setlet_set_rhs.inc"
}

void CCodeGen::visit(LetStmt& node) {
    if (!node.target || !node.value) return;

    // P7.5+P7.6: 控件属性写入 (Let语句)
    if (node.target->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maExpr = static_cast<MemberAccessExpr&>(*node.target);
        // P18-C: Printer.CurrentX / Printer.CurrentY 赋值
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& objId = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string objName = objId.name;
            std::transform(objName.begin(), objName.end(), objName.begin(), ::tolower);
            std::string memName = maExpr.memberName;
            std::transform(memName.begin(), memName.end(), memName.begin(), ::tolower);
            if (objName == "printer") {
                emitExpr(*node.value);
                if (memName == "currentx") { c_.emitLine("vb6_Printer_SetCurrentX((int32_t)(" + lastExpr_ + "));"); return; }
                if (memName == "currenty") { c_.emitLine("vb6_Printer_SetCurrentY((int32_t)(" + lastExpr_ + "));"); return; }
            }
            // Fix 161f: Screen.MousePointer 全局写入 (CWaitCursor 等类里用;
            // 早先落进类成员 fallback → me->Screen.MousePointer = ... → C2039)。
            if (objName == "screen") {
                if (memName == "mousepointer") {
                    emitExpr(*node.value);
                    c_.emitLine("vb6_Screen_SetMousePointer((int32_t)(" + lastExpr_ + "));");
                    return;
                }
            }
        }
        // P7.6: 控件数组属性写入 ctrlArr(idx).Property = value
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IndexOrCallExpr) {
            auto& idxExpr = static_cast<IndexOrCallExpr&>(*maExpr.object);
            if (idxExpr.callee && idxExpr.callee->kind == ASTNodeKind::IdentifierExpr) {
                auto& arrId = static_cast<IdentifierExpr&>(*idxExpr.callee);
                std::string arrLower = arrId.name;
                std::transform(arrLower.begin(), arrLower.end(), arrLower.begin(), ::tolower);
                if (knownControlArrays_.count(arrLower)) {
                    auto itCtrl = knownFormControls_.find(arrLower);
                    if (itCtrl != knownFormControls_.end()) {
                        std::string writeFn = getControlPropWriteFn(itCtrl->second, maExpr.memberName);
                        if (!writeFn.empty()) {
                            emitExpr(*node.value);
                            std::string valExpr = std::move(lastExpr_);
                            std::string idxArg;
                            if (!idxExpr.positional.empty()) {
                                emitExpr(*idxExpr.positional[0]);
                                idxArg = std::move(lastExpr_);
                                emitExpr(*node.value);
                                // Fix 142: 同 cgen_assign_stmt_special — 展开 RHS 的
                                // COM 属性读 marker (NewTab1(c).Theme).
                                if (isComMarker_) resolveComValue();
                                valExpr = std::move(lastExpr_);
                            } else {
                                idxArg = "0";
                                if (isComMarker_) resolveComValue();
                                valExpr = std::move(lastExpr_);
                            }
                            c_.emitLine(writeFn + "(" + ctrlArrWriteCalleeArgs(writeFn, arrId.name, idxArg) + ", " + valExpr + ");  /* Let Control Array Property */");
                            return;
                        }
                    }
                }
            }
        }
        // P7.5: 非数组控件属性写入
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& objId = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string objLower = objId.name;
            std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
            // P16: WithEvents控件属性写入 (Let)
            {
                auto itWECtrl = knownWithEventsCtrlVars_.find(objLower);
                if (itWECtrl != knownWithEventsCtrlVars_.end()) {
                    std::string writeFn = getControlPropWriteFn(itWECtrl->second, maExpr.memberName);
                    if (!writeFn.empty()) {
                        auto itOrig = knownWithEventsCtrlOrigNames_.find(objLower);
                        std::string weVarName = (itOrig != knownWithEventsCtrlOrigNames_.end()) ? itOrig->second : objLower;
                        emitExpr(*node.value);
                        std::string valExpr = std::move(lastExpr_);
                        valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);  // Fix 194
                        c_.emitLine(writeFn + "(" + weVarName + ", " + valExpr + ");  /* Let WithEvents ctrl prop */");
                        return;
                    }
                }
            }
        auto itCtrl = knownFormControls_.find(objLower);
        if (itCtrl != knownFormControls_.end()) {
            // Fix 292: 表单上的 **UserControl** 若与内置控件同名, 属性必须走 .ctl
            // 实例, 不能走内置 shim。ComCtls 的 ListView / ToolBar / MonthView / ...
            // 全都撞内置控件的名字, 一旦先试 shim 就永远到不了下面的 Fix 133y 分支。
            // 内置 shim 按 hwnd 找的是**它自己那套原生控件**(vb6forms_listview.c 的
            // lvFind), 对 VB6_UserControlHost 一律返回 NULL ⇒ 属性写入静默丢失。
            // 实测 ComCtlsDemo: `ListView1.Sorted = True` 被发成
            //   vb6_ListView_SetSorted(vb6_hwnd_ListView1, (-1))
            // ⇒ ComCtls 的 `Property Let Sorted` 没跑 ⇒ SortListItems 没调
            // ⇒ LVM_SORTITEMSEX 一次都没发(1000 项停在插入序, VB6 是排好序的)。
            // 12 个属性撞名: AllowColumnReorder/Checkboxes/ColumnWidth/FullRowSelect/
            // GridLines/HideColumnHeaders/LabelEdit/MultiSelect/Sorted/SortKey/
            // SortOrder/View。注意读路径 (cgen_expr_member_form_builtin.inc:289) 本来就
            // 先试 .ctl (Fix 112), 只有这里的**写**路径是反的 —— 只改了写路径。
            const bool isUC292 = knownUserControlCtrlVars_.find(objLower)
                                 != knownUserControlCtrlVars_.end();
            std::string writeFn = isUC292
                ? std::string()
                : getControlPropWriteFn(itCtrl->second, maExpr.memberName);
            if (!writeFn.empty()) {
                emitExpr(*node.value);
                std::string valExpr = std::move(lastExpr_);
                valExpr = controlPropValueExpr(writeFn, valExpr, *node.value);  // Fix 194
                c_.emitLine(writeFn + "(" + makeCtrlHwndArg(objLower, itCtrl->second) + ", " + valExpr + ");  /* Let Control Property */");
                return;
            }
                // Fix 133y: 工程内 UserControl 子控件属性写入
                // (`czLabel3.Caption = "..."` / `ucChartBar1.Value = 30`)。
                // builtin 控件 readFn/writeFn 只在表单代码访问标准属性时生效;
                // 自定义 UserControl 属性 (Caption/Value/ChartStyle...) 落在
                // getControlPropWriteFn default → 空 → 此前 "Unknown control
                // property write" 警告 + struct field access (裸标识符, C2065)。
                // 这里与读取路径 (cgen_expr_member_form_builtin.inc 的
                // resolveClassMemberCall) 对称: 按 .ctl 类查 Property Let 形参表,
                // this 取宿主窗口的实例。
                auto itUC132 = knownUserControlCtrlVars_.find(objLower);
                if (itUC132 != knownUserControlCtrlVars_.end()) {
                    std::vector<ParameterInfo> wp132;
                    if (findClassMemberWriteParams(itUC132->second, maExpr.memberName, false, wp132)
                        && !wp132.empty()) {
                        std::string canon132 = canonicalClassMemberName(itUC132->second, maExpr.memberName);
                        if (canon132.empty()) canon132 = maExpr.memberName;
                        std::string setter132 = "vb6_" + cIdent(itUC132->second) + "_"
                                              + propFnSegment("prop_let_", canon132);
                        std::string hwndArg132 = makeCtrlHwndArg(objLower, itCtrl->second);
                        std::string meExpr132 = "(vb6_cls_" + cIdent(itUC132->second)
                                              + "*)vb6_UC_InstanceOf(" + hwndArg132 + ")";
                        const ParameterInfo& p132 = wp132.back();
                        emitExpr(*node.value);
                        std::string val132 = std::move(lastExpr_);
                        // 形参是 String (ByVal BSTR) 时, 值表达式需是 BSTR。
                        // 其余按 QI (inferExprType) 标量直传。
                        if (p132.type == Vb6Type::String) {
                            val132 = wrapToBSTR(val132, *node.value);
                        } else if (p132.type == Vb6Type::Boolean) {
                            val132 = "(" + val132 + " ? (int16_t)-1 : (int16_t)0)";
                        }
                        if (p132.isByVal) {
                            c_.emitLine(setter132 + "(" + meExpr132 + ", " + val132 + ");  /* Let UserControl Property */");
                        } else {
                            c_.emitLine("{ " + mapType(p132.type) + " _uc132 = " + val132 + "; "
                                      + setter132 + "(" + meExpr132 + ", &_uc132); }");
                        }
                        return;
                    }
                }
            }
        }
    }

    // P17.1: WithMemberExpr作为Let目标
    if (node.target->kind == ASTNodeKind::WithMemberExpr && !withObjectVars_.empty() && !withObjectInfoStack_.empty()) {
        auto& wmExpr = static_cast<WithMemberExpr&>(*node.target);
        const auto& info = withObjectInfoStack_.back();
        const std::string& tempVar = withObjectVars_.back();

        // FormControl + WithEventsCtrl Only (COM/Class handled via AssignmentStmt)
        if (info.kind == WithObjKind::FormControl) {
            std::string writeFn = getControlPropWriteFn(info.ctrlType, wmExpr.memberName);
            if (!writeFn.empty()) {
                emitExpr(*node.value);
                if (info.ctrlType == FrmControlType::Menu) {  // P20-36
                    std::string mnuLower = info.ctrlOrigName;
                    std::transform(mnuLower.begin(), mnuLower.end(), mnuLower.begin(), ::tolower);
                    c_.emitLine(writeFn + "(" + makeCtrlHwndArg(mnuLower, info.ctrlType) + ", " + lastExpr_ + ");  /* Let With menu prop */");
                } else {
                    c_.emitLine(writeFn + "(" + tempVar + ", " + lastExpr_ + ");");
                }
                return;
            }
        } else if (info.kind == WithObjKind::WithEventsCtrl) {
            std::string writeFn = getControlPropWriteFn(info.ctrlType, wmExpr.memberName);
            if (!writeFn.empty()) {
                emitExpr(*node.value);
                c_.emitLine(writeFn + "(" + info.ctrlOrigName + ", " + lastExpr_ + ");");
                return;
            }
        }
    }
    // 左值求值期间关掉虚表派发 (ai/022 B08e-4 站点⑭: Pattern C/D2 要对这段文本做
    // `prop_get_` -> `prop_let_` 的字符串级改写, 派发表达式会被换成表里不存在的槽字段)
    const bool savedVirtLhs = suppressVirtDispatch_;
    suppressVirtDispatch_ = true;
    emitExpr(*node.target);
    suppressVirtDispatch_ = savedVirtLhs;
    std::string target = std::move(lastExpr_);
    emitExpr(*node.value);
    // COM属性值: 根据目标变量类型解封
    if (isComMarker_) {
        std::string unpackHint;
        if (node.target->kind == ASTNodeKind::IdentifierExpr) {
            auto& idExpr = static_cast<IdentifierExpr&>(*node.target);
            std::string lower = idExpr.name;
            std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
            if (knownLongPtrVars_.count(lower)) unpackHint = "LongPtr";
            else if (knownLongVars_.count(lower)) unpackHint = "Long";
            else if (knownDoubleVars_.count(lower)) unpackHint = "Double";
            else if (knownObjectVars_.count(lower)) unpackHint = "Object";
            else if (knownBstrVars_.count(lower)) unpackHint = "BSTR";
            else if (knownVariantVars_.count(lower)) unpackHint = "Variant";
        }
        resolveComValue(unpackHint);
    }
    std::string value = std::move(lastExpr_);

    // Fix 010r-16: Let 语句中, 当 LHS 是非左值的 COM 调用或 Property Get,
    // 重写为 vb6_ComSetProp / vb6_ComSetPropArg / prop_let_ 调用.
    if (tryRewriteCOMLvalue(target, value, node.value.get(), /*isSet=*/false)) {
        return;
    }

    // Fix 038 Group 1/2: Variant 数组元素赋值 (同 AssignmentStmt 路径)
    if (isVariantSlotElemRead(target)) {
        // 账 #305 §B150: 与 AssignmentStmt 那一支同一条改写 (一维/多维都在 variantSlotAssignLine 里)
        std::string slotLine;
        if (variantSlotAssignLine(target, "vb6_VariantFromValue(" + value + ")", slotLine)) {
            c_.emitLine(slotLine + ";  /* Let */");
        } else {
            c_.emitLine(target + " = " + value + ";  /* Let */");
        }
    } else if ([&] {
                   // 账 #305 §B164 形⑤ 站③ (第一百八十四刀): 目标是不是 Variant 槽只问这位主人
                   // (与 AssignmentStmt 同一枚); 命中就交给同一枚 emitVariantSlotAssign ——
                   // 装箱 / 所有权 / 清旧值三件都在那一处, LetStmt 不再自己拼形状。
                   bool letUdtField = false;
                   if (!variantSlotTargetOf(node.target.get(), target, letUdtField)) return false;
                   emitVariantSlotAssign(target, node.value.get(), value, letUdtField);
                   return true;
               }()) {
        return;
    } else {
        // Fix 038b-6/045: 具体类型目标 + Variant 值 → 自动提取 (同 AssignmentStmt 路径)
        // 使用 cExprIsVariant (C 字符串级) + knownVariantVars_ 检测.
        bool valueIsVariant = cExprIsVariant(value);
        if (!valueIsVariant && node.value && node.value->kind == ASTNodeKind::IdentifierExpr) {
            auto& id = static_cast<IdentifierExpr&>(*node.value);
            std::string idLower = id.name;
            std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
            if (knownVariantVars_.count(idLower)) valueIsVariant = true;
        }
        // Fix 089j: RHS 为类 Variant 字段 (me->mParentsColKey 等) 也视为
        // Variant 表达式 — cExprIsVariant 只查函数前缀, me-> 成员不命中,
        // 导致 PropertyGet 里 `ret = me->VarField` 不转 → C2440.
        if (!valueIsVariant && value.compare(0, 4, "me->") == 0) {
            std::string rhsMember = value.substr(4);
            std::transform(rhsMember.begin(), rhsMember.end(), rhsMember.begin(), ::tolower);
            if (classVariantMembers_.count(rhsMember)) valueIsVariant = true;
        }
        if (valueIsVariant) {
            std::string convertedValue = value;
            if (target.find("VB6_SA_AT(BSTR,") != std::string::npos) {
                convertedValue = "vb6_VariantToString(" + value + ")";
            } else if (target.find("VB6_SA_AT(int32_t,") != std::string::npos
                       || target.find("VB6_SA_AT(int16_t,") != std::string::npos
                       || target.find("VB6_SA_AT(uint8_t,") != std::string::npos) {
                convertedValue = "vb6_VariantToLong(" + value + ")";
            } else if (target.find("VB6_SA_AT(double,") != std::string::npos
                       || target.find("VB6_SA_AT(float,") != std::string::npos) {
                convertedValue = "vb6_VariantToDouble(" + value + ")";
            } else {
                std::string checkName = target;
                if (checkName.substr(0, 4) == "me->") checkName = checkName.substr(4);
                if (checkName.size() > 4 && checkName[0] == '(' && checkName[1] == '*'
                    && checkName.back() == ')') {
                    checkName = checkName.substr(2, checkName.size() - 3);
                }
                std::string lower = checkName;
                std::transform(lower.begin(), lower.end(), lower.begin(), ::tolower);
                if (knownLongVars_.count(lower)) {
                    convertedValue = "vb6_VariantToLong(" + value + ")";
                } else if (knownDoubleVars_.count(lower)) {
                    convertedValue = "vb6_VariantToDouble(" + value + ")";
                } else if (knownObjectVars_.count(lower)) {
                    // Fix 045: Object (void*) target + Variant value → extract object
                    convertedValue = "vb6_VariantToObjectVal(" + value + ")";
                } else if (node.target && node.target->kind == ASTNodeKind::IdentifierExpr) {
                    auto& id = static_cast<IdentifierExpr&>(*node.target);
                    std::string idLower = id.name;
                    std::transform(idLower.begin(), idLower.end(), idLower.begin(), ::tolower);
                    if (knownLongVars_.count(idLower)) {
                        convertedValue = "vb6_VariantToLong(" + value + ")";
                    } else if (knownDoubleVars_.count(idLower)) {
                        convertedValue = "vb6_VariantToDouble(" + value + ")";
                    } else if (knownObjectVars_.count(idLower)) {
                        convertedValue = "vb6_VariantToObjectVal(" + value + ")";
                    }
                }
            }
            c_.emitLine(target + " = " + convertedValue + ";  /* Let */");
        } else {
            c_.emitLine(target + " = " + value + ";  /* Let */");
        }
    }
}

void CCodeGen::visit(MidStmt& node) {
    // P18-A: Mid$(var, start, len) = expr → vb6_MidSet(&var, start, len, expr)
    emitExpr(*node.start);
    std::string startVar = std::move(lastExpr_);
    std::string lenVar;
    if (node.hasLength) {
        emitExpr(*node.length);
        lenVar = std::move(lastExpr_);
    } else {
        lenVar = "0";
    }
    emitExpr(*node.value);
    std::string valueVar = std::move(lastExpr_);
    // Fix 084b: Mid$(var, start, len) = VariantExpr (如 vSplit(i) 数组元素返回
    // vb6_VARIANT) → 需 vb6_VariantToString 转 BSTR, 否则 vb6_MidSet 第4参数
    // 类型不匹配触发 C2440.
    if (cExprIsVariant(valueVar)) {
        valueVar = "vb6_VariantToString(" + valueVar + ")";
    }
    // Emit target variable address
    // 左值求值期间关掉虚表派发 (ai/022 B08e-4 站点⑭: Pattern C/D2 要对这段文本做
    // `prop_get_` -> `prop_let_` 的字符串级改写, 派发表达式会被换成表里不存在的槽字段)
    const bool savedVirtLhs = suppressVirtDispatch_;
    suppressVirtDispatch_ = true;
    emitExpr(*node.target);
    suppressVirtDispatch_ = savedVirtLhs;
    std::string targetVar = std::move(lastExpr_);
    c_.emitLine("vb6_MidSet(&" + targetVar + ", " + startVar + ", " + lenVar + ", " + valueVar + ");");
}


} // namespace vb6c3
