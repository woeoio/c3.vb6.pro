// cgen_assign_host_pseudo.inc - visit(AssignmentStmt&) 函数体片段：宿主伪对象 / 内建 Parent 的成员写入
// 由 src/backend/stmt/cgen_assign.cpp 拆出（2026-09-29），在函数体内被 #include，非独立编译单元。
//
// Fix <vbeclipse>（VbEclipse.vbp 定向 C 错误清零）：接住两类在 VbEclipse 里大量出现、
// 但此前没有任何通道的 LHS 形态。二者共同点是**限定符不是工程内的类**：
//
//   ① `UserControl.BackColor() = v` / `Ambient.ForeColor = v`
//      VB6 里"带默认参数/索引的属性"写成就带空括号 —— ucButton.ctl 的
//      `Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)` 体内就是
//      `UserControl.BackColor() = New_BackColor`（源码原文，非笔误）。
//      P6.7 的属性写分派只认 target 是 MemberAccessExpr 的形态（cgen_assign_prop_write.inc:7），
//      空括号形态（IndexOrCallExpr 包裹、零实参）落到通用调用分支 → 发出
//      `vb6_UserControl_BackColor() = New_BackColor;` → C2063 "不是一个函数"
//      （ucButton.c:269 BackColor / :297 Enabled）。RTL 里这些宿主成员是
//      **全局变量**（vb6rtl_userctl.h: extern int32_t vb6_UserControl_BackColor），
//      直接按左值赋值即可。
//
//   ② `Parent.Active = True`
//      `Parent` 是 VB6 内建的 As Object 属性（宿主对象），其成员访问是 IDispatch
//      后期绑定；而 UserControl 的设计器元数据里**没有**宿主类型，
//      `Parent` 既不在 knownClassVars_ 也不在工程类表里 → 落到 class_module.inc
//      优先级3 的"模块名.成员"回退 → `vb6_Parent_Active = (-1);` C2065
//      （ucTabStrip.c:198，源 ucTabStrip.ctl:199）。走
//      `vb6_ComSetProp(vb6_UC_ParentObject(), L"Active", ...)` 后期绑定。
//
// 说明：同名真实局部变量优先（VB6 里局部遮蔽内建名），故先查 knownLocalVars_。

    {
        const MemberAccessExpr* maHp = nullptr;
        if (node.target && node.target->kind == ASTNodeKind::MemberAccessExpr) {
            maHp = static_cast<const MemberAccessExpr*>(node.target.get());
        } else if (node.target && node.target->kind == ASTNodeKind::IndexOrCallExpr) {
            // 空括号形态：`UserControl.BackColor() = v`
            const auto& icHp = static_cast<const IndexOrCallExpr&>(*node.target);
            if (icHp.positional.empty() && icHp.named.empty() && icHp.callee
                && icHp.callee->kind == ASTNodeKind::MemberAccessExpr) {
                maHp = static_cast<const MemberAccessExpr*>(icHp.callee.get());
            }
        }
        if (maHp && maHp->object && maHp->object->kind == ASTNodeKind::IdentifierExpr
            && isDesignerModule_) {
            const std::string objNameHp =
                static_cast<const IdentifierExpr&>(*maHp->object).name;
            const std::string objLowerHp = Symbol::toLower(objNameHp);
            const bool isParentHp = (objLowerHp == "parent");
            const bool isHostPseudoHp =
                (objLowerHp == "usercontrol" || objLowerHp == "ambient"
                 || objLowerHp == "extender" || objLowerHp == "propertypage");
            if ((isParentHp || isHostPseudoHp) && !knownLocalVars_.count(objLowerHp)) {
                emitExpr(*node.value);
                std::string valHp = std::move(lastExpr_);
                std::string packFnHp = comPackExpr(*node.value);
                if (isComMarker_) {
                    // RHS 本身是 COM 属性读 → 按 packer 反推解封类型
                    std::string hintHp = "Variant";
                    if (packFnHp == "vb6_ComPackDouble") hintHp = "Double";
                    else if (packFnHp == "vb6_ComPackInt") hintHp = "Long";
                    else if (packFnHp == "vb6_ComPackBSTR") hintHp = "BSTR";
                    else if (packFnHp == "vb6_ComPackObject") hintHp = "Object";
                    resolveComValue(hintHp);
                    valHp = std::move(lastExpr_);
                }
                if (isParentHp) {
                    // Parent.<prop> = v → 宿主对象上的 IDispatch 后期绑定写
                    c_.emitLine("vb6_ComSetProp(vb6_UC_ParentObject(), L\""
                                + escapeWideCString(maHp->memberName) + "\", "
                                + packFnHp + "(" + valHp + "));  /* Parent.<prop> = v */");
                    return;
                }
                // 宿主伪对象成员 = RTL 全局变量（vb6rtl_userctl.h），直接赋值。
                // 成员名按 RTL 约定规范化（源码 hDc → RTL hDC，见 canonicalHostPseudoMember）。
                const std::string tgtHp = "vb6_" + cIdent(objNameHp) + "_"
                    + cIdent(canonicalHostPseudoMember(objNameHp, maHp->memberName));
                // 数值型宿主成员：Variant RHS 先转 Long（其余是句柄/对象成员，原样直传）。
                static const char* kNumericHostMembers[] = {
                    "backcolor", "forecolor", "enabled", "mousepointer", "oledropmode",
                    "righttoleft", "scalewidth", "scaleheight", "scalemode", "appearance",
                    "borderstyle", "containerhwnd", "autoredraw", "width", "height",
                };
                if (cExprIsVariant(valHp)) {
                    const std::string memLowerHp = Symbol::toLower(maHp->memberName);
                    for (const char* numHp : kNumericHostMembers) {
                        if (memLowerHp == numHp) { valHp = "vb6_VariantToLong(" + valHp + ")"; break; }
                    }
                }
                c_.emitLine(tgtHp + " = " + valHp + ";  /* <host pseudo>.<prop> = v */");
                return;
            }
        }
    }
