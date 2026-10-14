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
            // 账 #278 §B120: 这一问由那张表答 (抄了五份的那一份就在这)。
            const bool isHostPseudoHp = hostPseudoIsObject(objLowerHp);
            if ((isParentHp || isHostPseudoHp) && !knownLocalVars_.count(objLowerHp)) {
                emitExpr(*node.value);
                std::string valHp = std::move(lastExpr_);
                std::string packFnHp = comPackExpr(*node.value);
                // 账 #260: 这一份 hint 表与写侧另两处同源，收进 comMarkerValueForWrite
                // （取值顺序与改前一致: 先摘 lastExpr_、再算 packer、最后才消费标记）。
                // 它原来把「其余 packer」一律按 Variant 解封，其中 Boolean 那一档会发出
                // vb6_ComPackBool(vb6_VARIANT 结构体) —— 现在跟着同一张表走 Long 档
                // (与 resolveComMarkerForPack 对 Boolean 的答案一致)。
                valHp = comMarkerValueForWrite(packFnHp, valHp);
                if (isParentHp) {
                    // Parent.<prop> = v → 宿主对象上的 IDispatch 后期绑定写
                    c_.emitLine("vb6_ComSetProp(vb6_UC_ParentObject(), L\""
                                + escapeWideCString(maHp->memberName) + "\", "
                                + packFnHp + "(" + valHp + "));  /* Parent.<prop> = v */");
                    return;
                }
                // 宿主伪对象成员 = RTL 全局变量（vb6rtl_userctl.h），直接赋值。
                // 成员名按 RTL 约定规范化、对象名取表里那一列，装配只经 hostPseudoRtlSymbol（§B120）。
                const std::string tgtHp = hostPseudoRtlSymbol(objNameHp, maHp->memberName);
                // 账 #159: 要不要拆数值由那张表答 (目标槽是整数才拆; 句柄/对象/字符串
                // 成员按原样直传)。此前这里是又一份成员名清单 kNumericHostMembers。
                if (cExprIsVariant(valHp) && hostPseudoIsNumeric(objLowerHp, maHp->memberName)) {
                    valHp = "vb6_VariantToLong(" + valHp + ")";
                }
                c_.emitLine(tgtHp + " = " + valHp + ";  /* <host pseudo>.<prop> = v */");
                return;
            }
        }
    }
