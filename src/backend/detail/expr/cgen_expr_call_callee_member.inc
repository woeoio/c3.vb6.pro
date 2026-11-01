// cgen_expr_call_callee_member.inc - visit(IndexOrCallExpr&) 函数体片段：Fix 037 MemberAccessExpr callee（Pattern A / B / C / D）
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 806~1009 行，逐行未改。

    // 账 (Outlook2013): **自类型过程/局部 COM 变量作链基的多级调用** —
    //   同窗体:  `RibbonBar.RedrawBar` / `RibbonBar.Controls.Add(...)` (frmMain:1421)
    //   跨窗体:  `frmMain.RibbonBar.FindTab(12000)` (frmShortcutBar:83)
    //   PropertyGet: `CalendarControl.Options.SetScale2TimeZone tzi` (frmTimeZone:216)
    // 基 = 过程名与返回的 typelib 类同名 (符号表同键, 类型符号被遮蔽 —— 必须问
    // lookupModuleByKind, 全表 lookup 先命中 ComClass) 或**本窗体无同名局部变量的**
    // 自类型过程, 或局部 COM 对象变量。中间成员逐层 vb6_ComGetObjectProp,
    // 尾方法 vb6_ComCallObject。此前落 Fix 037 fallback 拼成结构体文本 (C2224)。
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && node.named.empty()) {
        // 收集链: chain[0]=最外层(方法), ... chain[n-1]=贴着基的成员
        std::vector<const MemberAccessExpr*> chainSt;
        const Expr* curSt = node.callee.get();
        while (curSt && curSt->kind == ASTNodeKind::MemberAccessExpr) {
            auto& maC = static_cast<const MemberAccessExpr&>(*curSt);
            chainSt.push_back(&maC);
            curSt = maC.object.get();
        }
        std::string recvSt;
        bool haveSt0 = false;
        int stArmSt = 0;
        if (curSt && curSt->kind == ASTNodeKind::IdentifierExpr && !chainSt.empty()) {
            auto& oidSt0 = static_cast<const IdentifierExpr&>(*curSt);
            std::string oLowerSt0 = Symbol::toLower(oidSt0.name);
            bool otherSt0 = knownClassVars_.count(oLowerSt0)
                || knownUdtVars_.count(oLowerSt0) || knownIfaceVars_.count(oLowerSt0)
                || knownNewVars_.count(oLowerSt0);
            if (!otherSt0) {
                // (A) 同模块自类型过程 (无同名局部变量/对象变量/控件时)
                if (!knownLocalVars_.count(oLowerSt0)
                    && !knownObjectVars_.count(oLowerSt0)
                    && !knownTypedComVars_.count(oLowerSt0)
                    && knownFormControls_.count(oLowerSt0) == 0
                    && !(currentProc_ && Symbol::toLower(currentProc_->name) == oLowerSt0)) {
                    Symbol* stSym0 = symTab_.lookupModuleByKind(oidSt0.name, SymbolKind::Function);
                    if (!stSym0) stSym0 = symTab_.lookupModuleByKind(oidSt0.name, SymbolKind::PropertyGet);
                    if (stSym0
                        && !stSym0->variableTypeName.empty()
                        && Symbol::toLower(stSym0->variableTypeName) == oLowerSt0) {
                        std::string stName0 = stSym0->name;
                        if (stSym0->kind == SymbolKind::PropertyGet)
                            stName0 = "prop_get_" + stName0;
                        std::string stMod0 = stSym0->isExternal ? stSym0->sourceModule : moduleName_;
                        recvSt = cProcName(stName0, stSym0->access, stMod0)
                               + (isClassModule_ ? "((void*)me)" : "()");
                        haveSt0 = true;
                        stArmSt = 1;
                    }
                }
                // (B) 跨窗体: 基是已知外部窗体名。**控件优先** —— 控件名与 typelib
                // 类同名时 (frmMain.CommandBars: 控件 XtremeCommandBars.CommandBars
                // vs 库里同名的 coclass) 必须发控件句柄访问器, 发过程调用 = LNK2019
                // (该窗体根本没有那个过程)。控件不是再问自类型过程/限定键过程。
                if (!haveSt0 && chainSt.size() >= 2
                    && oLowerSt0 != Symbol::toLower(moduleName_)
                    && knownFormModuleNames_.count(oLowerSt0)) {
                    const std::string& midName0 = chainSt[chainSt.size() - 1]->memberName;
                    std::string midLower0 = Symbol::toLower(midName0);
                    const FrmControl* extCtrlB = findExternalFormControl(oidSt0.name, midName0);
                    if (extCtrlB) {
                        std::string accB = "vb6_form_ctrl_" + cIdent(oidSt0.name)
                                         + "_" + cIdent(midName0);
                        if (extCtrlB->controlType == FrmControlType::TreeView)
                            externalTreeViewAccs_[accB] = accB;
                        // 控件数组: 访问器带 (int32_t idx), 下标由外层 IndexOrCallExpr 补。
                        recvSt = (extCtrlB->index >= 0) ? accB : (accB + "()");
                        haveSt0 = true;
                        stArmSt = 2;
                    } else {
                        bool otherMid0 = knownLocalVars_.count(midLower0)
                            || knownFormControls_.count(midLower0);
                        Symbol* midCls0 = symTab_.lookup(midName0);
                        Symbol* midQk0 = symTab_.moduleScope()
                            ? symTab_.moduleScope()->lookupLocal(oLowerSt0 + "_" + midLower0)
                            : nullptr;
                        if (!otherMid0
                            && ((midCls0 && midCls0->kind == SymbolKind::ComClass
                                 && Symbol::toLower(midCls0->name) == midLower0)
                                || (midQk0 && midQk0->isExternal
                                    && midQk0->kind == SymbolKind::Function))) {
                            recvSt = cProcName(midName0, AccessLevel::Public, oidSt0.name) + "()";
                            haveSt0 = true;
                            stArmSt = 2;
                        }
                    }
                }
                // (C) 局部 COM 对象变量 (Dim rb As <typelib 类> 类型解析失败成 void*)
                // **只接多级链** (chainSt.size() >= 2)。单级 `d.Exists("K")` 不归这一支管:
                // 它落到本文件末尾那条通用晚绑定出口, 发的是交回 VARIANT 的 `vb6_ComCall`,
                // 值语境/条件语境才能接 `vb6_ComResultToBool` / `vb6_VariantTo*`。
                // 这一支只交对象出口 (`vb6_ComCallObject`) —— 单级也走它的话, 条件里的
                // `If d.Exists(k)` 就成了 `if (指针)` 恒真 (test_late_cond / test_com_* 实测)。
                if (!haveSt0 && chainSt.size() >= 2 && knownLocalVars_.count(oLowerSt0)
                    && (knownObjectVars_.count(oLowerSt0)
                        || knownTypedComVars_.count(oLowerSt0)
                        || knownVariantVars_.count(oLowerSt0))) {
                    // 账 (Outlook2013 / b216_slot_arm_answers): **Variant 槽不是指针**。
                    // C 侧 v1 是 `vb6_VARIANT` 结构体, 直接当接收者交出去就是
                    // `vb6_ComCallObject(v1, L"Twice", …)` ⇒ C2172 (实参不是指针) ×2。
                    // 与 comObjectRefFromCallExpr 的兜底同一句口径: 从 Variant 取对象指针
                    // 只走 vb6_VariantToObjectVal 这一位主人。
                    recvSt = knownVariantVars_.count(oLowerSt0)
                           ? ("vb6_VariantToObjectVal(" + cIdent(oidSt0.name) + ")")
                           : cIdent(oidSt0.name);
                    haveSt0 = true;
                    stArmSt = 1;
                }
            }
        }
        if (haveSt0) {
            // 中间成员逐层 GetObjectProp。链序: chain[0]=最外层(方法) ...
            // chain[n-1]=贴基成员。(A)/(C): 贴基那位也是普通成员 → 从 n-1 起;
            // (B): 贴基那位已被过程调用替代 → 从 n-2 起。
            size_t midFromSt = (stArmSt == 2) ? chainSt.size() - 2 : chainSt.size() - 1;
            std::string objSt = recvSt;
            for (size_t i = midFromSt; i >= 1; i--) {
                objSt = "vb6_ComGetObjectProp(" + objSt + ", L\""
                      + escapeWideCString(chainSt[i]->memberName) + "\")";
            }
            std::vector<std::string> packedSt0;
            for (size_t i = 0; i < node.positional.size(); i++) {
                std::string packFnSt0 = comPackExpr(*node.positional[i]);
                emitExpr(*node.positional[i]);
                { std::string rSt0 = resolveComMarkerForPack(packFnSt0);
                  if (!rSt0.empty()) lastExpr_ = rSt0; }
                packedSt0.push_back(packArgValue(packFnSt0, lastExpr_));
            }
            std::string argsArrSt0 = "(void*[]){";
            for (size_t i = 0; i < packedSt0.size(); i++) {
                if (i > 0) argsArrSt0 += ", ";
                argsArrSt0 += packedSt0[i];
            }
            argsArrSt0 += "}";
            lastExpr_ = "vb6_ComCallObject(" + objSt + ", L\""
                      + escapeWideCString(chainSt[0]->memberName) + "\", "
                      + (packedSt0.empty() ? std::string("NULL") : argsArrSt0) + ", "
                      + std::to_string(packedSt0.size()) + ")";
            return;
        }
    }

    // Fix 245c: `<返回 Object/Variant 的无参 Function>.方法(args)` —— 后期绑定调用。
    // ComCtlsDemo Pager.ctl:1249 `PropBuddyControl.Move PropBuddyControl.Left,
    // PropBuddyControl.Top, …` (Pager.ctl:1203 `Private Function PropBuddyControl()
    // As Object`)。callee 的对象部分既不是变量也不是工程内模块名 ⇒ 唯一合理的语义是
    // 「(调用 Fn) 得到的 IDispatch 上做方法调用」→ 直接发 vb6_ComCall。
    // **必须排在下面的 Fix 037 之前**: 否则 callee 被当"成员访问"取成属性值
    // (`.Move` 读成 L"Move" 属性), 外层实参再被当默认成员索引, 生成
    // vb6_ComCall(<VARIANT>, L"Item", …) → C2172 (vb6_ComCall 第一参数不是指针)
    // ×2 + C2440 vb6_VARIANT→double ×2 (Pager.c)。
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && node.named.empty()) {
        auto& ma245c = static_cast<MemberAccessExpr&>(*node.callee);
        if (ma245c.object && ma245c.object->kind == ASTNodeKind::IdentifierExpr) {
            auto& oid245c = static_cast<IdentifierExpr&>(*ma245c.object);
            const std::string oLower245c = Symbol::toLower(oid245c.name);
            const bool otherKind245c =
                knownLocalVars_.count(oLower245c) || knownClassVars_.count(oLower245c)
                || knownObjectVars_.count(oLower245c) || knownTypedComVars_.count(oLower245c)
                || knownUdtVars_.count(oLower245c) || knownIfaceVars_.count(oLower245c)
                || knownNewVars_.count(oLower245c) || knownFormControls_.count(oLower245c);
            if (!otherKind245c) {
                Symbol* fn245c = symTab_.lookup(oid245c.name);
                if (!fn245c) fn245c = symTab_.lookupModule(oid245c.name);
                if (fn245c && fn245c->kind == SymbolKind::Function
                    && (Symbol::toLower(fn245c->variableTypeName) == "object"
                        || Symbol::toLower(fn245c->variableTypeName) == "variant")
                    && !(currentProc_ && Symbol::toLower(currentProc_->name) == oLower245c)) {
                    std::string fnMod245c = fn245c->isExternal ? fn245c->sourceModule
                                          : (isClassModule_ ? moduleName_ : std::string());
                    if (!fnMod245c.empty()) {
                        std::string fnCall245c = cProcName(fn245c->name, fn245c->access, fnMod245c)
                                               + "((void*)me)";
                        std::vector<std::string> packed245c;
                        for (size_t i = 0; i < node.positional.size(); i++) {
                            std::string packFn245c = comPackExpr(*node.positional[i]);
                            emitExpr(*node.positional[i]);
                            { std::string r245c = resolveComMarkerForPack(packFn245c);
                              if (!r245c.empty()) lastExpr_ = r245c; }
                            // Fix 266: 打包器按 AST 推的型 (ComPackDouble), 而 marker 解析交回
                            // 的可能是**无型** VARIANT ⇒ 交 packArgValue 一处收口
                            // (Pager.ctl:1249 `PropBuddyControl.Move .Left, .Top` 的两个
                            //  vb6_ComPackDouble(vb6_VARIANT) 就是这里发出的)。
                            packed245c.push_back(packArgValue(packFn245c, lastExpr_));
                        }
                        int32_t argc245c = (int32_t)packed245c.size();
                        std::string argsArr245c = "(void*[]){";
                        for (int i = 0; i < argc245c; i++) {
                            if (i > 0) argsArr245c += ", ";
                            argsArr245c += packed245c[i];
                        }
                        argsArr245c += "}";
                        lastExpr_ = "vb6_VariantFromComResult(vb6_ComCall(" + fnCall245c
                                  + ", L\"" + escapeWideCString(ma245c.memberName) + "\", "
                                  + argsArr245c + ", " + std::to_string(argc245c) + "))";
                        return;
                    }
                }
            }
        }
    }

    // ---- Fix 037: MemberAccessExpr callee with positional args — array field/COM dispatch ----
    // 处理 obj.field(idx) 模式, field 可能是:
    //   - UDT 数组字段 (固定/动态)     — Pattern A/B → obj.member[idx] / VB6_SA_AT(...)
    //   - 类 Variant 字段持有 SafeArray — Pattern C → vb6_VariantArrayGet(&obj->member, idx)
    //   - 类 void* 字段持有 COM 对象    — Pattern D → vb6_ComCall(obj->member, L"Item", args, argc)
    // 若 member 是该类的方法/属性, resolveClassMemberCall 返回非空 → 不拦截 (正常函数调用路径).
    // 不拦截则 fallback path (line ~2158) 把 field 误当函数调用 → C2064.

    // ---- Fix 158a: MemberAccessExpr callee 带 **空括号** — UDT 动态/固定数组成员
    // 整体引用 obj.member() → 裸数组表达式 obj.member (不带调用括号).
    // 此前 fallback 把它当函数调用 emit 成 obj.member() → C2064 (项不会计算为接受
    // N 参数的函数), 或作为 ByRef 数组实参被包裹成 &(obj.member()) (VBFlexGrid.ctl
    // SortGrid: Call MergeSortRec(..., VBFlexGridCells.Rows(), Sort) — ByRef
    // Data() As TCOLS, C 形参 vb6_SafeArray1D**). 与 Pattern A/B 相同判定
    // (inferUdtTypeOfExpr + udtMembers + isArrayDynamic / arraySize>0), 但空括号
    // 不消费 positional. 固定数组成员 (Pattern A) 空括号引用同样走这里:
    // C 侧固定数组字段是静态数组 obj.member[], 整体引用即裸名 obj.member.
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && node.positional.empty() && node.named.empty()) {
        auto& maExpr158 = static_cast<MemberAccessExpr&>(*node.callee);
        std::string memLower158 = Symbol::toLower(maExpr158.memberName);
        if (maExpr158.object) {
            std::string udtCType158 = inferUdtTypeOfExpr(*maExpr158.object);
            const std::string prefix158 = "vb6_type_";
            if (udtCType158.size() > prefix158.size()
                && udtCType158.compare(0, prefix158.size(), prefix158) == 0) {
                std::string udtName158 = udtCType158.substr(prefix158.size());
                Symbol* udtSym158 = symTab_.lookupModule(udtName158);
                if (udtSym158 && udtSym158->kind == SymbolKind::UserDefinedType) {
                    for (auto& mi158 : udtSym158->udtMembers) {
                        if (Symbol::toLower(mi158.name) == memLower158) {
                            if (mi158.isArrayDynamic || mi158.arraySize > 0) {
                                emitExpr(*node.callee);
                                return;
                            }
                            break;
                        }
                    }
                }
            }
        }
    }

    // ============================================================
    // P20-44: OLE 拖放的方法调用 —— 两种形状 (独立于下面的 positional 大门,
    // 因为 X.OLEDrag 通常**无实参**, `!node.positional.empty()` 挡不住它):
    //   ① DataObject 形参的方法: Data.SetData(v) / Data.Files(i) / Data.Clear()
    //   ② 控件的拖动方法:        X.OLEDrag
    // 都不能掉 COM 派发 (本机 hwnd 不是 IDispatch)。
    // ============================================================
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maDD = static_cast<MemberAccessExpr&>(*node.callee);
        if (maDD.object && maDD.object->kind == ASTNodeKind::IdentifierExpr
            && node.named.empty() && node.positional.size() <= 1) {
            auto& objIdDD = static_cast<IdentifierExpr&>(*maDD.object);
            std::string objLowerDD = Symbol::toLower(objIdDD.name);
            std::string memLowerDD = Symbol::toLower(maDD.memberName);
            // ① DataObject 形参方法
            if (dataObjectParams_.count(objLowerDD)) {
                std::string ddArg = "(void*)(*" + objIdDD.name + ")";
                if (memLowerDD == "setdata" && node.positional.size() == 1) {
                    emitExpr(*node.positional[0]);
                    lastExpr_ = "{ vb6_oleDD_SetText(" + ddArg + ", "
                              + std::move(lastExpr_) + "); (void)0; }";
                    return;
                }
                if (memLowerDD == "clear" && node.positional.empty()) {
                    lastExpr_ = "{ vb6_oleDD_Clear(" + ddArg + "); (void)0; }";
                    return;
                }
                if (memLowerDD == "files" && node.positional.size() == 1) {
                    emitExpr(*node.positional[0]);
                    lastExpr_ = "vb6_oleDD_GetFileBstr(" + ddArg + ", "
                              + std::move(lastExpr_) + ")  /* DataObject.Files(i) */";
                    return;
                }
            }
            // ② 控件的 OLEDrag 方法 → 造 DataObject(控件文本) + 源事件注册 + DoDragDrop
            auto itCtlDD = knownFormControls_.find(objLowerDD);
            if (itCtlDD != knownFormControls_.end() && memLowerDD == "oledrag") {
                std::string hwndDD = makeCtrlHwndArg(objLowerDD, itCtlDD->second);
                static const char* ddSrcEv[4] = { "_OLEStartDrag", "_OLESetData",
                                                  "_OLEGiveFeedback", "_OLECompleteDrag" };
                std::string regDD;
                for (int ddK = 0; ddK < 4; ddK++) {
                    auto* ddSym = symTab_.lookup(objIdDD.name + ddSrcEv[ddK]);
                    if (!ddSym) continue;
                    std::string ddFn = cProcName(ddSym->name, AccessLevel::Private);
                    regDD += "  { extern void " + ddFn + "(); vb6_OLEDrop_SetHandler("
                           + hwndDD + ", " + std::to_string(ddK + 2) + ", (void*)" + ddFn + "); }" + "\n";
                }
                std::string allowedDD = "3";   // 默认 vbCFCopy|vbCFMove
                if (node.positional.size() == 1) {
                    emitExpr(*node.positional[0]);
                    allowedDD = std::move(lastExpr_);
                }
                lastExpr_ = (std::string)"{ void* ddObj44 = (void*)vb6_oleDD_MakeDataObject("
                          "vb6_GetControlText(" + hwndDD + "), NULL, 0);" + "\n"
                          + regDD
                          + "  vb6_OLEDrag_Start((void*)" + hwndDD + ", ddObj44, "
                          + allowedDD + "); }";
                return;
            }
        }
    }

    // ============================================================
    // Fix 162a: X.Move l[, t[, w[, h]]] —— VB6 语言级方法 (Form/控件通用)。
    // 之前落 vb6_ComGetObjectProp(hwnd, L"Move") → 在 HWND 上取属性必失败,
    // Form_Load/Form_Resize 的布局全部静默失效 (extlist 实测: ListView 停在
    // 设计宽 5835 twips 只 5 列可见; 基准 VB6 里被 Move 拉到全宽 10 列全见)。
    // 生成 vb6_ControlMove(hwnd, L,T,W,H, mask); mask 位 1=L 2=T 4=W 8=H,
    // RTL 侧缺失位保持当前值, 单位 twips 与 Left/Top/Width/Height 一致。
    // ============================================================
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr) {
        auto& maMV = static_cast<MemberAccessExpr&>(*node.callee);
        if (Symbol::toLower(maMV.memberName) == "move" && node.named.empty()
            && node.positional.size() <= 4 && maMV.object) {
            std::string hwndMV;
            if (maMV.object->kind == ASTNodeKind::IdentifierExpr) {
                auto& objIdMV = static_cast<IdentifierExpr&>(*maMV.object);
                std::string objLowerMV = Symbol::toLower(objIdMV.name);
                auto itCtlMV = knownFormControls_.find(objLowerMV);
                if (itCtlMV != knownFormControls_.end()) {
                    hwndMV = makeCtrlHwndArg(objLowerMV, itCtlMV->second);
                } else if (listViewSlotVars_.count(objLowerMV)) {
                    hwndMV = objLowerMV;   // Fix 161f: 槽变量本身就是 HWND
                } else if (objLowerMV == "me") {   // 保险: Me 被归为 IdentifierExpr 的解析器
                    auto itOrigMV = knownFormControlOriginalNames_.find(knownFormName_);
                    hwndMV = "vb6_hwnd_" + (itOrigMV != knownFormControlOriginalNames_.end()
                                            ? cIdent(itOrigMV->second) : moduleName_);
                }
            } else if (maMV.object->kind == ASTNodeKind::MeExpr
                       && isFormModule_ && !knownFormName_.empty()) {
                auto itOrigMV = knownFormControlOriginalNames_.find(knownFormName_);
                hwndMV = "vb6_hwnd_" + (itOrigMV != knownFormControlOriginalNames_.end()
                                        ? cIdent(itOrigMV->second) : moduleName_);
            }
            if (!hwndMV.empty()) {
                std::string aMV[4] = {"0", "0", "0", "0"};
                // Fix 261: 实参是**值上下文**, 必须清掉 asCallCallee_ 再发射。
                // 本段是从 callee 解析路径进来的 —— 语句级 `X.Move a,b,c,d` 由
                // visit(CallStmt) 把 asCallCallee_ 置 true 后才 emitExpr(callee),
                // 落到这里时它仍是 true。于是实参里的成员访问 (如 `ToolBar1.Height`)
                // 被 emitUcInstanceMemberExpr / 类成员路径当成"要接实参的 callee",
                // 交付**裸函数名** + pendingChainObj_, 而实参位没人消费 pendingChainObj_
                // ⇒ 裸名泄漏成实参 (接收者丢失)。
                // 实测 ComCtlsDemo: RichTextBoxForm.frm:165
                //   `RichTextBox1.Move 0, ToolBar1.Height, .ScaleWidth, .ScaleHeight - ToolBar1.Height`
                // 生成 `vb6_ControlMove(vb6_hwnd_RichTextBox1, 0, vb6_ToolBar_prop_get_Height, ...)`
                // ⇒ C2440×2 (float(__cdecl*)(vb6_cls_ToolBar*) / void* → double) +
                //    C2113 (指针减法) + C2198 (vb6_ControlMove 实参太少)。共 4 条。
                bool savedCcMV = asCallCallee_;
                asCallCallee_ = false;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    // Fix 261b: 实参里的 COM 成员读要**就地消费标记** —— 这里原来只
                    // 搬走 lastExpr_, 标记留在原地, 于是 `With Me` 块里的 `.ScaleWidth`
                    // 交出的是 With 入口那枚裸 HWND (void*) ⇒
                    // C2440 "void* → double" (vb6_ControlMove 的四个 double 槽)。
                    // vb6_ControlMove 清一色 double ⇒ 按 Double 档取值。
                    // 实测 ComCtlsDemo RichTextBoxForm.frm:165
                    //   `RichTextBox1.Move 0, ToolBar1.Height, .ScaleWidth, .ScaleHeight - ToolBar1.Height`
                    if (isComMarker_) {
                        lastExpr_ = resolveComValue("Double");
                    } else {
                        // 账 (Outlook2013 frmInbox:224/227): Move 实参是**未声明类型的
                        // 隐式变量** (nTop = IIf(...)) — C 侧 vb6_VARIANT 槽裸传
                        // vb6_ControlMove 的 double 形参 (C2440 ×2)。裸标识符
                        // (无括号/无成员链) + Variant 档 → 折 Double。
                        const std::string& aMVarg = lastExpr_;
                        bool bareIdentMV = !aMVarg.empty()
                            && aMVarg.find('(') == std::string::npos
                            && aMVarg.find('-') == std::string::npos
                            && aMVarg.find('.') == std::string::npos
                            && aMVarg.find('>') == std::string::npos;
                        if (bareIdentMV
                            && inferExprType(*node.positional[i]) == Vb6Type::Variant) {
                            lastExpr_ = "vb6_VariantToDouble(" + aMVarg + ")";
                        }
                    }
                    aMV[i] = std::move(lastExpr_);
                }
                asCallCallee_ = savedCcMV;
                int maskMV = (1 << (int)node.positional.size()) - 1;
                lastExpr_ = "{ vb6_ControlMove(" + hwndMV + ", " + aMV[0] + ", "
                          + aMV[1] + ", " + aMV[2] + ", " + aMV[3] + ", "
                          + std::to_string(maskMV) + "); (void)0; }";
                return;
            }
        }
    }

    // Fix <vbeclipse>: UserControl 的隐式 `Parent` 伪对象 (VB6 声明 As Object) 上的
    // 成员调用必须走**后期绑定**, 不能按"模块名.成员"做全工程按名查找 ——
    //   ucFolder.ctl:338  Parent.EventRaise "RemoveFolder", FolderId, ViewId
    //   ucTab.ctl:298    Parent.EventRaise "DblClick", Me.ViewId
    //   ucSplitBar.ctl:138  Parent.Refresh          (无括号语句式)
    //   ucTabStrip.ctl:199  Parent.Active = True    (属性写 → cgen_assign_host_pseudo.inc)
    // `Parent` 不是已知变量 → 落到 class_module.inc 优先级3, 用 lookupModule 命中
    // 工程内**别的**类的同名方法 (ucTabStrip.EventRaise), 发出裸函数名
    // vb6_ucTabStrip_EventRaise: ① 丢接收者 (原型首参 me) ② 形参表按被误选的 2 参
    // 签名裁剪, 3 实参被 Fix 041c 截成 2 个 → C2198 ×18; 无括号形态则发
    // vb6_Parent_Refresh (C2065, ucSplitBar.c:205).
    // VB6 里 Parent As Object 的成员调用本就是 IDispatch 后期绑定, 这里同构发
    // vb6_ComCall(vb6_UC_ParentObject(), L"<成员>", 实参包, argc).
    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && node.named.empty() && isDesignerModule_) {
        auto& maPar = static_cast<MemberAccessExpr&>(*node.callee);
        if (maPar.object && maPar.object->kind == ASTNodeKind::IdentifierExpr
            && Symbol::toLower(static_cast<IdentifierExpr&>(*maPar.object).name) == "parent") {
            std::vector<std::string> packedPar;
            for (size_t i = 0; i < node.positional.size(); i++) {
                std::string packFnPar = comPackExpr(*node.positional[i]);
                emitExpr(*node.positional[i]);
                { std::string resolvedPar = resolveComMarkerForPack(packFnPar);
                  if (!resolvedPar.empty()) lastExpr_ = resolvedPar; }
                packedPar.push_back(packArgValue(packFnPar, lastExpr_));
            }
            int32_t argcPar = (int32_t)packedPar.size();
            std::string argsArrayPar = "NULL";
            if (argcPar > 0) {
                argsArrayPar = "(void*[]){";
                for (int i = 0; i < argcPar; i++) {
                    if (i > 0) argsArrayPar += ", ";
                    argsArrayPar += packedPar[i];
                }
                argsArrayPar += "}";
            }
            lastExpr_ = "vb6_ComCall(vb6_UC_ParentObject(), L\""
                      + escapeWideCString(maPar.memberName) + "\", "
                      + argsArrayPar + ", " + std::to_string(argcPar) + ")";
            return;
        }
    }

    if (node.callee && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && !node.positional.empty() && node.named.empty()) {
        auto& maExpr = static_cast<MemberAccessExpr&>(*node.callee);
        std::string memLower = Symbol::toLower(maExpr.memberName);
        std::string memLowerM = "m_" + memLower;
        bool handled = false;

        // ---- P20-42: SSTab 的带下标索引属性 TabCaption(i) / TabVisible(i) ----
        // 与 StatusBar 的 `Panels(i).Text` 不是一回事: 那个是"集合对象再取成员属性",
        // 这里 TabCaption 是**控件自身的参数化属性**, VB6 写作 SSTab1.TabCaption(0)。
        // 不拦就会掉到下面的 COM 派发, 生成 vb6_ComCall(vb6_hwnd_SSTab1, L"TabCaption", …)
        // —— 可本机 hwnd 不是 IDispatch, 运行期静默答错或直接崩。
        // 写完这一侧还要配 cgen_util_comwrite.cpp 的对应写分支 (Get→Set)。
        if (maExpr.object && maExpr.object->kind == ASTNodeKind::IdentifierExpr
            && node.positional.size() == 1) {
            auto& objIdTs = static_cast<IdentifierExpr&>(*maExpr.object);
            std::string objLowerTs = Symbol::toLower(objIdTs.name);
            auto itCtlTs = knownFormControls_.find(objLowerTs);
            if (itCtlTs != knownFormControls_.end()
                && itCtlTs->second == FrmControlType::SSTab) {
                std::string fnTs;
                if (memLower == "tabcaption")      fnTs = "vb6_SSTab_GetTabCaption";
                else if (memLower == "tabvisible") fnTs = "vb6_SSTab_GetTabVisible";
                else if (memLower == "tabtooltiptext") fnTs = "vb6_SSTab_GetTabToolTipText";
                if (!fnTs.empty()) {
                    emitExpr(*node.positional[0]);
                    std::string idxTs = std::move(lastExpr_);
                    lastExpr_ = fnTs + "(" + makeCtrlHwndArg(objLowerTs, itCtlTs->second)
                              + ", " + idxTs + ")  /* SSTab indexed prop */";
                    return;
                }
            }
        }

        // ---- Pattern A/B: UDT 数组字段 ----
        // 仅当 maExpr.object 的 UDT 类型可推断 (IdentifierExpr / WithMemberExpr / 嵌套
        // MemberAccessExpr). inferUdtTypeOfExpr 找到 UDT C 类型 → 在 udtMembers 中查 member.
        if (maExpr.object) {
            std::string udtCType = inferUdtTypeOfExpr(*maExpr.object);
            if (!udtCType.empty()) {
                const std::string prefix = "vb6_type_";
                if (udtCType.size() > prefix.size()
                    && udtCType.compare(0, prefix.size(), prefix) == 0) {
                    std::string udtName = udtCType.substr(prefix.size());
                    Symbol* udtSym = symTab_.lookupModule(udtName);
                    if (udtSym && udtSym->kind == SymbolKind::UserDefinedType) {
                        // 账 #476 §B190 (拍板: 路由只消费出口): 这个成员是不是**动态**数组、
                        // 元素是哪一档、是不是 UDT、秩多少 —— 一律先问那一枚出口, 本站不再自己
                        // 摸 udtMembers 答档。Pattern A (定长) 仍然走 mi.arraySize: 出口只认
                        // 动态数组成员, 定长那一形不归它。
                        MemberArrayDecl mad;
                        const bool memIsDynArr =
                            resolveMemberArrayDecl(udtCType, std::string(),
                                                   maExpr.memberName, mad);
                        for (auto& mi : udtSym->udtMembers) {
                            if (Symbol::toLower(mi.name) == memLower) {
                                if (mi.arraySize > 0) {
                                    // Pattern A: 固定大小数组成员 → obj.member[idx]
                                    // 账 #262: 括号层数问 mi.arrayRank —— 多维成员的发码形态是
                                    // C 多维数组 (`float M[5][5]`)，以前只折第一维，于是
                                    // `.M(3, 3)` 与 `.M(4, 4)` 写进同一格。
                                    emitExpr(*node.callee);  // emits "obj.member"
                                    std::string fieldExpr = std::move(lastExpr_);
                                    lastExpr_ = udtFixedMemberIndex(fieldExpr, node.positional,
                                                                    mi.arrayRank);
                                    handled = true;
                                } else if (memIsDynArr && node.positional.size() == 1) {
                                    // Pattern B: 动态数组成员 → VB6_SA_AT(elemCType, obj.member, idx)
                                    // Fix 092s: UDT 元素发 vb6_type_<UDT> (mapSaElemCType 对
                                    // UserDefinedType 会退化成 vb6_VARIANT ⇒ 元素大小/取址全错,
                                    // cTlsSocket 4086/4093) —— 这件事现在由那一枚出口答。
                                    emitExpr(*node.callee);
                                    std::string fieldExpr = std::move(lastExpr_);
                                    emitExpr(*node.positional[0]);
                                    std::string idx = std::move(lastExpr_);
                                    lastExpr_ = "VB6_SA_AT(" + mad.elemCType + ", "
                                              + fieldExpr + ", " + idx + ")";
                                    handled = true;
                                } else if (memIsDynArr && node.positional.empty()) {
                                    // 账 #476 §B190: 整数组读 (`v = t.Items`) —— 让字段指针原样
                                    // 进 vb6_VariantFromValue, _Generic 的 `struct
                                    // vb6_SafeArray1D*` 臂选 vb6_VariantArray ⇒ VT 8204。
                                    // 以前这一形落到底层通用路, 发成 `vb6_VariantLong(t->Items)`。
                                    emitExpr(*node.callee);
                                    std::string fieldExpr = std::move(lastExpr_);
                                    lastExpr_ = "vb6_VariantFromValue(" + fieldExpr + ")";
                                    handled = true;
                                } else if (mi.type == Vb6Type::Variant) {
                                    // Pattern J: UDT Variant 字段 (持有 SafeArray) →
                                    // vb6_VariantArrayGet(&obj.field, idx)
                                    emitExpr(*node.callee);  // emits "obj.field"
                                    std::string fieldExpr = std::move(lastExpr_);
                                                                        // 账 #305 §B150: 秩数 = 实参个数, 这一问只在 variantSlotElemRead 一处答。
                                    lastExpr_ = variantSlotElemRead("&" + fieldExpr, node.positional);
                                    handled = true;
                                } else if (mi.type == Vb6Type::Object) {
                                    // Pattern H: UDT void* (Object) 字段 (持有 COM 对象) →
                                    // vb6_ComCall(obj.field, L"Item", args, argc)
                                    emitExpr(*node.callee);
                                    std::string objExpr = std::move(lastExpr_);
                                    std::vector<std::string> packedArgs;
                                    for (size_t i = 0; i < node.positional.size(); i++) {
                                        std::string packFn = comPackExpr(*node.positional[i]);
                                        emitExpr(*node.positional[i]);
                                        { std::string resolved = resolveComMarkerForPack(packFn);
                                          if (!resolved.empty()) lastExpr_ = resolved; }
                                        packedArgs.push_back(packArgValue(packFn, lastExpr_));
                                    }
                                    int32_t argc = (int32_t)packedArgs.size();
                                    std::string argsArray = "(void*[]){";
                                    for (int i = 0; i < argc; i++) {
                                        if (i > 0) argsArray += ", ";
                                        argsArray += packedArgs[i];
                                    }
                                    argsArray += "}";
                                    lastExpr_ = "vb6_ComCall(" + objExpr + ", L\"Item\", "
                                              + argsArray + ", " + std::to_string(argc) + ")";
                                    handled = true;
                                }
                                break;
                            }
                        }
                    }
                }
            }
        }

        // ---- 账 #305 §B221 形②: 接收者的值已经是接口视图 (MakeDirect().Twice(21)) ----
        // 下面 Pattern C/D 会把接口名当工程类名用 ⇒ 发成字段访问 `X()->Twice(21)`
        // (两台实测 C2039: 视图结构体里只有一枚 vt)。视图的成员调用只有一条正解 =
        // `x->vt-><slot>(x, args)`，而那条发射的主人刚收在 ivrefVtableCall 里；接收者
        // 必须只求值一次 (返回对象的调用内部自带 AddRef) ⇒ 走 Fix 090q 的 addPending。
        if (!handled && maExpr.object
            && !(maExpr.object->kind == ASTNodeKind::IdentifierExpr
                 && knownIvrefVars_.count(Symbol::toLower(
                        static_cast<const IdentifierExpr&>(*maExpr.object).name)))) {
            // 裸接口变量那种接收者 (Dim o As I 之后 o.M(...)) 早由 visit(MemberAccessExpr)
            // 的 B04 臂接住、发成 `o->vt->m(o, ...)`，这一臂刻意不插手 —— 否则每条存量
            // 接口调用都会多出一枚临时量，把「零行为改动」那条护栏打成一片假差异。
            const std::string ivIface221 = ifaceViewOfExpr(*maExpr.object);
            const IfaceView* ivV221 = ivIface221.empty() ? nullptr : ivLookupIface(ivIface221);
            const std::string ivSlot221 =
                ivV221 ? ivSlotKeyForMember(*ivV221, maExpr.memberName) : std::string();
            if (!ivSlot221.empty()) {
                emitExpr(*maExpr.object);
                const std::string ivRecv221 = "__ivrecv" + std::to_string(ivrefSetSeq_++);
                c_.addPending(ivrefCType(ivIface221) + " " + ivRecv221 + " = "
                              + lastExpr_ + ";");
                std::vector<std::string> ivArgs221;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    ivArgs221.push_back(lastExpr_);
                }
                for (auto& named : node.named) {
                    emitExpr(*named.value);
                    ivArgs221.push_back(lastExpr_);
                }
                lastExpr_ = ivrefVtableCall(ivIface221, maExpr.memberName, ivRecv221, ivArgs221);
                handled = true;
            }
        }

        // ---- Pattern C/D: 类实例字段访问 ----
        if (!handled && maExpr.object) {
            std::string className = inferClassTypeOfExpr(*maExpr.object);
            if (getenv("C3_DBG476")) {
                MemberArrayDecl dbg476;
                const bool dbgOk = resolveMemberArrayDecl(std::string(), className,
                                                          maExpr.memberName, dbg476);
                fprintf(stderr, "[DBG476] member=%s className=%s npos=%zu exit=%d ok=%d elem=%s\n",
                        maExpr.memberName.c_str(), className.c_str(), node.positional.size(),
                        (int)dbgOk, (int)dbg476.ok, dbg476.elemCType.c_str());
            }
            if (!className.empty()) {
                // resolveClassMemberCall 验证 member 是否为方法/属性:
                // 是 → 跳过 (正常函数调用 fallback 处理); 否 → 数据字段 → 检查 Variant/void*
                std::string resolvedFn = resolveClassMemberCall(className, maExpr.memberName);
                if (resolvedFn.empty()) {
                    // 账 #476 §B190 (拍板: 路由只消费出口): 类字段是**动态数组成员** ⇒ 元素
                    // 读 / 元素写 / 整数组读一律先问那一枚出口, 档位与形状都由它给。
                    // 这一臂以前**不存在** —— `o.Items(i)` 于是落到底层通用路, 发成
                    // `o->Items(i)`: 那是在调用 `vb6_SafeArray1D*` 这个指针 (cl C2064,
                    // 11 处实测), 而 `v = o.Items` 发成 `vb6_VariantLong(o->Items)` (C2440)。
                    // 只接**一维** (positional.size()==1) 与零参两种形状: 多维动态数组成员
                    // 今天没有实物 (Pattern B 同样只发一维), 多维那形照旧落底层, 不猜。
                    MemberArrayDecl madC;
                    const bool clsIsDynArr =
                        resolveMemberArrayDecl(std::string(), className, maExpr.memberName, madC);
                    if (clsIsDynArr && node.positional.size() == 1) {
                        emitExpr(*node.callee);  // emits "obj->member" (带字段尾注)
                        std::string fieldExpr = std::move(lastExpr_);
                        emitExpr(*node.positional[0]);
                        std::string idx = std::move(lastExpr_);
                        if (madC.isUdt) {
                            // 账 #476 §B190: 元素是 UDT 且这一枚还要被外层 `.Field` 接住
                            // (真语料 `TextureTool.Vertices(i).tU`) —— 必须发**指针形**。
                            // `VB6_SA_AT` 是宏 `(*(type*)vb6_SaElemPtr(...))`, 外层再拼
                            // `->X` 会接在宏调用的括号外 ⇒ cl C2816「宏调用缺少终止 )」。
                            lastExpr_ = "(" + madC.elemCType
                                      + "*)vb6_SaElemPtr(" + fieldExpr + ", (int32_t)("
                                      + idx + "), (int32_t)sizeof(" + madC.elemCType + "))";
                        } else {
                            lastExpr_ = "VB6_SA_AT(" + madC.elemCType + ", "
                                      + fieldExpr + ", " + idx + ")";
                        }
                        handled = true;
                    } else if (clsIsDynArr && node.positional.empty()) {
                        // 整数组读: 字段指针原样进 vb6_VariantFromValue ⇒ _Generic 选
                        // vb6_VariantArray ⇒ VT 8204 (VB6 的形状)。
                        emitExpr(*node.callee);
                        std::string fieldExpr = std::move(lastExpr_);
                        lastExpr_ = "vb6_VariantFromValue(" + fieldExpr + ")";
                        handled = true;
                    } else if (classVariantMembers_.count(memLower)
                        || classVariantMembers_.count(memLowerM)) {
                        emitExpr(*node.callee);  // emits "obj->member  /* class var .X field */"
                        std::string fieldExpr = std::move(lastExpr_);
                                                // 账 #305 §B150: 秩数 = 实参个数, 这一问只在 variantSlotElemRead 一处答。
                        lastExpr_ = variantSlotElemRead("&" + fieldExpr, node.positional);
                        handled = true;
                    } else {
                        // Pattern D: void* COM 字段 → vb6_ComCall(obj->member, L"Item", args, argc)
                        bool isVoidPtr = false;
                        if (classVoidFieldMap_) {
                            auto itV = classVoidFieldMap_->find(className);
                            if (itV != classVoidFieldMap_->end()) {
                                if (itV->second.count(memLower)
                                    || itV->second.count(memLowerM)) {
                                    isVoidPtr = true;
                                }
                            }
                        }
                        if (isVoidPtr) {
                            // 模仿 line 1880-1938 COM 默认属性调用 (comPackExpr/resolveComMarkerForPack)
                            emitExpr(*node.callee);  // emits "obj->member  /* ... voidptr */"
                            std::string objExpr = std::move(lastExpr_);
                            std::vector<std::string> packedArgs;
                            for (size_t i = 0; i < node.positional.size(); i++) {
                                std::string packFn = comPackExpr(*node.positional[i]);
                                emitExpr(*node.positional[i]);
                                { std::string resolved = resolveComMarkerForPack(packFn);
                                  if (!resolved.empty()) lastExpr_ = resolved; }
                                packedArgs.push_back(packArgValue(packFn, lastExpr_));
                            }
                            int32_t argc = (int32_t)packedArgs.size();
                            std::string argsArray = "(void*[]){";
                            for (int i = 0; i < argc; i++) {
                                if (i > 0) argsArray += ", ";
                                argsArray += packedArgs[i];
                            }
                            argsArray += "}";
                            // 默认成员名 "Item" — VB6 Collection / ADODB.Recordset 等大多数
                            // void* 字段都是用 .Item(idx) 索引的 (DISPID_VALUE 默认成员).
                            lastExpr_ = "vb6_ComCall(" + objExpr + ", L\"Item\", "
                                      + argsArray + ", " + std::to_string(argc) + ")";
                            handled = true;
                        } else if (classTypedFieldMap_) {
                            // Fix 037b: typed (非 void*) 对象字段 — 项目类或 COM 接口
                            auto itT = classTypedFieldMap_->find(className);
                            if (itT != classTypedFieldMap_->end()) {
                                auto itF = itT->second.find(memLower);
                                if (itF == itT->second.end()) itF = itT->second.find(memLowerM);
                                if (itF != itT->second.end()) {
                                    const std::string& fieldType = itF->second;
                                    if (fieldType.compare(0, 4, "COM:") == 0) {
                                        // Pattern K: COM 接口字段 (如 Header As Dictionary)
                                        // → vb6_ComCall((void*)obj->member, L"Item", args, argc)
                                        emitExpr(*node.callee);
                                        std::string objExpr = std::move(lastExpr_);
                                        std::vector<std::string> packedArgs;
                                        for (size_t i = 0; i < node.positional.size(); i++) {
                                            std::string packFn = comPackExpr(*node.positional[i]);
                                            emitExpr(*node.positional[i]);  // Fix: Pattern K 实参丢失 (漏 emit 致 vb6_ComPackXxx() 零实参), 与 Pattern D 对齐

                                            { std::string resolved = resolveComMarkerForPack(packFn);
                                              if (!resolved.empty()) lastExpr_ = resolved; }
                                            packedArgs.push_back(packArgValue(packFn, lastExpr_));
                                        }
                                        int32_t argc = (int32_t)packedArgs.size();
                                        std::string argsArray = "(void*[]){";
                                        for (int i = 0; i < argc; i++) {
                                            if (i > 0) argsArray += ", ";
                                            argsArray += packedArgs[i];
                                        }
                                        argsArray += "}";
                                        lastExpr_ = "vb6_ComCall((void*)" + objExpr + ", L\"Item\", "
                                                  + argsArray + ", " + std::to_string(argc) + ")";
                                        handled = true;
                                    } else if (node.positional.size() == 1) {
                                        // Pattern L: 项目类字段 (如 Rows As cCollection)
                                        // → vb6_<Type>_prop_get_Item(obj->member, vb6_VariantFromValue(arg))
                                        std::string itemFn = resolveClassMemberCall(fieldType, "Item");
                                        if (!itemFn.empty()) {
                                            emitExpr(*node.callee);
                                            std::string objExpr = std::move(lastExpr_);
                                            emitExpr(*node.positional[0]);
                                            std::string arg = std::move(lastExpr_);
                                        // 类虚表 (tB, ai/022 B08e-5 第⑦站): 默认属性调用 `m_up(4)` 此前静默发
                                        // `vb6_<Base>_prop_get_Item(接收者, ...)`, 接收者声明成基类而实际持派生实例时会切片回
                                        // 基类的 Item; 显式写 `m_up.Item(4)` 走的是 ①/B08d 那条早已派发的路, 两种写法必须一致。
                                        // 口径同 ①⑤⑪: 派发串与实参里的接收者是同一个 objExpr, 不纯时由 mustDispatch 出诊断,
                                        // 左值求值 (suppressVirtDispatch_) 不改派。
                                        if (!suppressVirtDispatch_) {
                                            const std::string dispatchFnL = virtDispatchCallee(
                                                fieldType, "Item", objExpr, /*mustDispatch=*/true, node.loc);
                                            if (!dispatchFnL.empty()) itemFn = dispatchFnL;
                                        }
                                            lastExpr_ = itemFn + "(" + objExpr + ", vb6_VariantFromValue(" + arg + "))";
                                            lastExprNeedsObjectUnpack_ = true;  // Set 语句需转 void*
                                            handled = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if (handled) return;
    }
