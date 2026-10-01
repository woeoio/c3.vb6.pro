// cgen_form_wndproc_dispatch.inc - emitFormFramework 函数体片段：WndProc 后半 —— 焦点 / 滚动 / 键盘 / 关闭 / 清理 / default
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 922~1272 行，逐行未改。

    // P18-B: GotFocus/LostFocus events (EN_SETFOCUS/EN_KILLFOCUS/CBN_SETFOCUS/CBN_KILLFOCUS)
    {
        // 账 #158：这张焦点码表**只写一遍**，顶层与容器子控件**两条创建路共用**。以前只有顶层
        // 那一趟（ID 100+）⇒ 放在 Frame/PictureBox 里的按钮与文本框连 arm 都没发过
        // —— 发码侧的读数：同一份 `.frm`，`cmdTop_GotFocus` 有 `id == 101 && code == 6`、
        // `cmdIn_GotFocus` 一条都没有（容器子控件的 ID 从 200 起，那一趟根本没被遍历）。
        // ID 口径与创建循环一致（跳过 Menu 与没有 Win32 窗口类的控件，DFS 顺序同步）。
        auto emitFocusArms158 = [&](const FrmControl& ctrl, int focusId) {
            int ctrlId = focusId;
            std::string gotFn = cProcName(ctrl.controlName + "_GotFocus", AccessLevel::Private);
            std::string lostFn = cProcName(ctrl.controlName + "_LostFocus", AccessLevel::Private);
            bool hasGot = symTab_.lookup(ctrl.controlName + "_GotFocus") != nullptr;
            bool hasLost = symTab_.lookup(ctrl.controlName + "_LostFocus") != nullptr;
            bool hasValidate = symTab_.lookup(ctrl.controlName + "_Validate") != nullptr;
            if (hasGot || hasLost) {
                if (ctrl.controlType == FrmControlType::TextBox) {
                    // EN_SETFOCUS=256, EN_KILLFOCUS=512
                    if (hasGot) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 256) { extern void " + gotFn + "(); " + gotFn + "(); }");
                    if (hasLost) { if (hasValidate) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 512 && !GetPropW((HWND)lParam, L\"VB6_ValidateCancel\")) { extern void " + lostFn + "(); " + lostFn + "(); }"); else c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 512) { extern void " + lostFn + "(); " + lostFn + "(); }"); }
                } else if (ctrl.controlType == FrmControlType::ComboBox) {
                    // 账 #158：这里原本发的是 1024 / 2048 —— 那是 **CBEM_*（发给 ComboBoxEx 窗口的
                    // 消息号）**，根本不是 WM_COMMAND 的 notification code，所以这两个分支永远等不到。
                    // 真正的码表是 `CBN_SETFOCUS=3` / `CBN_KILLFOCUS=4`（对照同一段里
                    // EN_=256/512、LBN_=4/5、BN_=6/7 都是对的 —— 只有 Combo 这一格抄错了表）。
                    if (hasGot) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 3) { extern void " + gotFn + "(); " + gotFn + "(); }");
                    if (hasLost) { if (hasValidate) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 4 && !GetPropW((HWND)lParam, L\"VB6_ValidateCancel\")) { extern void " + lostFn + "(); " + lostFn + "(); }"); else c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 4) { extern void " + lostFn + "(); " + lostFn + "(); }"); }
                } else if (ctrl.controlType == FrmControlType::ListBox) {
                    // LBN_SETFOCUS=4, LBN_KILLFOCUS=5
                    if (hasGot) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 4) { extern void " + gotFn + "(); " + gotFn + "(); }");
                    if (hasLost) { if (hasValidate) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 5 && !GetPropW((HWND)lParam, L\"VB6_ValidateCancel\")) { extern void " + lostFn + "(); " + lostFn + "(); }"); else c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 5) { extern void " + lostFn + "(); " + lostFn + "(); }"); }
                } else if (ctrl.controlType == FrmControlType::CommandButton ||
                           ctrl.controlType == FrmControlType::CheckBox ||
                           ctrl.controlType == FrmControlType::OptionButton) {
                    // BN_SETFOCUS=6, BN_KILLFOCUS=7 (button class notifications via WM_COMMAND)
                    if (hasGot) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 6) { extern void " + gotFn + "(); " + gotFn + "(); }");
                    if (hasLost) { if (hasValidate) c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 7 && !GetPropW((HWND)lParam, L\"VB6_ValidateCancel\")) { extern void " + lostFn + "(); " + lostFn + "(); }"); else c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 7) { extern void " + lostFn + "(); " + lostFn + "(); }"); }
                }
                // Label: no GotFocus/LostFocus in VB6 (windowless control)
                // PictureBox: GotFocus/LostFocus handled via subclassing (P18-F)
            }
        };
        ctrlId = 100;
        for (const auto& ctrl : frmDesc.formControl.children) {
            emitFocusArms158(ctrl, ctrlId);
            ctrlId++;
        }
        int childFocusId158 = 200;
        auto walkFocusChildren158 = [&](auto&& self, const FrmControl& container) -> void {
            for (const auto& child : container.children) {
                if (child.controlType == FrmControlType::Menu) continue;
                if (!FrmParser::controlTypeToWin32Class(child.controlType)) continue;
                emitFocusArms158(child, childFocusId158);
                childFocusId158++;
                if (!child.children.empty() && controlIsContainerType(child)) {   // 账 #167: 原来少认 SSTab
                    self(self, child);
                }
            }
        };
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (!ctrl.children.empty() && controlIsContainerType(ctrl)) {   // 账 #167: 原来少认 SSTab
                walkFocusChildren158(walkFocusChildren158, ctrl);
            }
        }
    }
    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");

    // ============================================================
    // P20-42: WM_NOTIFY —— C3 此前**一个 WM_NOTIFY 分支都没有**。
    // comctl32 的通用控件 (SysTabControl32 等) 不用 WM_COMMAND 报事件, 一律走
    // WM_NOTIFY: lParam 指向 NMHDR{hwndFrom, idFrom, code}。
    // 这里只接 SSTab 的 TCN_SELCHANGE (code = -551): 交给 RTL 换算成"切换前的
    // 页号"并刷新子控件可见性, 再把该值作为参数调 <控件名>_Click。
    // ============================================================
    {
        bool anyNotify = false;
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (ctrl.controlType == FrmControlType::SSTab) {
                if (!symTab_.lookup(ctrl.controlName + "_Click")) continue;
                anyNotify = true;
                break;
            }
            // C29-7: ListView 的 ItemClick / ColumnClick 也走这条 WM_NOTIFY。
            // 两个事件共用一个 case 分支, 只要有一个处理器存在就得把分支建起来。
            if (ctrl.controlType == FrmControlType::ListView) {
                if (!symTab_.lookup(ctrl.controlName + "_ItemClick") &&
                    !symTab_.lookup(ctrl.controlName + "_ColumnClick")) continue;
                anyNotify = true;
                break;
            }
            // C29-8c: TreeView 的三条通知事件同理 (NodeClick / Expand / Collapse)。
            if (ctrl.controlType == FrmControlType::TreeView) {
                if (!symTab_.lookup(ctrl.controlName + "_NodeClick") &&
                    !symTab_.lookup(ctrl.controlName + "_Expand") &&
                    !symTab_.lookup(ctrl.controlName + "_Collapse")) continue;
                anyNotify = true;
                break;
            }
            // C29-4: StatusBar 的 PanelClick / PanelDblClick 同理。
            if (ctrl.controlType == FrmControlType::StatusBar) {
                if (!symTab_.lookup(ctrl.controlName + "_PanelClick") &&
                    !symTab_.lookup(ctrl.controlName + "_PanelDblClick")) continue;
                anyNotify = true;
                break;
            }
            // C29-5c: Toolbar 只有 _ButtonMenuClick 走这条 (TBN_DROPDOWN)；
            // _ButtonClick 在 WM_COMMAND 那一路 (见 cgen_form_wndproc_create.inc)。
            if (ctrl.controlType == FrmControlType::Toolbar) {
                if (!symTab_.lookup(ctrl.controlName + "_ButtonMenuClick")) continue;
                anyNotify = true;
                break;
            }
            // C29-DT-c: DTPicker 的三条 DTN_* 通知同理 (Change / DropDown / CloseUp)。
            if (ctrl.controlType == FrmControlType::DTPicker) {
                if (!symTab_.lookup(ctrl.controlName + "_Change") &&
                    !symTab_.lookup(ctrl.controlName + "_DropDown") &&
                    !symTab_.lookup(ctrl.controlName + "_CloseUp")) continue;
                anyNotify = true;
                break;
            }
            // C29-MV-c: MonthView 的 DateClick (原生 MCN_SELCHANGE) 同理。
            if (ctrl.controlType == FrmControlType::MonthView) {
                if (!symTab_.lookup(ctrl.controlName + "_DateClick")) continue;
                anyNotify = true;
                break;
            }
            // C29-RT-d: RichTextBox 的 SelChange (原生 EN_SELCHANGE) 同理。
            // 前提在 RTL 那侧：ENM_SELCHANGE 得由 vb6_RTB_Init 显式开（默认掩码是空的）。
            if (ctrl.controlType == FrmControlType::RichTextBox) {
                if (!symTab_.lookup(ctrl.controlName + "_SelChange")) continue;
                anyNotify = true;
                break;
            }
        }
        if (anyNotify) {
            c_.emitLine("case WM_NOTIFY: {");
            c_.indent();
            c_.emitLine("if (vb6_formLoading_" + cIdent(formName) + ") break;  /* block events during form init */");
            c_.emitLine("NMHDR* pNM42 = (NMHDR*)lParam;");
            c_.emitLine("if (!pNM42) break;");
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::SSTab) continue;
                auto* clickSym = symTab_.lookup(ctrl.controlName + "_Click");
                if (!clickSym) continue;
                std::string clickFn = cProcName(clickSym->name, AccessLevel::Private);
                std::string ctlNLower = ctrl.controlName;
                std::transform(ctlNLower.begin(), ctlNLower.end(), ctlNLower.begin(), ::tolower);
                std::string hwndVar = knownControlArrays_.count(ctlNLower)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                // TCN_SELCHANGE = -551 (TCN_FIRST=-550, TCN_SELCHANGE=TCN_FIRST-1)
                c_.emitLine("if (pNM42->code == -551 && (void*)pNM42->hwndFrom == "
                            + hwndVar + ") {");
                c_.indent();
                c_.emitLine("extern void " + clickFn + "(int16_t*);");
                c_.emitLine("int16_t vb6_prev42 = (int16_t)vb6_SSTab_OnSelChange((void*)"
                            + hwndVar + ");");
                // 页号没变就不触发: 重建标签项 (Tabs/TabVisible 改动后) 时 comctl32
                // 会自发一条 prev==cur 的噪声 TCN_SELCHANGE, VB6 语义里"没切换"不算 Click。
                c_.emitLine("if (vb6_prev42 != (int16_t)vb6_SSTab_GetTab((void*)"
                            + hwndVar + ")) {");
                c_.indent();
                c_.emitLine(clickFn + "(&vb6_prev42);");
                c_.dedent();
                c_.emitLine("}");
                c_.dedent();
                c_.emitLine("}");
            }
            // ---- C29-7: ListView 的 ItemClick / ColumnClick ----
            // 两个事件都从 WM_NOTIFY 来; RTL 的 vb6_ListView_OnNotify 把通知换算成
            // **1 基下标**, 再按下标取成员对象回调 —— VB6 的事件参数是**对象**:
            //   Sub ListView1_ItemClick(ByVal Item As ListItem)
            //   Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
            // 生成的回调形参因此是 void** (对象指针的地址)。
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::ListView) continue;
                auto* icSym = symTab_.lookup(ctrl.controlName + "_ItemClick");
                auto* ccSym = symTab_.lookup(ctrl.controlName + "_ColumnClick");
                if (!icSym && !ccSym) continue;
                std::string lvNameLower = ctrl.controlName;
                std::transform(lvNameLower.begin(), lvNameLower.end(), lvNameLower.begin(), ::tolower);
                std::string lvHwndVar = knownControlArrays_.count(lvNameLower)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                if (icSym) {
                    // LVN_ITEMACTIVATE = LVN_FIRST - 14 = -114 (双击/回车激活一行)
                    std::string icFn = cProcName(icSym->name, AccessLevel::Private);
                    c_.emitLine("if (pNM42->code == -114 && (void*)pNM42->hwndFrom == "
                                + lvHwndVar + ") {");
                    c_.indent();
                    c_.emitLine("extern void " + icFn + "(void*);");
                    c_.emitLine("int32_t vb6_lvIdx7 = vb6_ListView_OnNotify((void*)"
                                + lvHwndVar + ", -114, (void*)lParam);");
                    c_.emitLine("if (vb6_lvIdx7 > 0) {");
                    c_.indent();
                    c_.emitLine("void* vb6_lvItem7 = vb6_ListView_ListItemAt((void*)"
                                + lvHwndVar + ", vb6_lvIdx7);");
                    c_.emitLine("if (vb6_lvItem7) " + icFn + "(vb6_lvItem7);");
                    c_.dedent();
                    c_.emitLine("}");
                    c_.dedent();
                    c_.emitLine("}");
                }
                if (ccSym) {
                    // LVN_COLUMNCLICK = LVN_FIRST - 8 = -108 (点列头)
                    std::string ccFn = cProcName(ccSym->name, AccessLevel::Private);
                    c_.emitLine("if (pNM42->code == -108 && (void*)pNM42->hwndFrom == "
                                + lvHwndVar + ") {");
                    c_.indent();
                    c_.emitLine("extern void " + ccFn + "(void*);");
                    c_.emitLine("int32_t vb6_lvCol7 = vb6_ListView_OnNotify((void*)"
                                + lvHwndVar + ", -108, (void*)lParam);");
                    c_.emitLine("if (vb6_lvCol7 > 0) {");
                    c_.indent();
                    c_.emitLine("void* vb6_lvHdr7 = vb6_ListView_ColumnHeaderAt((void*)"
                                + lvHwndVar + ", vb6_lvCol7);");
                    c_.emitLine("if (vb6_lvHdr7) " + ccFn + "(vb6_lvHdr7);");
                    c_.dedent();
                    c_.emitLine("}");
                    c_.dedent();
                    c_.emitLine("}");
                }
            }
            // ---- C29-4: StatusBar 的 PanelClick / PanelDblClick ----
            // 原生 status bar 把**1 基面板号放在 NMHDR.idFrom**。RTL 的
            // vb6_StatusBar_OnNotify 校验 hwndFrom 并把 idFrom 折成返回值;
            // 再拿它造 Panel 对象 (vb6_StatusBar_PanelAt) 回调 —— VB6 的参数是对象:
            //   Sub StatusBar1_PanelClick(ByVal Panel As Panel)
            // 形状与上面 ListView 那两条完全同构。
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::StatusBar) continue;
                auto* pcSym = symTab_.lookup(ctrl.controlName + "_PanelClick");
                auto* pdSym = symTab_.lookup(ctrl.controlName + "_PanelDblClick");
                if (!pcSym && !pdSym) continue;
                std::string sbNameLower = ctrl.controlName;
                std::transform(sbNameLower.begin(), sbNameLower.end(), sbNameLower.begin(), ::tolower);
                std::string sbHwndVar = knownControlArrays_.count(sbNameLower)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                // (code, handler) 对: NM_CLICK=-2 → PanelClick, NM_DBLCLK=-6 → PanelDblClick
                for (auto evt : { std::make_pair(-2, pcSym), std::make_pair(-6, pdSym) }) {
                    if (!evt.second) continue;
                    std::string fn = cProcName(evt.second->name, AccessLevel::Private);
                    c_.emitLine("if (pNM42->code == " + std::to_string(evt.first)
                                + " && (void*)pNM42->hwndFrom == " + sbHwndVar + ") {");
                    c_.indent();
                    c_.emitLine("extern void " + fn + "(void*);");
                    c_.emitLine("int32_t vb6_sbPnl4 = vb6_StatusBar_OnNotify((void*)"
                                + sbHwndVar + ", " + std::to_string(evt.first)
                                + ", (void*)lParam);");
                    c_.emitLine("if (vb6_sbPnl4 > 0) {");
                    c_.indent();
                    c_.emitLine("void* vb6_sbPanel4 = vb6_StatusBar_PanelAt((void*)"
                                + sbHwndVar + ", vb6_sbPnl4);");
                    c_.emitLine("if (vb6_sbPanel4) " + fn + "(vb6_sbPanel4);");
                    c_.dedent();
                    c_.emitLine("}");
                    c_.dedent();
                    c_.emitLine("}");
                }

            }
            // ---- C29-8c: TreeView 的 NodeClick / Expand / Collapse ----
            // TVN_SELCHANGEDW(-451) 走 NodeClick；TVN_ITEMEXPANDEDW(-455) 按 action 分给
            // Expand / Collapse。通知负载 NMTREEVIEWW 里的 itemNew.hItem 由 RTL 折成 1 基
            // 节点序号，再拿 vb6_TreeView_NodeAt 造出 Node 对象回调 —— VB6 的处理器参数就
            // 是这个对象:
            //   Sub TreeView1_NodeClick(ByVal Node As Node)
            //   Sub TreeView1_Expand(ByVal Node As Node)   /   _Collapse 同形
            // 形状与上面 ListView / StatusBar 那两条完全同构 (共用同一个 anyNotify 分支)。
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::TreeView) continue;
                auto* ncSym = symTab_.lookup(ctrl.controlName + "_NodeClick");
                auto* exSym = symTab_.lookup(ctrl.controlName + "_Expand");
                auto* coSym = symTab_.lookup(ctrl.controlName + "_Collapse");
                if (!ncSym && !exSym && !coSym) continue;
                std::string tvNameLower = ctrl.controlName;
                std::transform(tvNameLower.begin(), tvNameLower.end(), tvNameLower.begin(), ::tolower);
                std::string tvHwndVar = knownControlArrays_.count(tvNameLower)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                // wantExp == 999 = 这条不看 action (NodeClick)；否则 -1 = 展开、0 = 折回。
                struct TvEvt { int32_t code; Symbol* sym; const char* tag; int32_t wantExp; };
                TvEvt tvEvts[] = {
                    { -451, ncSym, "nc8c", 999 },
                    { -455, exSym, "ex8c",  -1 },
                    { -455, coSym, "co8c",   0 } };
                for (const TvEvt& ev : tvEvts) {
                    if (!ev.sym) continue;
                    std::string fn = cProcName(ev.sym->name, AccessLevel::Private);
                    // ev.tag 是 const char* —— 与字面量直接相加就是"两个指针相加" (C2110)，
                    // 先把中段包成 std::string 再拼。
                    std::string vIdx = std::string("vb6_tvIdx8c_") + ev.tag;
                    std::string vNode = std::string("vb6_tvNode8c_") + ev.tag;
                    c_.emitLine("if (pNM42->code == " + std::to_string(ev.code)
                                + " && (void*)pNM42->hwndFrom == " + tvHwndVar + ") {");
                    c_.indent();
                    c_.emitLine("extern void " + fn + "(void*);");
                    c_.emitLine("int32_t " + vIdx + " = vb6_TreeView_NotifyNodeIndex((void*)"
                                + tvHwndVar + ", (void*)lParam);");
                    if (ev.wantExp == 999)
                        c_.emitLine("if (" + vIdx + " > 0) {");
                    else
                        c_.emitLine("if (" + vIdx + " > 0 && vb6_TreeView_NotifyExpanded("
                                    + "(void*)lParam) == " + std::to_string(ev.wantExp) + ") {");
                    c_.indent();
                    c_.emitLine("void* " + vNode + " = vb6_TreeView_NodeAt((void*)"
                                + tvHwndVar + ", " + vIdx + ");");
                    c_.emitLine("if (" + vNode + ") " + fn + "(" + vNode + ");");
                    c_.dedent();
                    c_.emitLine("}");
                    c_.dedent();
                    c_.emitLine("}");
                }
            }
            // ---- C29-5c: Toolbar 的 ButtonMenuClick ----
            // TBN_DROPDOWN(-710) 只有带下拉箭头的按钮 (VB6 的 Style 5 = BTNS_DROPDOWN) 发得
            // 出来；hdr.idFrom 就是那颗按钮的 idCommand —— 与 WM_COMMAND 那一路同一个值，
            // 所以造对象照用 vb6_Toolbar_ButtonAt:
            //   Sub Toolbar1_ButtonMenuClick(ByVal Button As Button)
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::Toolbar) continue;
                auto* mcSym = symTab_.lookup(ctrl.controlName + "_ButtonMenuClick");
                if (!mcSym) continue;
                std::string mcNameLower = ctrl.controlName;
                std::transform(mcNameLower.begin(), mcNameLower.end(), mcNameLower.begin(), ::tolower);
                std::string tbHwndVar = knownControlArrays_.count(mcNameLower)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                std::string mcFn = cProcName(mcSym->name, AccessLevel::Private);
                c_.emitLine("if (pNM42->code == -710 && (void*)pNM42->hwndFrom == " + tbHwndVar + ") {");
                c_.indent();
                c_.emitLine("extern void " + mcFn + "(void*);");
                c_.emitLine("void* vb6_tbMenuBtn5c = vb6_Toolbar_ButtonAt((void*)" + tbHwndVar
                            + ", (int32_t)pNM42->idFrom);");
                c_.emitLine("if (vb6_tbMenuBtn5c) " + mcFn + "(vb6_tbMenuBtn5c);");
                c_.dedent();
                c_.emitLine("}");
            }
            // ---- C29-DT-c: DTPicker 的 Change / DropDown / CloseUp ----
            // 三条都由 DTN_* 通知来 (SDK: DTN_FIRST2 = (0U-753U)，本机那份头逐条对过)：
            //   DTN_DATETIMECHANGE = -759   DTN_DROPDOWN = -754   DTN_CLOSEUP = -753
            // 与 ListView / TreeView 那两条的**关键差别**：VB6 的这三个处理器都没有参数
            //   Sub DTPicker1_Change()   /   _DropDown()   /   _CloseUp()
            // 所以不需要按负载造成员对象，认来源 (hwndFrom) + 分流 (code) 就是全部工作。
            // DTN_DATETIMECHANGE 的负载里那位 dwFlags (GDT_VALID / GDT_NONE) 这里刻意**不读**:
            // VB6 的 Change 对"勾掉复选框"和"拨了日期"一视同仁，两条都是"值变了"。
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::DTPicker) continue;
                struct DtEvt { int32_t code; const char* evt; const char* tag; };
                const DtEvt dtEvts[] = { { -759, "_Change", "chg" },
                                         { -754, "_DropDown", "drp" },
                                         { -753, "_CloseUp", "clu" } };
                std::string dtNameLower = ctrl.controlName;
                std::transform(dtNameLower.begin(), dtNameLower.end(), dtNameLower.begin(), ::tolower);
                std::string dtHwndVar = knownControlArrays_.count(dtNameLower)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                for (const DtEvt& ev : dtEvts) {
                    auto* evSym = symTab_.lookup(ctrl.controlName + ev.evt);
                    if (!evSym) continue;
                    std::string dtFn = cProcName(evSym->name, AccessLevel::Private);
                    c_.emitLine("if (pNM42->code == " + std::to_string(ev.code)
                                + " && (void*)pNM42->hwndFrom == " + dtHwndVar + ") {");
                    c_.indent();
                    c_.emitLine("{ extern void " + dtFn + "(); " + dtFn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
            }
            // ---- C29-MV-c: MonthView 的 DateClick ----
            // 原生一条 MCN_SELCHANGE(-749)，负载 NMSELCHANGE{nmhdr, stSelStart, stSelEnd}。
            // 与 DT-c 那三条的**关键差别**：VB6 这条处理器带一个 Date 参数
            //   Sub MonthView1_DateClick(ByVal DateSelected As Date)
            // ⇒ 要把负载折成 Date（vb6_MV_NotifyDate 干这件事）再传出去。
            // ⚠ **ABI 由 ByVal 决定**（实测：形状写错时 x64 侥幸读数正确、x86 当场错值）——
            //   cgen_decl_proc.cpp:94 那条规则是"ByVal ⇒ 原样传值，非 ByVal(即 ByRef) ⇒ 传指针",
            //   所以 Form_MouseDown 那批**没写 ByVal** 的参数才是 int16_t*/float*, 而这条
            //   `ByVal DateSelected As Date` 的定义发的是 `double DateSelected` ⇒ 按值传。
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::MonthView) continue;
                auto* mvSym = symTab_.lookup(ctrl.controlName + "_DateClick");
                if (!mvSym) continue;
                std::string mvNLower = ctrl.controlName;
                std::transform(mvNLower.begin(), mvNLower.end(), mvNLower.begin(), ::tolower);
                std::string mvNHwnd = knownControlArrays_.count(mvNLower)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                std::string mvFn = cProcName(mvSym->name, AccessLevel::Private);
                c_.emitLine("if (pNM42->code == -749 && (void*)pNM42->hwndFrom == "
                            + mvNHwnd + ") {");
                c_.indent();
                c_.emitLine("extern void " + mvFn + "(double);");
                c_.emitLine(mvFn + "(vb6_MV_NotifyDate((void*)lParam));");
                c_.dedent();
                c_.emitLine("}");
            }
            // ---- C29-RT-d: RichTextBox 的 SelChange ----
            // EN_SELCHANGE = 0x0702 = 1794（richedit.h:526，**正数** —— 这一族的码都不取负，
            // 与 DTN_*/MCN_*/TVN_* 那批 (0U-n) 反着来，抄惯 comctl32 的人最容易在这里写错号）。
            // VB6 这条处理器没有参数（Sub RichTextBox1_SelChange()），所以又不需要造负载对象。
            // 掩码前提见 vb6_RTB_Init 那段；同一事件的 WM_COMMAND/1815 那条 arm 在
            // cgen_form_wndproc_create.inc —— 原生到底走哪一条不稳定（实测同产物连跑三次
            // 1 次 1794 / 2 次 1815），所以两条都接；判据各钉一条，见夹具的 SimNotify 那几行。
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::RichTextBox) continue;
                auto* rtSelSym = symTab_.lookup(ctrl.controlName + "_SelChange");
                if (!rtSelSym) continue;
                std::string rtLower42 = ctrl.controlName;
                std::transform(rtLower42.begin(), rtLower42.end(), rtLower42.begin(), ::tolower);
                std::string rtNHwnd = knownControlArrays_.count(rtLower42)
                    ? ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)")
                    : ("vb6_hwnd_" + cIdent(ctrl.controlName));
                std::string rtSelFn = cProcName(rtSelSym->name, AccessLevel::Private);
                c_.emitLine("if (pNM42->code == 1794 && (void*)pNM42->hwndFrom == "
                            + rtNHwnd + ") {");
                c_.indent();
                c_.emitLine("{ extern void " + rtSelFn + "(); " + rtSelFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }

    // WM_CLOSE: 调用Form_Unload判断是否允许关闭
    c_.emitLine("case WM_CLOSE: {");
    c_.indent();
    {
        std::string formUnloadFn = cProcName("Form_Unload", AccessLevel::Private);
        auto* formUnloadSym2 = symTab_.lookup("Form_Unload");
        if (formUnloadSym2) {
            // 生成trampoline: 桥接 int(*)(void) 回调和 void(int16_t*) VB6签名
            std::string trampoline = formUnloadFn + "_trampoline";
            // P18-B: Form_QueryUnload - check if close is allowed before Unload
            {
                std::string queryUnloadFn = cProcName("Form_QueryUnload", AccessLevel::Private);
                auto* quSym = symTab_.lookup("Form_QueryUnload");
                if (quSym) {
                    c_.emitLine("{ int16_t vb6_cancel = 0; int16_t vb6_unloadmode = 0; /* 0=vbFormControlMenu */");
                    c_.emitLine("extern void " + queryUnloadFn + "(int16_t*, int16_t*); " + queryUnloadFn + "(&vb6_cancel, &vb6_unloadmode);");
                    c_.emitLine("if (vb6_cancel != 0) return 0; /* Cancel close */ }");
                }
            }
            c_.emitLine("vb6_SetFormUnloadCallback((void*)" + trampoline + ");");
            c_.emitLine("if (vb6_QueryFormUnload() == 0) {");
            c_.indent();
            c_.emitLine("DestroyWindow(hwnd);");
            c_.dedent();
            c_.emitLine("}");
        } else {
            c_.emitLine("DestroyWindow(hwnd);");
        }
    }
    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");
    // WM_DESTROY: PostQuitMessage (如果是主窗体)
    c_.emitLine("case WM_DESTROY: {");
    c_.indent();
    c_.emitLine("vb6_Forms_Unregister((void*)hwnd);");
    // P18-F: 移除所有控件子类化
    {
        // ⚠ 这一句判据在仓里有**三份**（这里 + frame_menu 的"装" + subclass 的"发 arm"），
        // 不同步就是"装了不拆"或"发了不装"。C29-SL-g 把三份一起对齐：焦点档改成按
        // "有没有另一条原生来源"排除，并补上 C29-SL-d 那一条 `_Click`（上一批只同步了
        // 前两份，这一份漏了 —— 漏了不会崩，只是拆的时候少拆一枚，记在这里免得再漏）。
        // C29-SL-n（账 #141）: 这份「要不要子类化」原本在这里又抄了一遍，而且比发 arm 那一趟
        // 少了 `_DblClick` 与 `_Paint` ⇒ 只挂这两条处理器之一的控件，子类过程与消息臂都生成了、
        // 却一次也没被 install（编得过、永不触发）。三份现在共用 controlNeedsSubclass。
        auto needsSubclass = [&](const FrmControl& ctrl) -> bool {
            return controlNeedsSubclass(ctrl);
        };
        std::unordered_set<std::string> subEmitted;
        for (const auto& ctrl : frmDesc.formControl.children) {
            std::string ctrlNameLower = ctrl.controlName;
            std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
            if (subEmitted.count(ctrlNameLower)) continue;
            subEmitted.insert(ctrlNameLower);
            if (!needsSubclass(ctrl)) continue;
            if (knownControlArrays_.count(ctrlNameLower)) {
                c_.emitLine("{ int vb6_si; for (vb6_si = 0; vb6_si < VB6_CTRLARR_MAX; vb6_si++) {");
                c_.emitLine("  void* vb6_sub_hwnd = vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", vb6_si);");
                c_.emitLine("  if (vb6_sub_hwnd) vb6_RemoveControlSubclass(vb6_sub_hwnd); } }");
            } else {
                c_.emitLine("vb6_RemoveControlSubclass((void*)vb6_hwnd_" + cIdent(ctrl.controlName) + ");");
            }
        }
        // P0-5修复: 容器内子控件（账 #167: 这一份「谁是容器」抄得比"装"那一趟还少 ——
        // 只有 Frame ⇒ PictureBox / SSTab 的孩子在 install 那一趟发了子类化、这里不拆。
        // 三份合到 controlIsContainerType 这一个出口。）
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (!controlIsContainerType(ctrl)) continue;
            for (const auto& child : ctrl.children) {
                std::string childNameLower = child.controlName;
                std::transform(childNameLower.begin(), childNameLower.end(), childNameLower.begin(), ::tolower);
                if (subEmitted.count(childNameLower)) continue;
                subEmitted.insert(childNameLower);
                if (!needsSubclass(child)) continue;
                c_.emitLine("vb6_RemoveControlSubclass((void*)vb6_hwnd_" + cIdent(child.controlName) + ");");
            }
        }
    }
    // M12-FIX: Cleanup Form.Picture and Form.Icon from .frx
    {
        auto picIt = frmDesc.formControl.properties.find("Picture");
        if (picIt != frmDesc.formControl.properties.end() && picIt->second.type == FrmValueType::FrxReference && frxLoaded) {
            c_.emitLine("{ HANDLE vb6_bg = GetPropW(hwnd, L\"VB6_Picture\");");
            c_.emitLine("  if (vb6_bg) { DeleteObject(vb6_bg); RemovePropW(hwnd, L\"VB6_Picture\"); } }");
        }
        auto iconIt = frmDesc.formControl.properties.find("Icon");
        if (iconIt != frmDesc.formControl.properties.end() && iconIt->second.type == FrmValueType::FrxReference && frxLoaded) {
            c_.emitLine("{ HANDLE vb6_ico = (HANDLE)SendMessageW(hwnd, WM_GETICON, ICON_BIG, 0);");
            c_.emitLine("  if (vb6_ico) DestroyIcon((HICON)vb6_ico); }");
        }
    }
    // Fix 144: 只有**启动窗体**的 WM_DESTROY 才 PostQuitMessage.
    // 此前无条件投 WM_QUIT: 用 Show vbModal 打开的子窗体一旦销毁就把主消息循环
    // 一起结束掉 — 症状是"点按钮打开子窗体后整个程序退出" (退出码 0).
    // Startup="Sub Main" 时由 main 返回结束进程, 窗体都不投.
    {
        bool postQuit = true;
        if (!startupObject_.empty()) {
            std::string so = Symbol::toLower(startupObject_);
            std::string fn = Symbol::toLower(formName);
            std::string modLower = Symbol::toLower(moduleName_);
            postQuit = (so == fn || so == modLower);
            if (so == "sub main") postQuit = false;
        }
        if (postQuit) {
            c_.emitLine("PostQuitMessage(0);  /* Fix 144: 仅启动窗体结束消息循环 */");
        }
    }
    // P20-43: Form_Terminate —— VB6 语义是"窗体实例被销毁时最后触发"。
    // 挂在 WM_DESTROY 里、PostQuitMessage 之前, 这是窗体生命周期的最后一站。
    {
        auto* termSym43 = symTab_.lookup(formName + "_Terminate");
        if (!termSym43) termSym43 = symTab_.lookup("Form_Terminate");
        if (termSym43) {
            std::string termFn43 = cProcName(termSym43->name, AccessLevel::Private);
            c_.emitLine("{ extern void " + termFn43 + "(); " + termFn43 + "(); }");
        }
    }
    // P20-44: 窗体销毁时撤销该窗体控件的 OLE 放置目标登记 (RevokeDragDrop 对
    // 已销毁的 hwnd 无害, 但不清表会留悬垂 hwnd 在 g_targets 里)。
    c_.emitLine("vb6_OLEDrop_RevokeAll();");
    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");

    // P14.4.1b: WM_HSCROLL / WM_VSCROLL - ScrollBar事件
    c_.emitLine("case WM_HSCROLL:");
    c_.emitLine("case WM_VSCROLL: {");
    c_.indent();
    bool hasScrollBar = false;
    for (const auto& ctrl : frmDesc.formControl.children) {
        if (ctrl.controlType == FrmControlType::HScrollBar || ctrl.controlType == FrmControlType::VScrollBar) { hasScrollBar = true; break; }
        // C29-SL-c: Slider 走的就是这一条通道（实测：横杆发 WM_HSCROLL、竖杆发 WM_VSCROLL，
        // lParam = 控件句柄）⇒ 同一扇门，认来源的写法也照抄 ScrollBar 那一套。
        if (ctrl.controlType == FrmControlType::Slider) { hasScrollBar = true; break; }
    }
    if (hasScrollBar) {
    c_.emitLine("{ void* scrollHwnd = (void*)lParam; int scrollCode = LOWORD(wParam);");
    // C29-SL-k: VB6 的 Slider 在拖滑块时把那颗气泡摆出来（ISlider.Text = 0x0010 就是气泡里那句
    // 串），而原生轨道条自己那枚 TBS_TOOLTIPS 服务不了自定义串（实测 slmeasure17.c：它自带一条
    // 工具、TTM_POP 无头里什么都不触发）⇒ 气泡由 RTL 自己持一枚 TRACK 型 tooltip 来摆。
    // 这一条**只发一次、不按控件展开**：RTL 里按窗口类名把 ScrollBar 筛掉。值取 wParam 高字
    // （实测 4/5 两档带新值、8 那档为 0 ⇒ 传 0 时 RTL 自己现问 TBM_GETPOS）。
    bool hasSlider = false;
    for (const auto& ctrl : frmDesc.formControl.children) {
        if (ctrl.controlType == FrmControlType::Slider) { hasSlider = true; break; }
    }
    if (hasSlider)
        c_.emitLine("vb6_Slider_BubbleNotify(scrollHwnd, scrollCode, HIWORD(wParam));");
    for (const auto& ctrl : frmDesc.formControl.children) {
        if (ctrl.controlType == FrmControlType::HScrollBar || ctrl.controlType == FrmControlType::VScrollBar) {
            std::string ctrlHwnd = "vb6_hwnd_" + cIdent(ctrl.controlName);
            std::string scrollFn = cProcName(ctrl.controlName + "_Scroll", AccessLevel::Private);
            std::string changeFn = cProcName(ctrl.controlName + "_Change", AccessLevel::Private);
            bool hasScroll = symTab_.lookup(ctrl.controlName + "_Scroll") != nullptr;
            bool hasChange = symTab_.lookup(ctrl.controlName + "_Change") != nullptr;
            c_.emitLine("if (scrollHwnd == (void*)" + ctrlHwnd + ") {");
            c_.indent();
            if (hasScroll) c_.emitLine("if (scrollCode == 5 || scrollCode == 4) { extern void " + scrollFn + "(); " + scrollFn + "(); }");
            if (hasChange) c_.emitLine("else { extern void " + changeFn + "(); " + changeFn + "(); }");
            c_.dedent();
            c_.emitLine("}");
        }
    }
    // C29-SL-c: Slider 的 Change / Scroll —— 同一条通道、同一扇认来源的门。
    // 两条分法都在 RTL 那边定（029 §C29-SL-c 有实测理由）：Change 问"值真的变了没"
    // （`vb6_Slider_FireChange` 自带基准，拖拽中那串 TB_THUMBTRACK 因此连续触发），
    // Scroll 只认滑块那两档（4/5），与上面 ScrollBar 已发货的口径同一档。
    for (const auto& ctrl : frmDesc.formControl.children) {
        if (ctrl.controlType != FrmControlType::Slider) continue;
        std::string slHwnd = "vb6_hwnd_" + cIdent(ctrl.controlName);
        std::string slScrollFn = cProcName(ctrl.controlName + "_Scroll", AccessLevel::Private);
        std::string slChangeFn = cProcName(ctrl.controlName + "_Change", AccessLevel::Private);
        bool slHasScroll = symTab_.lookup(ctrl.controlName + "_Scroll") != nullptr;
        bool slHasChange = symTab_.lookup(ctrl.controlName + "_Change") != nullptr;
        if (!slHasScroll && !slHasChange) continue;
        c_.emitLine("if (scrollHwnd == (void*)" + slHwnd + ") {");
        c_.indent();
        // FireChange 只在**真挂了 Change 处理器**时才调：它会推进基准，白调一次等于吞掉一条 Change。
        if (slHasChange)
            c_.emitLine("if (vb6_Slider_FireChange((void*)" + slHwnd + ")) { extern void "
                        + slChangeFn + "(); " + slChangeFn + "(); }");
        if (slHasScroll)
            c_.emitLine("if (scrollCode == 4 || scrollCode == 5) { extern void "
                        + slScrollFn + "(); " + slScrollFn + "(); }");
        c_.dedent();
        c_.emitLine("}");
    }
    c_.emitLine("}");
    }
    // P16: WithEvents HScrollBar/VScrollBar事件分发
    for (const auto& we : knownWithEventsCtrlVars_) {
        if (we.second == FrmControlType::HScrollBar || we.second == FrmControlType::VScrollBar) {
            std::string varName = knownWithEventsCtrlOrigNames_.count(we.first) ? knownWithEventsCtrlOrigNames_[we.first] : cIdent(we.first);
            auto* scrollSym = symTab_.lookup(we.first + "_scroll");
            auto* changeSym = symTab_.lookup(we.first + "_change");
            c_.emitLine("if (" + varName + " != NULL && (void*)" + varName + " == scrollHwnd) {");
            c_.indent();
            if (scrollSym) {
                std::string scrollFn = cProcName(scrollSym->name, AccessLevel::Private);
                c_.emitLine("if (scrollCode == 5 || scrollCode == 4) { extern void " + scrollFn + "(); " + scrollFn + "(); }");
            }
            if (changeSym) {
                std::string changeFn = cProcName(changeSym->name, AccessLevel::Private);
                c_.emitLine("else { extern void " + changeFn + "(); " + changeFn + "(); }");
            }
            c_.dedent();
            c_.emitLine("}");
        }
    }    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");

    // P14.4.1d: WM_SIZE -> Form_Resize() (仅当处理器存在时生成)
    {
        std::string resizeFn = cProcName("Form_Resize", AccessLevel::Private);
        if (symTab_.lookup(formName + "_Resize") || symTab_.lookup("Form_Resize")) {
            c_.emitLine("case WM_SIZE: {");
            c_.indent();
            // Fix 112: Form_Load 由 PostMessage 延迟执行, 而首次 WM_SIZE 在
            // CreateWindow 期间就到 — 若照搬 VB6 的 Load→Resize 顺序, 这里会在
            // Form_Load 之前触发 Form_Resize, 模块级对象 (如 cResizer) 还是
            // Nothing (空指针) · 故 Load 完成前跳过, 由 0x7FF0 分支补一次 Resize.
            c_.emitLine("if (!vb6_formLoaded_" + cIdent(formName) + ") break;");
            c_.emitLine("{ extern void " + resizeFn + "(); " + resizeFn + "(); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }

    // P14.4.1e: WM_KEYDOWN/WM_CHAR -> Form_KeyDown/KeyPress (仅当处理器存在时)
    {
        std::string keyDownFn = cProcName("Form_KeyDown", AccessLevel::Private);
        if (symTab_.lookup(formName + "_KeyDown") || symTab_.lookup("Form_KeyDown")) {
            c_.emitLine("case WM_KEYDOWN: {");
            c_.indent();
            c_.emitLine("{ int16_t vb6_keyCode = (int16_t)wParam; int16_t vb6_shift = 0; if (GetKeyState(VK_SHIFT) & 0x8000) vb6_shift |= 1; if (GetKeyState(VK_CONTROL) & 0x8000) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
            c_.emitLine("extern void " + keyDownFn + "(int16_t*, int16_t*); " + keyDownFn + "(&vb6_keyCode, &vb6_shift); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }
    {
        std::string keyPressFn = cProcName("Form_KeyPress", AccessLevel::Private);
        if (symTab_.lookup(formName + "_KeyPress") || symTab_.lookup("Form_KeyPress")) {
            c_.emitLine("case WM_CHAR: {");
            c_.indent();
            c_.emitLine("{ int16_t vb6_keyAscii = (int16_t)wParam;");
            c_.emitLine("extern void " + keyPressFn + "(int16_t*); " + keyPressFn + "(&vb6_keyAscii); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }
    // P18-B: Form_KeyUp (WM_KEYUP)
    {
        std::string keyUpFn = cProcName("Form_KeyUp", AccessLevel::Private);
        if (symTab_.lookup(formName + "_KeyUp") || symTab_.lookup("Form_KeyUp")) {
            c_.emitLine("case WM_KEYUP: {");
            c_.indent();
            c_.emitLine("{ int16_t vb6_keyCode = (int16_t)wParam; int16_t vb6_shift = 0; if (GetKeyState(VK_SHIFT) & 0x8000) vb6_shift |= 1; if (GetKeyState(VK_CONTROL) & 0x8000) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
            c_.emitLine("extern void " + keyUpFn + "(int16_t*, int16_t*); " + keyUpFn + "(&vb6_keyCode, &vb6_shift); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }
    // P18-B: Form_MouseDown/MouseUp/MouseMove (WM_LBUTTONDOWN/UP/MOVE)
    {
        std::string mouseDownFn = cProcName("Form_MouseDown", AccessLevel::Private);
        std::string mouseUpFn = cProcName("Form_MouseUp", AccessLevel::Private);
        std::string mouseMoveFn = cProcName("Form_MouseMove", AccessLevel::Private);
        bool hasMouseDown = symTab_.lookup(formName + "_MouseDown") || symTab_.lookup("Form_MouseDown");
        bool hasMouseUp = symTab_.lookup(formName + "_MouseUp") || symTab_.lookup("Form_MouseUp");
        bool hasMouseMove = symTab_.lookup(formName + "_MouseMove") || symTab_.lookup("Form_MouseMove");
        // P20-43: Form_Click —— VB6 在"按下并抬起"时触发。WM_LBUTTONUP 与 MouseUp
        // 是同一条消息, **不能各自发一个 case** (C2196 case 值重复), 必须并进同一组:
        // 有 MouseUp 处理器就共用它那组 case, 在里面补一次 Click 调用。
        auto* clickSym43 = symTab_.lookup(formName + "_Click");
        if (!clickSym43) clickSym43 = symTab_.lookup("Form_Click");
        bool hasFormClick43 = clickSym43 != nullptr;
        std::string formClickFn43 = hasFormClick43 ? cProcName(clickSym43->name, AccessLevel::Private) : std::string();
        if (hasMouseDown) {
            c_.emitLine("case WM_LBUTTONDOWN: case WM_RBUTTONDOWN: case WM_MBUTTONDOWN: {");
            c_.indent();
            c_.emitLine("{ int16_t vb6_button = 0; int16_t vb6_shift = 0;");
            c_.emitLine("if (msg == WM_LBUTTONDOWN) vb6_button = 1;");
            c_.emitLine("else if (msg == WM_RBUTTONDOWN) vb6_button = 2;");
            c_.emitLine("else if (msg == WM_MBUTTONDOWN) vb6_button = 4;");
            c_.emitLine("if (wParam & MK_SHIFT) vb6_shift |= 1; if (wParam & MK_CONTROL) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
            c_.emitLine("float vb6_x = (float)(int16_t)LOWORD(lParam); float vb6_y = (float)(int16_t)HIWORD(lParam);");
            c_.emitLine("extern void " + mouseDownFn + "(int16_t*, int16_t*, float*, float*); " + mouseDownFn + "(&vb6_button, &vb6_shift, &vb6_x, &vb6_y); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
        if (hasMouseUp || hasFormClick43) {
            c_.emitLine("case WM_LBUTTONUP: case WM_RBUTTONUP: case WM_MBUTTONUP: {");
            c_.indent();
            // Fix 161f: 参数计算裸块只在真有 Form_MouseUp 处理器时才发。
            // 此前 "{ int16_t vb6_button ... " 的闭合 "}" 挂在 MouseUp 调用行尾,
            // 工程"只有 Form_Click 没有 Form_MouseUp"时裸块永不闭合 ⇒ 后续所有
            // 函数定义被吞进 WndProc ⇒ C2143/C2065 雪崩 (extlist 实测)。
            if (hasMouseUp) {
                c_.emitLine("{ int16_t vb6_button = 0; int16_t vb6_shift = 0;");
                c_.emitLine("if (msg == WM_LBUTTONUP) vb6_button = 1;");
                c_.emitLine("else if (msg == WM_RBUTTONUP) vb6_button = 2;");
                c_.emitLine("else if (msg == WM_MBUTTONUP) vb6_button = 4;");
                c_.emitLine("if (wParam & MK_SHIFT) vb6_shift |= 1; if (wParam & MK_CONTROL) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
                c_.emitLine("float vb6_x = (float)(int16_t)LOWORD(lParam); float vb6_y = (float)(int16_t)HIWORD(lParam);");
                c_.emitLine("extern void " + mouseUpFn + "(int16_t*, int16_t*, float*, float*); " + mouseUpFn + "(&vb6_button, &vb6_shift, &vb6_x, &vb6_y); }");
            }
            // P20-43: 同一条消息里补 Form_Click (有处理器才发)
            if (hasFormClick43) {
                c_.emitLine("{ extern void " + formClickFn43 + "(); " + formClickFn43 + "(); }");
            }
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
        if (hasMouseMove) {
            c_.emitLine("case WM_MOUSEMOVE: {");
            c_.indent();
            c_.emitLine("{ int16_t vb6_button = 0; int16_t vb6_shift = 0;");
            c_.emitLine("if (wParam & MK_LBUTTON) vb6_button |= 1; if (wParam & MK_RBUTTON) vb6_button |= 2; if (wParam & MK_MBUTTON) vb6_button |= 4;");
            c_.emitLine("if (wParam & MK_SHIFT) vb6_shift |= 1; if (wParam & MK_CONTROL) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
            c_.emitLine("float vb6_x = (float)(int16_t)LOWORD(lParam); float vb6_y = (float)(int16_t)HIWORD(lParam);");
            c_.emitLine("extern void " + mouseMoveFn + "(int16_t*, int16_t*, float*, float*); " + mouseMoveFn + "(&vb6_button, &vb6_shift, &vb6_x, &vb6_y); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }
    // P18-B: Form_DblClick (WM_LBUTTONDBLCLK)
    {
        // P20-43: 调用名必须取**符号真名**。此前写死 formName + "_DblClick",
        // 而用户写的是 `Form_DblClick` → 定义落在 vb6_Form_DblClick, 调用却发
        // vb6_<窗体名>_DblClick → LNK2019 (与 Activate 同一处缺陷)。
        auto* dblSym43 = symTab_.lookup(formName + "_DblClick");
        if (!dblSym43) dblSym43 = symTab_.lookup("Form_DblClick");
        std::string dblClickFn = dblSym43 ? cProcName(dblSym43->name, AccessLevel::Private) : std::string();
        if (dblSym43) {
            c_.emitLine("case WM_LBUTTONDBLCLK: {");
            c_.indent();
            c_.emitLine("extern void " + dblClickFn + "(); " + dblClickFn + "();");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }
    // P18-B: WM_ACTIVATE -> Form_Activate/Deactivate
    {
        // P20-43: 同 DblClick —— 调用名取符号真名。写死 formName + "_Activate"
        // 会让 `Form_Activate` 的定义 (vb6_Form_Activate) 连不上 → LNK2019。
        auto* actSym43 = symTab_.lookup(formName + "_Activate");
        if (!actSym43) actSym43 = symTab_.lookup("Form_Activate");
        auto* deactSym43 = symTab_.lookup(formName + "_Deactivate");
        if (!deactSym43) deactSym43 = symTab_.lookup("Form_Deactivate");
        bool hasActivate = actSym43 != nullptr;
        bool hasDeactivate = deactSym43 != nullptr;
        if (hasActivate || hasDeactivate) {
            c_.emitLine("case WM_ACTIVATE: {");
            c_.indent();
            if (hasActivate) {
                std::string activateFn = cProcName(actSym43->name, AccessLevel::Private);
                c_.emitLine("if (LOWORD(wParam) != WA_INACTIVE) { extern void " + activateFn + "(); " + activateFn + "(); }");
            }
            if (hasDeactivate) {
                std::string deactivateFn = cProcName(deactSym43->name, AccessLevel::Private);
                c_.emitLine("if (LOWORD(wParam) == WA_INACTIVE) { extern void " + deactivateFn + "(); " + deactivateFn + "(); }");
            }
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }
    // ============================================================
    // P20-43: 补齐窗体常用事件 (调用名一律取**符号真名**, 与 Activate 同一修法)。
    // ============================================================
    // Form_Paint (WM_PAINT) —— **已有**: create.inc 的 WM_PAINT 分支里
    // (BeginPaint/EndPaint 之间, 还捎带 OCX 绘制), 那边的命名 bug 已一并修掉,
    // 这里不能再发一个 case, 否则 C2196 case 值重复。
    // Form_GotFocus / Form_LostFocus (WM_SETFOCUS / WM_KILLFOCUS)
    {
        auto* gotSym43 = symTab_.lookup(formName + "_GotFocus");
        if (!gotSym43) gotSym43 = symTab_.lookup("Form_GotFocus");
        auto* lostSym43 = symTab_.lookup(formName + "_LostFocus");
        if (!lostSym43) lostSym43 = symTab_.lookup("Form_LostFocus");
        if (gotSym43) {
            std::string gotFn43 = cProcName(gotSym43->name, AccessLevel::Private);
            c_.emitLine("case WM_SETFOCUS: {");
            c_.indent();
            c_.emitLine("{ extern void " + gotFn43 + "(); " + gotFn43 + "(); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
        if (lostSym43) {
            std::string lostFn43 = cProcName(lostSym43->name, AccessLevel::Private);
            c_.emitLine("case WM_KILLFOCUS: {");
            c_.indent();
            c_.emitLine("{ extern void " + lostFn43 + "(); " + lostFn43 + "(); }");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }

    // P24-Timer: WM_TIMER分发 (窗口关联定时器触发)
    c_.emitLine("case WM_TIMER: {");
    c_.indent();
    c_.emitLine("vb6_DispatchTimer((int)wParam);");
    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");

    // default: DefWindowProc (P7.7: MDI窗体使用DefFrameProc/DefMDIChildProc)
    // Fix 190: 必须用 W 版。窗口类现在由 RegisterClassExW 注册, 窗体的 Unicode
    // 标志随之打开 —— DefWindowProcA 会把 WM_SETTEXT/WM_GETTEXT 的 **宽字符串**
    // 当 ANSI 字节串处理, 于是标题/文本只剩第一个字符 (实测 Form1 -> "F",
    // 探针窗口 "PROBE" -> "P"; 同文件的控件子类过程本来就用 DefWindowProcW)。
    c_.emitLine("default:");
    c_.indent();
    if (isMDIForm) {
        c_.emitLine("return DefFrameProcW(hwnd, (HWND)vb6_GetMDIClient(hwnd), msg, wParam, lParam);");
    } else if (isMDIChild) {
        c_.emitLine("return DefMDIChildProcW(hwnd, msg, wParam, lParam);");
    } else {
        c_.emitLine("return DefWindowProcW(hwnd, msg, wParam, lParam);");
    }
    c_.dedent();

    c_.dedent();
    c_.emitLine("}");  // switch
    c_.emitLine("return 0;");
    c_.dedent();
    c_.emitLine("}");  // WndProc
    c_.emitBlank();
