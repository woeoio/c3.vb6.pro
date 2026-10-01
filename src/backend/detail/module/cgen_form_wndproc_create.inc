// cgen_form_wndproc_create.inc - emitFormFramework 函数体片段：WndProc 前半 —— WM_CREATE / 延迟 Form_Load / WM_COMMAND / WithEvents 派发
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 598~921 行，逐行未改。

    // WndProc
    c_.emitLine("// === " + formName + " WndProc ===");
    c_.emitLine("LRESULT CALLBACK " + wndProc + "(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {");
    c_.indent();
    c_.emitLine("switch (msg) {");

    // Fix 143d: WM_PAINT — 无窗口 (windowless) OCX 控件不会自己收 WM_PAINT,
    // 必须由宿主在 BeginPaint/EndPaint 之间调用 IViewObject::Draw 让控件绘制.
    // 不生成这个 case 时, 第三方 OCX (如 NewTab01.ocx) 激活成功但窗体上空白.
    c_.indent();
    c_.emitLine("case WM_PAINT: {");
    c_.indent();
    c_.emitLine("PAINTSTRUCT _vb6_ps;");
    c_.emitLine("HDC _vb6_hdc = BeginPaint(hwnd, &_vb6_ps);");
    if (std::getenv("C3_OCX_NO_PAINT") == nullptr) {
        c_.emitLine("vb6_OcxHost_PaintAll((void*)hwnd, (void*)_vb6_hdc);  /* Fix 143d: 无窗口/窗口化 OCX 绘制 */");
    }
    // Form_Paint 事件 (若有) —— P20-43: 调用名取**符号真名**。此前写死
    // formName + "_Paint", 而用户写的是 `Form_Paint` → 定义是 vb6_Form_Paint,
    // 调用却发 vb6_<窗体名>_Paint → LNK2019 (与 Activate/DblClick 同一处缺陷)。
    {
        auto* paintSym43 = symTab_.lookup(formName + "_Paint");
        if (!paintSym43) paintSym43 = symTab_.lookup("Form_Paint");
        if (paintSym43) {
            std::string paintFn43 = cProcName(paintSym43->name, AccessLevel::Private);
            c_.emitLine("{ extern void " + paintFn43 + "(); " + paintFn43 + "(); }");
        }
    }
    c_.emitLine("EndPaint(hwnd, &_vb6_ps);");
    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");
    c_.dedent();

    // WM_CREATE: 创建控件
    c_.indent();
    c_.emitLine("case WM_CREATE: {");
    c_.indent();
    c_.emitLine("CREATESTRUCTA* cs = (CREATESTRUCTA*)lParam;");
    c_.emitLine(createFn + "((void*)hwnd, (void*)cs->hInstance);");
    c_.emitLine("vb6_Forms_Register((void*)hwnd);");

    // P24: Form.Icon from .frx - load and set window icon
    // Fix 161f: 判定用 prelude 算好的 formIconHasData_ (readPicture 非空),
    // 与数组生成同一判定, 避免"引用发了、数组没生成"的 C2065。
    if (formIconHasData_) {
        std::string varName = "vb6_frx_icon_" + cIdent(formName);
        c_.emitLine("{ void* vb6_icon = vb6_LoadIconFromMemory(" + varName + ", " + varName + "_size);");
        c_.emitLine("  if (vb6_icon) { SendMessageW(hwnd, WM_SETICON, ICON_BIG, (LPARAM)vb6_icon);");
        c_.emitLine("    SendMessageW(hwnd, WM_SETICON, ICON_SMALL, (LPARAM)vb6_icon); } }");
    }
    // P24: Form.Picture from .frx - load and store as window property (for WM_PAINT)
    if (formPicHasData_) {
        std::string varName = "vb6_frx_bgpic_" + cIdent(formName);
        c_.emitLine("{ void* vb6_bgpic = vb6_LoadPictureFromMemory(" + varName + ", " + varName + "_size);");
        c_.emitLine("  if (vb6_bgpic) SetPropW(hwnd, L\"VB6_Picture\", (HANDLE)vb6_bgpic); }");
    }

    // P20-40: 从.frm属性初始化Form属性
    {
        auto kpIt = frmDesc.formControl.properties.find("KeyPreview");
        if (kpIt != frmDesc.formControl.properties.end() && kpIt->second.intValue != 0) {
            c_.emitLine("vb6_SetKeyPreview((void*)hwnd, -1);");
        }
        auto cbIt = frmDesc.formControl.properties.find("ControlBox");
        if (cbIt != frmDesc.formControl.properties.end() && cbIt->second.intValue == 0) {
            c_.emitLine("vb6_SetControlBox((void*)hwnd, 0);");
        }
        auto mbIt = frmDesc.formControl.properties.find("MaxButton");
        if (mbIt != frmDesc.formControl.properties.end() && mbIt->second.intValue == 0) {
            c_.emitLine("vb6_SetMaxButton((void*)hwnd, 0);");
        }
        auto mnIt = frmDesc.formControl.properties.find("MinButton");
        if (mnIt != frmDesc.formControl.properties.end() && mnIt->second.intValue == 0) {
            c_.emitLine("vb6_SetMinButton((void*)hwnd, 0);");
        }
        auto wsIt = frmDesc.formControl.properties.find("WindowState");
        if (wsIt != frmDesc.formControl.properties.end() && wsIt->second.intValue != 0) {
            c_.emitLine("vb6_SetWindowState((void*)hwnd, " + std::to_string((int)wsIt->second.intValue) + ");");
        }
    }

    // M22-fix: 在Form_Load之前先赋值vb6_hwnd_<Form>，这样Form_Load中
    // Me.Caption / Me.Text 等才能通过vb6_hwnd_引用到正确的HWND
    // CreateFormWindow还未返回，vb6_hwnd_尚未被赋值
    c_.emitLine("vb6_hwnd_" + cIdent(formName) + " = (void*)hwnd;  /* early assign for Form_Load */");

    // Fix 058: Form_Load 用 PostMessage 延迟执行, 避免在 WM_CREATE 中
    // 执行 COM 操作 (如 Controls.Add) 导致消息死锁
    // 调用VB6 Form_Load事件 (仅当存在时调用)
    std::string formLoadFn = cProcName("Form_Load", AccessLevel::Private);
    auto* formLoadSym = symTab_.lookup("Form_Load");
    if (formLoadSym) {
        c_.emitLine("PostMessageW(hwnd, 0x7FF0, 0, 0);  /* defer Form_Load */");
    }
    // P24-Timer: 在WM_CREATE中注册Timer (此时hwnd可用)
    for (const auto& ctrl : frmDesc.formControl.children) {
        if (ctrl.controlType == FrmControlType::Timer) {
            int interval = 1000;
            auto pIt = ctrl.properties.find("Interval");
            if (pIt != ctrl.properties.end()) interval = (int)pIt->second.intValue;
            int enabled = 1;
            pIt = ctrl.properties.find("Enabled");
            if (pIt != ctrl.properties.end()) enabled = (int)pIt->second.intValue;
            std::string timerFn = cProcName(ctrl.controlName + "_Timer", AccessLevel::Private);
            // C29-T: 改之前"设计期 Enabled=0 就整段不发"，于是运行期 `Timer1.Enabled = True`
            // 永远起不来（实测 tick 数 0）。现在只要有事件处理器就**照挂**，把设计期那个
            // Enabled 当参数交给 RTL 决定起不起；key 传 Timer 控件自己的不可见句柄，
            // 运行期的 Enabled/Interval 才找得回这一格。没处理器就没得派发，跳过（否则链接期
            // 会引用一个不存在的 vb6_<名>_Timer）。
            if (symTab_.lookup(ctrl.controlName + "_Timer")) {
                c_.emitLine("{ extern void " + timerFn + "(); vb6_TimerAttach((void*)hwnd, vb6_hwnd_"
                            + cIdent(ctrl.controlName) + ", " + std::to_string(interval) + ", (void*)"
                            + timerFn + ", " + (enabled ? "-1" : "0") + "); }");
            }
        }
    }
    // 账 #157: VB6 显示窗体时把焦点交给 TabIndex 最小那枚。目标在发码期算好（运行期排不出序 ——
    // 设计期的 TabIndex 从没下发到窗口），存到窗体句柄上，由 vb6_ShowForm 在激活之后应用一次。
    // 放在 Timer 那一趟之后：此时每枚控件的句柄全局已经赋完了。
    emitFormInitialFocus(frmDesc.formControl);
    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");

    // M12-FIX: Form.Picture background via WM_ERASEBKGND (single display at 0,0)
    {
        auto picIt = frmDesc.formControl.properties.find("Picture");
        if (picIt != frmDesc.formControl.properties.end() && picIt->second.type == FrmValueType::FrxReference && frxLoaded) {
            c_.emitLine("case WM_ERASEBKGND: {");
            c_.indent();
            c_.emitLine("HANDLE vb6_bg = GetPropW(hwnd, L\"VB6_Picture\");");
            c_.emitLine("if (vb6_bg && GetObjectType((HGDIOBJ)vb6_bg) == OBJ_BITMAP) {");
            c_.indent();
            c_.emitLine("HDC hdc = (HDC)(WPARAM)wParam;");
            c_.emitLine("RECT rc; GetClientRect(hwnd, &rc);");
            c_.emitLine("HBRUSH hBr = (HBRUSH)(COLOR_BTNFACE+1); FillRect(hdc, &rc, hBr);");
            c_.emitLine("HDC memDC = CreateCompatibleDC(hdc);");
            c_.emitLine("HBITMAP hBmp = (HBITMAP)vb6_bg;");
            c_.emitLine("BITMAP bm; GetObjectW(hBmp, sizeof(bm), &bm);");
            c_.emitLine("HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, hBmp);");
            c_.emitLine("BitBlt(hdc, 0, 0, bm.bmWidth, bm.bmHeight, memDC, 0, 0, SRCCOPY);");
            c_.emitLine("SelectObject(memDC, oldBmp);");
            c_.emitLine("DeleteDC(memDC);");
            c_.emitLine("return 1;");
            c_.dedent();
            c_.emitLine("}");
            c_.emitLine("break;");
            c_.dedent();
            c_.emitLine("}");
        }
    }

    // Fix 187: STATIC 类子控件 (picBackColor 等色块/Label) 通过 WM_CTLCOLORSTATIC
    // 向父窗要底色/前景刷 — 此前主窗体 WndProc 不处理, vb6_SetControlBackColor
    // 写的 VB6_BackColor 属性无人消费 → 颜色对话框 OK 后色块不变色
    // (Test.exe "颜色没有写入到颜色对话左边的控件里")。RTL 统一入口
    // vb6_ApplyCtlColorStatic 仅在子控件显式 Set 过 VB6_BackColor 时接管,
    // 其余情形 break 到 DefWindowProcW 维持原生外观。
    c_.emitLine("case WM_CTLCOLORSTATIC: {");
    c_.indent();
    c_.emitLine("HWND vb6_clrchld = (HWND)lParam;");
    c_.emitLine("if (vb6_clrchld && GetPropW(vb6_clrchld, L\"VB6_BackColorSet\")) {");
    c_.indent();
    c_.emitLine("LRESULT vb6_br187 = vb6_ApplyCtlColorStatic((HDC)wParam, vb6_clrchld);");
    c_.emitLine("if (vb6_br187) return vb6_br187;");
    c_.dedent();
    c_.emitLine("}");
    c_.emitLine("break;");
    c_.dedent();
    c_.emitLine("}");

    // Fix 162f-extlist: WM_CTLCOLORBTN —— Button 类子控件 (CheckBox/OptionButton/
    // Frame) 经这条消息向**父窗**要绘制刷子。
    //   CheckBox/OptionButton → 透明 (VB6: 文字直接坐在父窗灰底上, 不遮相邻框线);
    //   Frame(BS_GROUPBOX)    → **父窗底色的实心刷** —— 经典 groupbox 标题先
    //     FillRect 再 DrawText, 空刷会让标题后露一条**白块** (Fix 162f 实测)。
    // 统一走 RTL vb6_CtlColorBtnBrush (按子控件样式分流 + 实心刷按控件缓存)。
    c_.emitLine("case WM_CTLCOLORBTN:");
    c_.indent();
    c_.emitLine("return vb6_CtlColorBtnBrush((HWND)lParam, hwnd);  /* Fix 162f: 按钮类背景刷 */");
    c_.dedent();

    // Fix 058: 处理延迟的 Form_Load (WM_USER+0x100 = 0x7FF0)
    if (formLoadSym) {
        c_.emitLine("case 0x7FF0: {");
        c_.indent();
        // P20-43: Form_Initialize —— VB6 语义是"窗体实例创建时触发一次, 早于 Load"。
        // 我们的 Load 是 PostMessage 延迟的, 所以挂在同一条 0x7FF0 里、Load 之前,
        // 用 static 标志保证整个进程只走一次 (0x7FF0 理论上只会到一次, 保险起见)。
        {
            auto* initSym43 = symTab_.lookup(formName + "_Initialize");
            if (!initSym43) initSym43 = symTab_.lookup("Form_Initialize");
            if (initSym43) {
                std::string initFn = cProcName(initSym43->name, AccessLevel::Private);
                c_.emitLine("static int vb6_formInited_" + cIdent(formName) + " = 0;");
                c_.emitLine("if (!vb6_formInited_" + cIdent(formName) + ") {");
                c_.indent();
                c_.emitLine("vb6_formInited_" + cIdent(formName) + " = 1;");
                c_.emitLine("{ extern void " + initFn + "(); " + initFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        }
        c_.emitLine("{ /* Form_Load (deferred from WM_CREATE) */ extern void " + formLoadFn + "(); " + formLoadFn + "(); }");
        // Fix 112: Load 完成后再触发首次 Form_Resize (与 VB6 的 Load→Resize 顺序一致)
        c_.emitLine("vb6_formLoaded_" + cIdent(formName) + " = 1;");
        {
            std::string resizeFnLoad = cProcName("Form_Resize", AccessLevel::Private);
            if (symTab_.lookup(formName + "_Resize") || symTab_.lookup("Form_Resize")) {
                c_.emitLine("{ extern void " + resizeFnLoad + "(); " + resizeFnLoad + "(); }");
            }
        }
        c_.emitLine("break;");
        c_.dedent();
        c_.emitLine("}");
    }

    // WM_COMMAND: 按钮点击等 (P7.6: 支持控件数组Index参数)
    c_.emitLine("case WM_COMMAND: {");
    c_.indent();
    c_.emitLine("if (vb6_formLoading_" + cIdent(formName) + ") break;  /* block events during form init */");
    c_.emitLine("int id = LOWORD(wParam);");
    c_.emitLine("int code = HIWORD(wParam);");

    // 为每个CommandButton/CheckBox/OptionButton生成WM_COMMAND处理
    // Fix 142: 抽为 lambda, 顶层控件(ID 100+)与容器内子控件(ID 200+)共用同一派发逻辑
    auto emitCommandDispatch142 = [&](const FrmControl& ctrl, int cid) {
        ctrlId = cid;
        std::string ctrlNameLower = ctrl.controlName;
        std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
        bool isArrayCtrl = knownControlArrays_.count(ctrlNameLower) > 0;

        if (ctrl.controlType == FrmControlType::CommandButton ||
            ctrl.controlType == FrmControlType::CheckBox ||
            ctrl.controlType == FrmControlType::OptionButton) {
            if (symTab_.lookup(ctrl.controlName + "_Click")) {
                std::string clickFn = cProcName(ctrl.controlName + "_Click", AccessLevel::Private);
                // 账 #158：**只认 BN_CLICKED=0** —— 按钮族现在挂了 `BS_NOTIFY`，同一个 id 还会
                // 携着 `code=6/7`（焦点变化）进来，不过滤就把一次焦点移动算成一次点击。
                // 账 #171（C29-FS-h）：**单选钮这一型光按 code 筛不掉** —— 实测焦点进入组内另一枚时，
                // 它替父窗发的就是 `BN_CLICKED`（号 0），而那一刻它自己并没被勾上（两头证人：`Value`
                // 与 OS 的 `BM_GETCHECK` 都答「没勾」）。VB6 的 `Click` 只在用户真选了这枚时发生
                // ⇒ 只在 OptionButton 的 arm 上多问一句来路勾没勾；口径写在 RTL 的
                // `vb6_RadioClickCounts`（唯一一处），别的型一律放行、条件也不多发（爆炸半径只管单选钮）。
                std::string guard171 = (ctrl.controlType == FrmControlType::OptionButton)
                    ? " && vb6_RadioClickCounts((void*)lParam)" : std::string();
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 0" + guard171 + ") {");
                c_.indent();
                if (isArrayCtrl) {
                    // 控件数组: 事件处理带Index参数
                    std::string idxVar = "vb6_idx_" + std::to_string(ctrl.index >= 0 ? ctrl.index : 0);
                    c_.emitLine("{ int16_t " + idxVar + " = " + std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + "; extern void " + clickFn + "(int16_t*); " + clickFn + "(&" + idxVar + "); }");
                } else {
                    c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                }
                c_.dedent();
                c_.emitLine("}");
            }
        }
        else if (ctrl.controlType == FrmControlType::TextBox) {
            // P14.4.1: EN_CHANGE (code=0x0300=768) -> _Change() (仅当处理器存在时)
            if (symTab_.lookup(ctrl.controlName + "_Change")) {
                std::string changeFn = cProcName(ctrl.controlName + "_Change", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 768) {");
                c_.indent();
                c_.emitLine("{ extern void " + changeFn + "(); " + changeFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        }
        else if (ctrl.controlType == FrmControlType::RichTextBox) {
            // C29-RT-d: RichTextBox 的 Change **不在 768 这一格** —— 上面那条是 EDIT 控件的
            // EN_CHANGE，Msftedit 一次都不发它（探针六份读数里 768 出现 0 次）。它发的是
            // EN_UPDATE(0x0400=1024)，而且**不受事件掩码管**（掩码全 0 也照发），所以这条
            // 不需要 vb6_RTB_Init 那步配合（对照：同一条 Init 里的 ENM_SELCHANGE 少一句就哑）。
            // 读数 .build\rtdprobe5_c*_s*.txt；程序化写 Text 也算"变了"，与 VB6 一致。
            // ⚠ 但**时机**与 VB6 不同：这一条要等重绘才落地 —— `Text = x` 之后得过一次消息循环
            //   才收得到（夹具里那句 DoEvents 因此是判据的一部分，不 DoEvents 计数还是旧的）。
            if (symTab_.lookup(ctrl.controlName + "_Change")) {
                std::string rtbChgFn = cProcName(ctrl.controlName + "_Change", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1024) {");
                c_.indent();
                c_.emitLine("{ extern void " + rtbChgFn + "(); " + rtbChgFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // C29-RT-d: SelChange 的**第二条通道**（与下面 WM_NOTIFY 那一条同一个事件）。
            // 同一份产物、同一个夹具连跑三次：选区通知 1 次走 WM_NOTIFY/1794、2 次走
            // WM_COMMAND/1815（0x0717，SDK 头里没有这枚码）；裸码读数 .build\rtdmin_out\tr_*.txt。
            // 两条 arm 都接上，事件才谈得上"会来"；一条事件双发会让计数变 2 当场翻红，所以
            // 双发不可能被这两条 arm 一起放过。判据用 vb6_RTB_SimNotify 各钉一条。
            if (symTab_.lookup(ctrl.controlName + "_SelChange")) {
                std::string rtbSelFn = cProcName(ctrl.controlName + "_SelChange", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1815) {");
                c_.indent();
                c_.emitLine("{ extern void " + rtbSelFn + "(); " + rtbSelFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        } else if (ctrl.controlType == FrmControlType::ListBox) {
            // P14.4.1: LBN_SELCHANGE (code=1) -> _Click() (仅当处理器存在时)
            if (symTab_.lookup(ctrl.controlName + "_Click")) {
                std::string clickFn = cProcName(ctrl.controlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1) {");
                c_.indent();
                c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // LBN_DBLCLK (code=2) -> _DblClick() (仅当处理器存在时)
            if (symTab_.lookup(ctrl.controlName + "_DblClick")) {
                std::string dblClickFn = cProcName(ctrl.controlName + "_DblClick", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 2) {");
                c_.indent();
                c_.emitLine("{ extern void " + dblClickFn + "(); " + dblClickFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        } else if (ctrl.controlType == FrmControlType::DriveListBox) {
            // C29-1b: 换盘 = CBN_SELCHANGE(1) → Drive1_Change；VB6 同一条通知也发 Click，
            // 所以两个处理器各自挂一次（有哪个就发哪个，缺一不报错）。
            if (symTab_.lookup(ctrl.controlName + "_Change")) {
                std::string drvChgFn = cProcName(ctrl.controlName + "_Change", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1) {");
                c_.indent();
                c_.emitLine("{ extern void " + drvChgFn + "(); " + drvChgFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            if (symTab_.lookup(ctrl.controlName + "_Click")) {
                std::string drvClkFn = cProcName(ctrl.controlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1) {");
                c_.indent();
                c_.emitLine("{ extern void " + drvClkFn + "(); " + drvClkFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        } else if (ctrl.controlType == FrmControlType::DirListBox) {
            // C29-1b: 双击列表项 = 下钻一层。VB6 的 Dir1_Change 是"Path 真的变了"才发,
            // 所以先让 RTL 试下钻、返回 1 才调处理器 —— 双击卷标行/空项不惊动事件。
            if (symTab_.lookup(ctrl.controlName + "_Change")) {
                std::string dirChgFn = cProcName(ctrl.controlName + "_Change", AccessLevel::Private);
                std::string dirHwnd = ctrlHwndExprForInit(ctrl);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 2) {");
                c_.indent();
                c_.emitLine("if (vb6_DirListBoxDescendSelected((void*)" + dirHwnd + ")) {");
                c_.indent();
                c_.emitLine("{ extern void " + dirChgFn + "(); " + dirChgFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
                c_.dedent();
                c_.emitLine("}");
            }
            if (symTab_.lookup(ctrl.controlName + "_Click")) {
                std::string dirClkFn = cProcName(ctrl.controlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1) {");
                c_.indent();
                c_.emitLine("{ extern void " + dirClkFn + "(); " + dirClkFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        } else if (ctrl.controlType == FrmControlType::FileListBox) {
            // C29-1b: 选中文件 = LBN_SELCHANGE(1) → Click；双击 = LBN_DBLCLK(2) → DblClick
            if (symTab_.lookup(ctrl.controlName + "_Click")) {
                std::string fileClkFn = cProcName(ctrl.controlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1) {");
                c_.indent();
                c_.emitLine("{ extern void " + fileClkFn + "(); " + fileClkFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            if (symTab_.lookup(ctrl.controlName + "_DblClick")) {
                std::string fileDblFn = cProcName(ctrl.controlName + "_DblClick", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 2) {");
                c_.indent();
                c_.emitLine("{ extern void " + fileDblFn + "(); " + fileDblFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        } else if (ctrl.controlType == FrmControlType::ComboBox) {
            // P14.4.1: CBN_SELCHANGE (code=1) -> _Click() (仅当处理器存在时)
            if (symTab_.lookup(ctrl.controlName + "_Click")) {
                std::string clickFn = cProcName(ctrl.controlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 1) {");
                c_.indent();
                c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // CBN_EDITCHANGE (code=5) -> _Change() (仅当处理器存在时)
            if (symTab_.lookup(ctrl.controlName + "_Change")) {
                std::string changeFn = cProcName(ctrl.controlName + "_Change", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 5) {");
                c_.indent();
                c_.emitLine("{ extern void " + changeFn + "(); " + changeFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // CBN_DROPDOWN (code=7) -> _DropDown() (仅当处理器存在时)
            if (symTab_.lookup(ctrl.controlName + "_DropDown")) {
                std::string dropDownFn = cProcName(ctrl.controlName + "_DropDown", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 7) {");
                c_.indent();
                c_.emitLine("{ extern void " + dropDownFn + "(); " + dropDownFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // CBN_CLOSEUP (code=8) -> _CloseUp() (仅当处理器存在时)
            if (symTab_.lookup(ctrl.controlName + "_CloseUp")) {
                std::string closeUpFn = cProcName(ctrl.controlName + "_CloseUp", AccessLevel::Private);
                c_.emitLine("if (id == " + std::to_string(ctrlId) + " && code == 8) {");
                c_.indent();
                c_.emitLine("{ extern void " + closeUpFn + "(); " + closeUpFn + "(); }");
                c_.dedent();
                c_.emitLine("}");
            }
        }
    };

    ctrlId = 100;
    for (const auto& ctrl : frmDesc.formControl.children) {
        emitCommandDispatch142(ctrl, ctrlId);
        ctrlId++;
    }
    // Fix 142: 容器(Frame/PictureBox)内子控件 — ID 从 200 起, DFS 顺序与
    // cgen_form_frame_menu.inc 的创建循环完全一致 (跳过 Menu 与无 Win32 类控件).
    {
        int childCmdId142 = 200;
        auto walkCmdChildren142 = [&](auto&& self, const FrmControl& container) -> void {
            for (const auto& child : container.children) {
                if (child.controlType == FrmControlType::Menu) continue;
                const char* childClass142 = FrmParser::controlTypeToWin32Class(child.controlType);
                if (!childClass142) continue;
                emitCommandDispatch142(child, childCmdId142);
                childCmdId142++;
                if (!child.children.empty() && controlIsContainerType(child)) {   // 账 #167: 原来少认 SSTab
                    self(self, child);
                }
            }
        };
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (!ctrl.children.empty() && controlIsContainerType(ctrl)) {   // 账 #167: 原来少认 SSTab
                walkCmdChildren142(walkCmdChildren142, ctrl);
            }
        }
    }

    // C29-5c: Toolbar 的 ButtonClick —— 原生工具栏报按钮点击**不走 WM_NOTIFY**，而是按
    // 标准做法给父窗发 WM_COMMAND: LOWORD(wParam)=按钮的 idCommand (5a 写的是"槽号+1")、
    // HIWORD=0、lParam=工具栏自己的 HWND。所以按 lParam 认来源、按 id 取按钮，两条都来自
    // 同一个 idCommand ⇒ 派发与表错位会当场翻红:
    //   Sub Toolbar1_ButtonClick(ByVal Button As Button)
    // (_ButtonMenuClick 才走 WM_NOTIFY 的 TBN_DROPDOWN，见 cgen_form_wndproc_dispatch.inc。)
    for (const auto& ctrl : frmDesc.formControl.children) {
        if (ctrl.controlType != FrmControlType::Toolbar) continue;
        if (!symTab_.lookup(ctrl.controlName + "_ButtonClick")) continue;
        std::string tbFn5c = cProcName(ctrl.controlName + "_ButtonClick", AccessLevel::Private);
        std::string tbSrc5c = ctrlHwndExprForInit(ctrl);
        c_.emitLine("{");
        c_.indent();
        c_.emitLine("HWND vb6_tbSrc5c = (HWND)" + tbSrc5c + ";");
        c_.emitLine("if (code == 0 && (HWND)lParam == vb6_tbSrc5c) {");
        c_.indent();
        c_.emitLine("void* vb6_tbBtn5c = vb6_Toolbar_ButtonAt((void*)vb6_tbSrc5c, id);");
        c_.emitLine("if (vb6_tbBtn5c) { extern void " + tbFn5c + "(void*); " + tbFn5c + "(vb6_tbBtn5c); }");
        c_.dedent();
        c_.emitLine("}");
        c_.dedent();
        c_.emitLine("}");
    }

    // P16: WithEvents控件事件分发 (按HWND匹配, 在标准命名handler之后)
    // 对于每个WithEvents控件变量, 检查其HWND是否匹配lParam, 是则调用对应handler
    if (!knownWithEventsCtrlVars_.empty()) {
        c_.emitLine("if ((HWND)lParam != NULL) {");
        c_.indent();
        c_.emitLine("HWND hCtrlWE = (HWND)lParam;");
        for (const auto& we : knownWithEventsCtrlVars_) {
            std::string varLower = we.first;
            FrmControlType ctrlType = we.second;
            std::string varName = knownWithEventsCtrlOrigNames_.count(varLower) ? knownWithEventsCtrlOrigNames_[varLower] : cIdent(varLower);  // 用原始大小写变量名
            // 根据控件类型生成不同的事件分发
            if (ctrlType == FrmControlType::CommandButton ||
                ctrlType == FrmControlType::CheckBox ||
                ctrlType == FrmControlType::OptionButton) {
                // BN_CLICKED (code=0) → _Click()
                auto* clickSym = symTab_.lookup(varLower + "_click");
                if (clickSym) {
                    std::string clickFn = cProcName(clickSym->name, AccessLevel::Private);
                    c_.emitLine("if (code == 0 && " + varName + " != NULL && " + varName + " == hCtrlWE) {");
                    c_.indent();
                    c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
            } else if (ctrlType == FrmControlType::TextBox) {
                // EN_CHANGE (code=768) → _Change()
                auto* changeSym = symTab_.lookup(varLower + "_change");
                if (changeSym) {
                    std::string changeFn = cProcName(changeSym->name, AccessLevel::Private);
                    c_.emitLine("if (code == 768 && " + varName + " != NULL && " + varName + " == hCtrlWE) {");
                    c_.indent();
                    c_.emitLine("{ extern void " + changeFn + "(); " + changeFn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
            } else if (ctrlType == FrmControlType::ListBox) {
                // LBN_SELCHANGE (code=1) → _Click()
                auto* clickSym = symTab_.lookup(varLower + "_click");
                if (clickSym) {
                    std::string clickFn = cProcName(clickSym->name, AccessLevel::Private);
                    c_.emitLine("if (code == 1 && " + varName + " != NULL && " + varName + " == hCtrlWE) {");
                    c_.indent();
                    c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
                // LBN_DBLCLK (code=2) → _DblClick()
                auto* dblClickSym = symTab_.lookup(varLower + "_dblclick");
                if (dblClickSym) {
                    std::string dblClickFn = cProcName(dblClickSym->name, AccessLevel::Private);
                    c_.emitLine("if (code == 2 && " + varName + " != NULL && " + varName + " == hCtrlWE) {");
                    c_.indent();
                    c_.emitLine("{ extern void " + dblClickFn + "(); " + dblClickFn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
            } else if (ctrlType == FrmControlType::ComboBox) {
                // CBN_SELCHANGE (code=1) → _Click()
                auto* clickSym = symTab_.lookup(varLower + "_click");
                if (clickSym) {
                    std::string clickFn = cProcName(clickSym->name, AccessLevel::Private);
                    c_.emitLine("if (code == 1 && " + varName + " != NULL && " + varName + " == hCtrlWE) {");
                    c_.indent();
                    c_.emitLine("{ extern void " + clickFn + "(); " + clickFn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
                // CBN_EDITCHANGE (code=5) → _Change()
                auto* changeSym = symTab_.lookup(varLower + "_change");
                if (changeSym) {
                    std::string changeFn = cProcName(changeSym->name, AccessLevel::Private);
                    c_.emitLine("if (code == 5 && " + varName + " != NULL && " + varName + " == hCtrlWE) {");
                    c_.indent();
                    c_.emitLine("{ extern void " + changeFn + "(); " + changeFn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
            }
            // P24-Timer: WM_TIMER由WndProc分发 (不再由消息循环拦截)
            // HScrollBar/VScrollBar: 在WM_HSCROLL/WM_VSCROLL中处理(下方)
        }
        c_.dedent();
        c_.emitLine("}");  /* close if ((HWND)lParam != NULL) */
    }
    // P7.8: 菜单项点击事件派发 (Menu控件ID从1000开始)
    {
        int menuId = 1000;
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (ctrl.controlType == FrmControlType::Menu) {
                emitMenuClickDispatch(ctrl, menuId);
            }
        }
    }
