// cgen_form_wndproc_subclass.inc - emitFormFramework 函数体片段：.c 实现开头 —— Form_Unload trampoline 与需子类化控件的 WndProc
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 262~597 行，逐行未改。

    // --- .c文件: 实现 ---

    // Form_Unload trampoline (如果存在Form_Unload)
    {
        std::string formUnloadFn = cProcName("Form_Unload", AccessLevel::Private);
        auto* unloadSym = symTab_.lookup("Form_Unload");
        if (unloadSym) {
            c_.emitLine("// Form_Unload trampoline: bridge callback int(*)(void) to void(int16_t*)");
            c_.emitLine("static int " + formUnloadFn + "_trampoline(void) {");
            c_.indent();
            c_.emitLine("int16_t vb6_cancel = 0;");
            c_.emitLine("extern void " + formUnloadFn + "(int16_t*);");
            c_.emitLine(formUnloadFn + "(&vb6_cancel);");
            c_.emitLine("return (int)vb6_cancel;");
            c_.dedent();
            c_.emitLine("}");
            c_.emitBlank();
        }
    }
    // P18-F: 生成需要子类化的控件的 subclass WndProc
    // 需要子类化的条件: 控件有MouseEnter/MouseLeave/MouseDown/MouseUp/MouseMove/KeyPress/KeyDown/KeyUp事件处理器
    // 或者 PictureBox/Frame/Label 有 GotFocus/LostFocus 事件处理器(这些控件不通过WM_COMMAND发送焦点通知)
    {
        struct SubclassInfo {
            std::string ctrlName;      // 控件名(原始大小写)
            std::string hwndVar;       // HWND变量名
            bool hasGotFocus = false;
            bool hasLostFocus = false;
            bool hasMouseEnter = false;
            bool hasMouseLeave = false;
            bool hasMouseDown = false;
            bool hasMouseUp = false;
            bool hasMouseMove = false;
            bool hasKeyPress = false;
            bool hasKeyDown = false;
            bool hasKeyUp = false;
            bool hasDblClick = false;
            bool hasClick = false;     // C29-SL-d: 子类化那一条路补上的常规 Click
            bool hasMouseHover = false;
            bool hasValidate = false;
            bool hasPaint = false;   // Fix 185: 控件级 _Paint (PictureBox.Print/Cls 的落点)
            FrmControlType ctrlType;
        };
        std::vector<SubclassInfo> subclassCtrls;
        std::unordered_set<std::string> emitted;
        for (const auto& ctrl : frmDesc.formControl.children) {
            std::string ctrlNameLower = ctrl.controlName;
            std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
            if (emitted.count(ctrlNameLower)) continue;
            emitted.insert(ctrlNameLower);

            // 跳过不可见控件(Timer/Menu等)和没有Win32窗口类的控件
            if (ctrl.controlType == FrmControlType::Timer ||
                ctrl.controlType == FrmControlType::Menu ||
                ctrl.controlType == FrmControlType::CommonDialog ||
                ctrl.controlType == FrmControlType::WebBrowser) continue;

            SubclassInfo info;
            info.ctrlName = ctrl.controlName;
            info.ctrlType = ctrl.controlType;
            if (knownControlArrays_.count(ctrlNameLower)) {
                info.hwndVar = "(void*)vb6_arr_" + cIdent(ctrl.controlName) + ".hwnds[0]";
            } else {
                info.hwndVar = "vb6_hwnd_" + cIdent(ctrl.controlName);
            }

            // C29-SL-g: 这一档以前是**四类白名单**（PictureBox/Frame/Label/Image），与
            // "装不装"、"拆不拆"那两份一起把剩下的所有窗口态控件 —— Slider 首当其冲，还有
            // ListView / TreeView / DTPicker / MonthView / RichTextBox / 两个滚动条 /
            // 文件系统三件套 —— 的 `_GotFocus` / `_LostFocus` 留在"编得过、永不触发"那一侧。
            // 改成按"有没有另一条原生来源"排除：那六类的焦点通知由父窗的 `WM_COMMAND` 码供
            // （EN_/CBN_/LBN_/BN_，见 controlFocusFromNativeNotify 的定义处），其余一律走
            // 控件自己的 `WM_SETFOCUS` / `WM_KILLFOCUS` —— 同一次焦点变化只会有一处发。
            if (!controlFocusFromNativeNotify(ctrl.controlType)) {
                info.hasGotFocus = symTab_.lookup(ctrl.controlName + "_GotFocus") != nullptr;
                info.hasLostFocus = symTab_.lookup(ctrl.controlName + "_LostFocus") != nullptr;
            }
            // P20-12: Validate event (all control types)
            info.hasValidate = symTab_.lookup(ctrl.controlName + "_Validate") != nullptr;

            // MouseEnter/MouseLeave 总是需要子类化(TrackMouseEvent)
            info.hasMouseEnter = symTab_.lookup(ctrl.controlName + "_MouseEnter") != nullptr;
            info.hasMouseLeave = symTab_.lookup(ctrl.controlName + "_MouseLeave") != nullptr;
            // VB6中MouseHover对应WM_MOUSEHOVER, 较少见, 暂不实现

            // 控件级鼠标事件(所有控件都可以有)需要子类化
            info.hasMouseDown = symTab_.lookup(ctrl.controlName + "_MouseDown") != nullptr;
            info.hasMouseUp = symTab_.lookup(ctrl.controlName + "_MouseUp") != nullptr;
            info.hasMouseMove = symTab_.lookup(ctrl.controlName + "_MouseMove") != nullptr;

            // 控件级键盘事件需要子类化
            info.hasKeyPress = symTab_.lookup(ctrl.controlName + "_KeyPress") != nullptr;
            info.hasKeyDown = symTab_.lookup(ctrl.controlName + "_KeyDown") != nullptr;
            info.hasKeyUp = symTab_.lookup(ctrl.controlName + "_KeyUp") != nullptr;
            info.hasDblClick = symTab_.lookup(ctrl.controlName + "_DblClick") != nullptr;
            // C29-SL-d: 只有**没有原生 Click 来源**的类型才由这一档补上（排除表与逐条来源
            // 见 controlClickFromNativeNotify 的定义处），否则同一次点击会双发。
            info.hasClick = !controlClickFromNativeNotify(ctrl.controlType) &&
                            symTab_.lookup(ctrl.controlName + "_Click") != nullptr;
            info.hasMouseHover = symTab_.lookup(ctrl.controlName + "_MouseHover") != nullptr;
            // Fix 185: 控件级 _Paint。原先整个控件家族都没有 WM_PAINT 转发，
            // MainForm.c 里的 vb6_MainForm_Picture2_Paint 是死代码。只接 PictureBox
            // （VB6 里有绘制表面的控件），本 demo 度量该事件仅此一处。
            info.hasPaint = (ctrl.controlType == FrmControlType::PictureBox) &&
                            (symTab_.lookup(ctrl.controlName + "_Paint") != nullptr);

            // 只要有任一子类化事件需求就加入列表
            // C29-SL-n（账 #141）: 这一趟原本自己按 info.hasXxx 汇一遍，与"装/拆"那两份各抄一遍
            // ⇒ 抄漏的那形（`_DblClick`、`_Paint`）就变成"arm 发了、没人 install"。三份共用
            // controlNeedsSubclass 之后，info.hasXxx 只管"发哪几条 arm"，"装不装"只有一处。
            if (controlNeedsSubclass(ctrl)) {
                subclassCtrls.push_back(std::move(info));
            }
        }

        // P0-5修复: 扫描容器内的子控件（账 #167: 这份「谁是容器」原本在这里手抄，
        // 抄的是 Frame + PictureBox 两型 ⇒ SSTab 的孩子没人发子类化定义；换成唯一出口）
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (!controlIsContainerType(ctrl)) continue;
            // Fix 160w: PictureBox 也可作为容器 (MainForm.frm Picture2 容器内
            // 有设置了 GotFocus/MouseDown 的子控件; 原来只扫 Frame → 子控件已
            // 在 cgen_form_frame_menu.inc 注册 install 但这里未生成子类化定义/
            // 转发 → MainForm.c:281 C2065 vb6_ctrl_subproc_Picture2).
            for (const auto& child : ctrl.children) {
                std::string childNameLower = child.controlName;
                std::transform(childNameLower.begin(), childNameLower.end(), childNameLower.begin(), ::tolower);
                if (emitted.count(childNameLower)) continue;
                emitted.insert(childNameLower);
                if (child.controlType == FrmControlType::Timer ||
                    child.controlType == FrmControlType::Menu ||
                    child.controlType == FrmControlType::CommonDialog ||
                    child.controlType == FrmControlType::WebBrowser) continue;
                SubclassInfo info;
                info.ctrlName = child.controlName;
                info.ctrlType = child.controlType;
                info.hwndVar = "vb6_hwnd_" + cIdent(child.controlName);
                // C29-SL-g: 与顶层那一趟同一口径（按"有没有另一条原生来源"排除）
                if (!controlFocusFromNativeNotify(child.controlType)) {
                    info.hasGotFocus = symTab_.lookup(child.controlName + "_GotFocus") != nullptr;
                    info.hasLostFocus = symTab_.lookup(child.controlName + "_LostFocus") != nullptr;
                }
                // P20-12: Validate event (all control types)
                info.hasValidate = symTab_.lookup(child.controlName + "_Validate") != nullptr;
                info.hasMouseEnter = symTab_.lookup(child.controlName + "_MouseEnter") != nullptr;
                info.hasMouseLeave = symTab_.lookup(child.controlName + "_MouseLeave") != nullptr;
                info.hasMouseDown = symTab_.lookup(child.controlName + "_MouseDown") != nullptr;
                info.hasMouseUp = symTab_.lookup(child.controlName + "_MouseUp") != nullptr;
                info.hasMouseMove = symTab_.lookup(child.controlName + "_MouseMove") != nullptr;
                info.hasKeyPress = symTab_.lookup(child.controlName + "_KeyPress") != nullptr;
                info.hasKeyDown = symTab_.lookup(child.controlName + "_KeyDown") != nullptr;
                info.hasKeyUp = symTab_.lookup(child.controlName + "_KeyUp") != nullptr;
                info.hasDblClick = symTab_.lookup(child.controlName + "_DblClick") != nullptr;
                // C29-SL-d: 同顶层那一趟的口径（排除表见 controlClickFromNativeNotify）。
                info.hasClick = !controlClickFromNativeNotify(child.controlType) &&
                                symTab_.lookup(child.controlName + "_Click") != nullptr;
                info.hasMouseHover = symTab_.lookup(child.controlName + "_MouseHover") != nullptr;
                info.hasPaint = (child.controlType == FrmControlType::PictureBox) &&
                                (symTab_.lookup(child.controlName + "_Paint") != nullptr);  // Fix 185
                // C29-SL-n: 容器子控件这一趟与上面顶层那趟共用同一份判据（原来是第二次抄写）。
                if (controlNeedsSubclass(child)) {
                    subclassCtrls.push_back(std::move(info));
                }
            }
        }
        // 为每个需要子类化的控件生成 subclass WndProc
        for (const auto& info : subclassCtrls) {
            std::string subProcName = "vb6_ctrl_subproc_" + cIdent(info.ctrlName);
            c_.emitLine("// P18-F: Subclass WndProc for " + info.ctrlName);
            c_.emitLine("static LRESULT CALLBACK " + subProcName + "(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {");
            c_.indent();

            // GotFocus (WM_SETFOCUS)
            if (info.hasGotFocus) {
                std::string fn = cProcName(info.ctrlName + "_GotFocus", AccessLevel::Private);
                c_.emitLine("if (msg == WM_SETFOCUS) { extern void " + fn + "(); " + fn + "(); }");
            }
            // LostFocus (WM_KILLFOCUS)
            if (info.hasLostFocus) {
                std::string fn = cProcName(info.ctrlName + "_LostFocus", AccessLevel::Private);
                if (info.hasValidate) {
                c_.emitLine("if (msg == WM_KILLFOCUS && !GetPropW(hwnd, L\"VB6_ValidateCancel\")) { extern void " + fn + "(); " + fn + "(); }");
            } else {
                c_.emitLine("if (msg == WM_KILLFOCUS) { extern void " + fn + "(); " + fn + "(); }");
            }
            }
            // P20-12: Validate event (WM_KILLFOCUS with CausesValidation + Cancel)
            // Validate fires BEFORE LostFocus when the gaining focus control has CausesValidation=True
            if (info.hasValidate) {
                std::string fn = cProcName(info.ctrlName + "_Validate", AccessLevel::Private);
                c_.emitLine("// Validate(Cancel As Boolean) - fires before LostFocus");
                c_.emitLine("if (msg == WM_KILLFOCUS) {");
                c_.indent();
                c_.emitLine("HWND vb6_gaining = (HWND)wp;");
                c_.emitLine("if (vb6_GetCausesValidation((void*)vb6_gaining)) {");
                c_.indent();
                c_.emitLine("int16_t vb6_vcancel = 0;");
                c_.emitLine("extern void " + fn + "(int16_t*);");
                c_.emitLine(fn + "(&vb6_vcancel);");
                c_.emitLine("if (vb6_vcancel) {");
                c_.indent();
                c_.emitLine("// Cancel=True: prevent focus transfer, restore focus");
                c_.emitLine("SetPropW(hwnd, L\"VB6_ValidateCancel\", (HANDLE)1);");
                c_.emitLine("PostMessage(hwnd, 0x7FFF, 0, 0); /* WM_VB6_RESTOREFOCUS */");
                c_.dedent();
                c_.emitLine("}");
                c_.dedent();
                c_.emitLine("}");
                c_.dedent();
                c_.emitLine("}");
            }

            // MouseEnter/MouseLeave via TrackMouseEvent
            if (info.hasMouseEnter || info.hasMouseLeave || info.hasMouseHover) {
                c_.emitLine("if (msg == WM_MOUSEMOVE) {");
                c_.indent();
                c_.emitLine("if (!GetPropW(hwnd, L\"VB6_MouseTracked\")) {");
                c_.indent();
                c_.emitLine("SetPropW(hwnd, L\"VB6_MouseTracked\", (HANDLE)1);");
                c_.emitLine("vb6_StartMouseTracking((void*)hwnd);");
                if (info.hasMouseEnter) {
                    std::string fn = cProcName(info.ctrlName + "_MouseEnter", AccessLevel::Private);
                    c_.emitLine("{ extern void " + fn + "(); " + fn + "(); }");
                }
                c_.dedent();
                c_.emitLine("}");
                // 如果同时有MouseMove事件，也在这里分发
                if (info.hasMouseMove) {
                    std::string fn = cProcName(info.ctrlName + "_MouseMove", AccessLevel::Private);
                    c_.emitLine("{ int16_t vb6_button = 0; int16_t vb6_shift = 0;");
                    c_.emitLine("if (wp & MK_LBUTTON) vb6_button |= 1; if (wp & MK_RBUTTON) vb6_button |= 2; if (wp & MK_MBUTTON) vb6_button |= 4;");
                    c_.emitLine("if (wp & MK_SHIFT) vb6_shift |= 1; if (wp & MK_CONTROL) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
                    c_.emitLine("float vb6_x = (float)(int16_t)LOWORD(lp); float vb6_y = (float)(int16_t)HIWORD(lp);");
                    c_.emitLine("extern void " + fn + "(int16_t*, int16_t*, float*, float*); " + fn + "(&vb6_button, &vb6_shift, &vb6_x, &vb6_y); }");
                }
                c_.dedent();
                c_.emitLine("}");
                c_.emitLine("if (msg == WM_MOUSELEAVE) {");
                c_.indent();
                c_.emitLine("RemovePropW(hwnd, L\"VB6_MouseTracked\");");
                if (info.hasMouseLeave) {
                    std::string fn = cProcName(info.ctrlName + "_MouseLeave", AccessLevel::Private);
                    c_.emitLine("{ extern void " + fn + "(); " + fn + "(); }");
                }
                c_.dedent();
                c_.emitLine("}");
                if (info.hasMouseHover) {
                    c_.emitLine("if (msg == WM_MOUSEHOVER) {");
                    c_.indent();
                    std::string hfn = cProcName(info.ctrlName + "_MouseHover", AccessLevel::Private);
                    c_.emitLine("{ extern void " + hfn + "(); " + hfn + "(); }");
                    c_.dedent();
                    c_.emitLine("}");
                }
            } else if (info.hasMouseMove) {
                // 只有MouseMove没有MouseEnter/Leave
                c_.emitLine("if (msg == WM_MOUSEMOVE) {");
                c_.indent();
                std::string fn = cProcName(info.ctrlName + "_MouseMove", AccessLevel::Private);
                c_.emitLine("{ int16_t vb6_button = 0; int16_t vb6_shift = 0;");
                c_.emitLine("if (wp & MK_LBUTTON) vb6_button |= 1; if (wp & MK_RBUTTON) vb6_button |= 2; if (wp & MK_MBUTTON) vb6_button |= 4;");
                c_.emitLine("if (wp & MK_SHIFT) vb6_shift |= 1; if (wp & MK_CONTROL) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
                c_.emitLine("float vb6_x = (float)(int16_t)LOWORD(lp); float vb6_y = (float)(int16_t)HIWORD(lp);");
                c_.emitLine("extern void " + fn + "(int16_t*, int16_t*, float*, float*); " + fn + "(&vb6_button, &vb6_shift, &vb6_x, &vb6_y); }");
                c_.dedent();
                c_.emitLine("}");
            }

            // MouseDown
            if (info.hasMouseDown) {
                std::string fn = cProcName(info.ctrlName + "_MouseDown", AccessLevel::Private);
                c_.emitLine("if (msg == WM_LBUTTONDOWN || msg == WM_RBUTTONDOWN || msg == WM_MBUTTONDOWN) {");
                c_.indent();
                c_.emitLine("{ int16_t vb6_button = 0; int16_t vb6_shift = 0;");
                c_.emitLine("if (msg == WM_LBUTTONDOWN) vb6_button = 1;");
                c_.emitLine("else if (msg == WM_RBUTTONDOWN) vb6_button = 2;");
                c_.emitLine("else if (msg == WM_MBUTTONDOWN) vb6_button = 4;");
                c_.emitLine("if (wp & MK_SHIFT) vb6_shift |= 1; if (wp & MK_CONTROL) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
                c_.emitLine("float vb6_x = (float)(int16_t)LOWORD(lp); float vb6_y = (float)(int16_t)HIWORD(lp);");
                c_.emitLine("extern void " + fn + "(int16_t*, int16_t*, float*, float*); " + fn + "(&vb6_button, &vb6_shift, &vb6_x, &vb6_y); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // MouseUp
            if (info.hasMouseUp) {
                std::string fn = cProcName(info.ctrlName + "_MouseUp", AccessLevel::Private);
                c_.emitLine("if (msg == WM_LBUTTONUP || msg == WM_RBUTTONUP || msg == WM_MBUTTONUP) {");
                c_.indent();
                c_.emitLine("{ int16_t vb6_button = 0; int16_t vb6_shift = 0;");
                c_.emitLine("if (msg == WM_LBUTTONUP) vb6_button = 1;");
                c_.emitLine("else if (msg == WM_RBUTTONUP) vb6_button = 2;");
                c_.emitLine("else if (msg == WM_MBUTTONUP) vb6_button = 4;");
                c_.emitLine("if (wp & MK_SHIFT) vb6_shift |= 1; if (wp & MK_CONTROL) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
                c_.emitLine("float vb6_x = (float)(int16_t)LOWORD(lp); float vb6_y = (float)(int16_t)HIWORD(lp);");
                c_.emitLine("extern void " + fn + "(int16_t*, int16_t*, float*, float*); " + fn + "(&vb6_button, &vb6_shift, &vb6_x, &vb6_y); }");
                c_.dedent();
                c_.emitLine("}");
            }

            // KeyPress (WM_CHAR)
            if (info.hasKeyPress) {
                std::string fn = cProcName(info.ctrlName + "_KeyPress", AccessLevel::Private);
                c_.emitLine("if (msg == WM_CHAR) {");
                c_.indent();
                c_.emitLine("{ int16_t vb6_keyascii = (int16_t)(unsigned char)wp;");
                c_.emitLine("extern void " + fn + "(int16_t*); " + fn + "(&vb6_keyascii); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // KeyDown (WM_KEYDOWN)
            if (info.hasKeyDown) {
                std::string fn = cProcName(info.ctrlName + "_KeyDown", AccessLevel::Private);
                c_.emitLine("if (msg == WM_KEYDOWN) {");
                c_.indent();
                c_.emitLine("{ int16_t vb6_keycode = (int16_t)wp; int16_t vb6_shift = 0;");
                c_.emitLine("if (GetKeyState(VK_SHIFT) & 0x8000) vb6_shift |= 1; if (GetKeyState(VK_CONTROL) & 0x8000) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
                c_.emitLine("extern void " + fn + "(int16_t*, int16_t*); " + fn + "(&vb6_keycode, &vb6_shift); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // KeyUp (WM_KEYUP)
            if (info.hasKeyUp) {
                std::string fn = cProcName(info.ctrlName + "_KeyUp", AccessLevel::Private);
                c_.emitLine("if (msg == WM_KEYUP) {");
                c_.indent();
                c_.emitLine("{ int16_t vb6_keycode = (int16_t)wp; int16_t vb6_shift = 0;");
                c_.emitLine("if (GetKeyState(VK_SHIFT) & 0x8000) vb6_shift |= 1; if (GetKeyState(VK_CONTROL) & 0x8000) vb6_shift |= 2; if (GetKeyState(VK_MENU) & 0x8000) vb6_shift |= 4;");
                c_.emitLine("extern void " + fn + "(int16_t*, int16_t*); " + fn + "(&vb6_keycode, &vb6_shift); }");
                c_.dedent();
                c_.emitLine("}");
            }
            // P1-13修复: DblClick (WM_LBUTTONDBLCLK)
            if (info.hasDblClick) {
                std::string fn = cProcName(info.ctrlName + "_DblClick", AccessLevel::Private);
                // C29-SL-n（账 #141）: VB6 的签名是 `Sub X_DblClick(Cancel As Integer)`，
                // 而这一条 arm 以前**写死无参** ⇒ 按标准签名写的处理器在产物里撞
                // `C2198: 用于调用的参数太多`（实测：声明是 `static void vb6_sldM6_DblClick(int16_t* Cancel)`，
                // arm 里却 `extern void vb6_sldM6_DblClick(); …`）。存量工程两种都见过
                // （本仓夹具就写着无参的 `sld6_DblClick()`），所以按**处理器自己声明的形参数**发。
                const Symbol* symDc = symTab_.lookup(info.ctrlName + "_DblClick");
                if (symDc && !symDc->params.empty()) {
                    c_.emitLine("if (msg == WM_LBUTTONDBLCLK) { int16_t vb6_dblcancel = 0;"
                                " extern void " + fn + "(int16_t*); " + fn + "(&vb6_dblcancel); }");
                } else {
                    c_.emitLine("if (msg == WM_LBUTTONDBLCLK) { extern void " + fn + "(); " + fn + "(); }");
                }
            }
            // C29-SL-d: 常规 Click (WM_LBUTTONUP)。VB6 的 Click 是"按下并抬起"之后发的，
            // 抬起这一条就是它的原生落点（探针 .build/slprobe/slmeasure10.c 实测：真手势的
            // WM_LBUTTONDOWN / WM_LBUTTONUP 都会经过控件自己的子类过程 —— 子类化换的是
            // WNDPROC，那条 hwnd 的每条消息都先过我们这一段；而原生轨道条自己在
            // button-down 里就把焦点抢过去了，所以按键那两条 arm 不需要 WS_TABSTOP 也到得了）。
            // 只发给**没有原生 Click 来源**的控件（那批由 BN_CLICKED / *_SELCHANGE /
            // TCN_SELCHANGE 供，见 controlClickFromNativeNotify），否则一次点击双发。
            if (info.hasClick) {
                std::string fn = cProcName(info.ctrlName + "_Click", AccessLevel::Private);
                c_.emitLine("if (msg == WM_LBUTTONUP) { extern void " + fn + "(); " + fn + "(); }");
            }

            // P20-12: WM_VB6_RESTOREFOCUS - restore focus after Validate cancel
            if (info.hasValidate) {
                c_.emitLine("if (msg == 0x7FFF) { RemovePropW(hwnd, L\"VB6_ValidateCancel\"); SetFocus(hwnd); return 0; }");
            }

            // Fix 185: 控件级 _Paint 派发 (PictureBox)。DC 用窗口属性传给
            // vb6_ControlPrint/vb6_ControlCls —— VB6 的 Print 在 _Paint 里画的正是
            // 这次 BeginPaint 拿到的 DC。必须 return 0: 否则消息落到原 STATIC 过程
            // 再 BeginPaint 一次，并且用户画的内容会被它的默认绘制覆盖。
            if (info.hasPaint) {
                std::string fn = cProcName(info.ctrlName + "_Paint", AccessLevel::Private);
                c_.emitLine("if (msg == WM_PAINT) {");
                c_.indent();
                c_.emitLine("PAINTSTRUCT vb6_ps;");
                c_.emitLine("HDC vb6_hdc = BeginPaint(hwnd, &vb6_ps);");
                c_.emitLine("if (vb6_hdc) {");
                c_.indent();
                c_.emitLine("SetPropW(hwnd, L\"VB6_PaintDC\", (HANDLE)vb6_hdc);");
                c_.emitLine("{ extern void " + fn + "(); " + fn + "(); }");
                c_.emitLine("RemovePropW(hwnd, L\"VB6_PaintDC\");");
                c_.emitLine("EndPaint(hwnd, &vb6_ps);");
                c_.dedent();
                c_.emitLine("}");
                c_.emitLine("return 0;");
                c_.dedent();
                c_.emitLine("}");
            }

            // WM_DESTROY: 移除子类化 + 清理ValidateCancel属性
            c_.emitLine("if (msg == WM_DESTROY) { RemovePropW(hwnd, L\"VB6_ValidateCancel\"); vb6_RemoveControlSubclass((void*)hwnd); }");

            // Call original WndProc for unhandled messages
            c_.emitLine("WNDPROC vb6_orig = (WNDPROC)vb6_GetOriginalWndProc((void*)hwnd);");
            c_.emitLine("if (vb6_orig) return CallWindowProcW(vb6_orig, hwnd, msg, wp, lp);");
            c_.emitLine("return DefWindowProcW(hwnd, msg, wp, lp);");
            c_.dedent();
            c_.emitLine("}");
            c_.emitBlank();
        }

        // Store subclass controls for WM_CREATE installation (used later in CreateControls)
        // We'll emit the install calls in the createFn, after all controls are created
    }
