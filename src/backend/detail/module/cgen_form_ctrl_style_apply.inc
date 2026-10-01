// cgen_form_ctrl_style_apply.inc - emitFormFramework 函数体片段：CreateControls 之二 —— 样式常量 + CreateWindow + 属性应用
// 由 src/backend/detail/module/cgen_form_create_controls.inc 再拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1476~1771 行，逐行未改。

        // Win32样式常量（不依赖windows.h，使用WinUser.h固定数值）
        // WS_CHILD=0x40000000, WS_VISIBLE=0x10000000, WS_BORDER=0x00800000
        // BS_PUSHBUTTON=0, BS_AUTOCHECKBOX=3, BS_AUTORADIOBUTTON=9
        // BS_GROUPBOX=7, SS_LEFT=0, ES_AUTOHSCROLL=0x80
        // LBS_NOTIFY=1, CBS_DROPDOWN=2
        constexpr long kWsChild    = 0x40000000L;
        constexpr long kWsClipSib  = 0x04000000L;  // WS_CLIPSIBLINGS (Fix 124)
        constexpr long kWsVisible  = 0x10000000L;
        constexpr long kWsBorder   = 0x00800000L;
        constexpr long kBsPush     = 0x00000000L;
        constexpr long kBsAutoChk  = 0x00000003L;
        constexpr long kBsAutoRad  = 0x00000009L;
        constexpr long kBsGroupBox = 0x00000007L;
        constexpr long kSsLeft     = 0x00000000L;
        constexpr long kSsNotify   = 0x00000100L;  // SS_NOTIFY: Label 接收鼠标消息 (VB6 语义 + tooltip 需要)
        constexpr long kEsAutoH    = 0x00000080L;
        constexpr long kLbsNotify  = 0x00000001L;
        constexpr long kCbsDrop    = 0x00000002L;
        constexpr long kSsBitmap   = 0x0000000EL;  // SS_BITMAP for PictureBox/Image
        constexpr long kSsCenterImg= 0x00000200L;  // SS_CENTERIMAGE
        constexpr long kBsBitmap   = 0x00000080L;  // BS_BITMAP for Graphical buttons
        constexpr long kBsPushLike = 0x00001000L;  // BS_PUSHLIKE: button-style toggle
        constexpr long kEsMulti    = 0x00000004L;  // ES_MULTILINE for TextBox
        constexpr long kEsAutoV    = 0x00000040L;  // ES_AUTOVSCROLL
        constexpr long kWsHscroll  = 0x00100000L;  // WS_HSCROLL
        constexpr long kWsVscroll  = 0x00200000L;  // WS_VSCROLL
        // P20-40: StatusBar — SBARS_SIZEGRIP/SBARS_TOOLTIPS 是 commctrl.h 的类专属样式
        // (不是 WS_*, 但仍要进 CreateWindow 的 style 参数)。RTL 自己算面板文本,
        // 故这里只给 sizegrip (与 VB6 默认观感一致), tooltip 由 RTL 存着不激活。
        constexpr long kSbarsGrip  = 0x00000100L;  // SBARS_SIZEGRIP
        constexpr long kSbarsTip   = 0x00000800L;  // SBARS_TOOLTIPS

        // Fix 124: WS_CLIPSIBLINGS —— 标准控件与 UC 宿主同层叠放时,
        // 重绘不得画到上层兄弟控件上 (否则出现"控件被其他控件擦掉")。
        long style = kWsChild | kWsVisible | kWsClipSib;
        long exStyle = 0;
        std::string createCaption = ctrlCaption;

        switch (ctrl.controlType) {
            case FrmControlType::CommandButton:
                style |= kBsPush;
                {
                    auto stIt = ctrl.properties.find("Style");
                    if (stIt != ctrl.properties.end() && stIt->second.intValue == 1)
                        style |= kBsPushLike;  // Graphical: button-style with picture (no BS_BITMAP, we draw ourselves)
                }
                break;
            case FrmControlType::TextBox:
                style |= kWsBorder | kEsAutoH;
                // Bug3-Fix: 不在CreateControl中传初始文本，避免EN_CHANGE在句柄赋值前触发
                // createCaption will be set to "" below; text set after hwnd assignment
                {
                    auto mlIt = ctrl.properties.find("MultiLine");
                    if (mlIt != ctrl.properties.end() && mlIt->second.intValue != 0)
                        style |= kEsMulti | kEsAutoV;
                    auto sbIt = ctrl.properties.find("ScrollBars");
                    if (sbIt != ctrl.properties.end()) {
                        int sb = (int)sbIt->second.intValue;
                        // 账 #108: VB6 的枚举是 0 无 / 1 水平 / 2 垂直 / 3 两者 —— 这里原先
                        // 1↔2 用反了（与 RichTextBox 那一格两套方向）。证人不是文档而是
                        // VB6 自己存出来的 .frm：tests/VBFlexGridDemo/InputForm.frm 19 行
                        // `Begin VB.TextBox Text2` 下挂着 `ScrollBars = 2  'Vertical`。
                        // 读侧 vb6_GetScrollBars 一直按 VB6 口径推（HSCROLL→1、VSCROLL→2），
                        // 所以写侧反着的结果是"设计期设 2、读回 1"——往返不自洽，不是观感问题。
                        if (sb == 1 || sb == 3) style |= kWsHscroll;  // Horizontal
                        if (sb == 2 || sb == 3) style |= kWsVscroll;  // Vertical
                    }
                }
                break;
            case FrmControlType::Label:
                // VB6 Label 有 Click/MouseDown 事件且是 tooltip 目标, 必须 SS_NOTIFY,
                // 否则 STATIC 对鼠标返回 HTTRANSPARENT, TTF_SUBCLASS 的 tooltip 永远不弹。
                style |= kSsLeft | kSsNotify;
                break;
            case FrmControlType::CheckBox:
                style |= kBsAutoChk;
                {
                    auto stIt = ctrl.properties.find("Style");
                    if (stIt != ctrl.properties.end() && stIt->second.intValue == 1)
                        style |= kBsPushLike;  // Graphical: button-style toggle with picture (no BS_BITMAP, we draw ourselves)
                }
                break;
            case FrmControlType::OptionButton:
                style |= kBsAutoRad;
                {
                    auto stIt = ctrl.properties.find("Style");
                    if (stIt != ctrl.properties.end() && stIt->second.intValue == 1)
                        style |= kBsPushLike;  // Graphical: button-style toggle with picture (no BS_BITMAP, we draw ourselves)
                }
                break;
            case FrmControlType::Frame:
                style |= kBsGroupBox;
                break;
            case FrmControlType::ListBox: {
                style |= kLbsNotify | kWsBorder | kWsVscroll;  // VB6 ListBox always shows vertical scrollbar
                // Sorted: LBS_SORT = 0x0002
                auto sortIt = ctrl.properties.find("Sorted");
                if (sortIt != ctrl.properties.end() && sortIt->second.intValue != 0) style |= 0x0002L;
                // MultiSelect: 1=LBS_MULTIPLESEL(0x8), 2=LBS_EXTENDEDSEL(0x800)
                auto msIt = ctrl.properties.find("MultiSelect");
                if (msIt != ctrl.properties.end()) {
                    if (msIt->second.intValue == 1) style |= 0x0008L;
                    else if (msIt->second.intValue == 2) style |= 0x0800L;
                }
                break;
            }
            case FrmControlType::ComboBox:
                // Bug3-Fix: ComboBox同TextBox，延迟设置初始文本
                // createCaption will be set to "" below
                {
                    // Style: 0=CBS_DROPDOWN(2), 1=CBS_SIMPLE(1), 2=CBS_DROPDOWNLIST(3)
                    auto stIt = ctrl.properties.find("Style");
                    if (stIt != ctrl.properties.end()) {
                        if (stIt->second.intValue == 1) style |= 0x0001L;  // CBS_SIMPLE
                        else if (stIt->second.intValue == 2) style |= 0x0003L;  // CBS_DROPDOWNLIST
                        else style |= kCbsDrop;  // default: CBS_DROPDOWN
                    } else {
                        style |= kCbsDrop;
                    }
                    // Fix 147: 同 cgen_util_ctrl.cpp — ComboBox 下拉列表需要
                    // WS_VSCROLL 才能滚动查看全部项 (否则只能看到前几项)
                    style |= kWsVscroll;
                    style |= kWsBorder;
                    // Sorted: CBS_SORT = 0x0100
                    auto sortIt = ctrl.properties.find("Sorted");
                    if (sortIt != ctrl.properties.end() && sortIt->second.intValue != 0) style |= 0x0100L;
                }
                break;
            case FrmControlType::PictureBox:
                // VB6 PictureBox: STATIC + SS_BITMAP + SS_CENTERIMAGE + border
                style |= kSsBitmap | kSsCenterImg;
                style |= kWsBorder;  // WS_BORDER for sunken border
                createCaption = "";   // No text for picture controls
                break;
            case FrmControlType::Image:
                // VB6 Image: STATIC + SS_BITMAP + SS_CENTERIMAGE (no border)
                style |= kSsBitmap | kSsCenterImg;
                createCaption = "";
                break;
            case FrmControlType::ProgressBar:
                // VB6 ProgressBar 无标题; 外观属性 (Orientation/Scrolling) 由
                // vb6_InitProgressBar 在窗口创建后改写 GWL_STYLE, 这里只给默认水平。
                createCaption = "";
                break;
            case FrmControlType::StatusBar:  // P20-40: msctls_status32
                style |= kSbarsGrip;
                createCaption = "";   // 面板文本走 SB_SETTEXTW, 创建时不给标题
                break;
            case FrmControlType::SSTab:      // P20-42: SysTabControl32
                // 位置/外观类 TCS_* 由 vb6_SSTab_Init 在窗口创建后改写 GWL_STYLE
                // (创建时读设计期属性容易与"未写即默认"的口径打架), 这里只给基类样式。
                createCaption = "";   // 页标题走 TCM_INSERTITEM, 创建时不给标题
                break;
            case FrmControlType::DriveListBox:
                // C29-1b: VB6 的 DriveListBox 外观就是"只能选、不能改"的组合框
                style |= 0x0003L | kWsVscroll | kWsBorder;  // CBS_DROPDOWNLIST
                createCaption = "";
                break;
            case FrmControlType::DirListBox:
            case FrmControlType::FileListBox:
                // 列表框同款: LBS_NOTIFY 才有 SELCHANGE / DBLCLK 通知可派
                style |= kLbsNotify | kWsBorder | kWsVscroll;
                createCaption = "";
                break;
            case FrmControlType::Shape:
            case FrmControlType::Line:
                // C29-1a: 自绘控件 (VB6_SHAPE / VB6_LINE)，无文本、无边框样式
                createCaption = "";
                break;
            case FrmControlType::CommonDialog:
                // D6 / C29-9: 这枚"控件"没有外观 —— 它是原生对话框的属性袋。
                // 窗口照建（要有句柄存属性），但必须**不可见**：基础样式里有 WS_VISIBLE，
                // 这里清掉，尺寸也由下面的 0x0 兜住。
                createCaption = "";
                style &= ~kWsVisible;
                width = 0;
                height = 0;
                break;
            case FrmControlType::Timer:
                // C29-T: 与 CommonDialog 同一手法 —— Timer 要有句柄（当计时器的身份），
                // 但绝不能有外观：不可见、0x0，且不吃 Tab 焦点。
                createCaption = "";
                style &= ~kWsVisible;
                width = 0;
                height = 0;
                break;
            case FrmControlType::Winsock:
                // C29-WS-a: 又一枚"无外观但有身份"的控件，外形处理与 Timer 完全同法。
                // **实质差别在窗口过程**：Timer 那枚的 WndProc 是裸 DefWindowProc（纯属性袋），
                // 而 Winsock 的身份窗要**真的收 WSAAsyncSelect 投来的 FD_\***（通知投给它自己，
                // 不是父窗）⇒ 类注册在 RTL（vb6_RegisterWinsockClass），那里才是有逻辑的一头。
                createCaption = "";
                style &= ~kWsVisible;
                width = 0;
                height = 0;
                break;
            case FrmControlType::TreeView:
                // C29-8a: SysTreeView32 的创建路早就有 (controlTypeToWin32Class 有格)，
                // 缺的是属性面。这里只给 VB6 的**默认观感**: LineStyle=0 (画树线、根层不画)
                // + BorderStyle=1 (Fixed)。其余四位由设计期 Init 那条补 (下面 vb6_TreeView_Init)。
                // WS_BORDER: 与 ListBox/Dir/File 同口径，否则 VB6 那圈凹边框没了。
                style |= 0x0002L | kWsBorder;   // TVS_HASLINES
                createCaption = "";
                break;
            case FrmControlType::DTPicker: {
                // C29-DT-a: SysDateTimePick32 的三条样式属性一律从 .frm 立进**创建参数**。
                // 为什么不能事后 SET（实测，读数见 tests\ctrldatetime 的 W 那行与 DT9-DT11）：
                //   · CheckBox / UpDown 是复选框/微调按钮那两枚**子窗口**的创建参数，控件只在
                //     WM_CREATE 读一次；事后写 GWL_STYLE，它立刻把这两位抹回去 ⇒ 只能创建时给。
                //   · Format 切档运行期倒是有效，但设计期本来就该在创建时对，所以三条同一条路。
                // 不加 WS_BORDER：日期框自带主题边框（TreeView/ListBox 那几枚要边框是因为原生
                // 类自己不画），再套一圈就是两道框。
                createCaption = "";
                auto dtProp = [&](const char* key, int& slot) {
                    auto it = ctrl.properties.find(key);
                    if (it != ctrl.properties.end()) slot = (int)it->second.intValue;
                };
                int dtFmt = -999, dtChk = -999, dtUp = -999;
                dtProp("Format", dtFmt);
                dtProp("CheckBox", dtChk);
                dtProp("UpDown", dtUp);
                // VB6 的 dtpFormat 0..3：0 短日期(=原生 0) / 1 长日期 / 2 时间 / 3 自定义。
                // 3 不改格式位（画成什么由 CustomFormat 决定），与 vb6_DTP_SetFormat 同一口径。
                if (dtFmt == 1) style |= 0x0004L;          // DTS_LONGDATEFORMAT
                else if (dtFmt == 2) style |= 0x0008L;     // DTS_TIMEFORMAT 的格式段(0x9 的 bit3；
                                                           //   bit0 是 UPDOWN 那位，不在这儿塞)
                if (dtChk != -999 && dtChk != 0) style |= 0x0002L;   // DTS_SHOWNONE
                if (dtUp  != -999 && dtUp  != 0) style |= 0x0001L;   // DTS_UPDOWN
                break;
            }
            case FrmControlType::MonthView: {
                // C29-MV-a: SysMonthCal32 的三位样式一律从 .frm 立进创建参数。
                //   MultiSelect=→MCS_MULTISELECT(0x2)  ShowWeekNumbers=→MCS_WEEKNUMBERS(0x4)
                //   ShowToday=False→MCS_NOTODAY(0x10)  —— 注意 ShowToday 与原生那位是**反**的，
                // 所以只有明确写 False 才挂 NOTODAY；没写就不动（原生默认显示 today 那一行）。
                // 不加 WS_BORDER：月历自带主题边框，再套一圈就是两道框（与 DTPicker 同因）。
                createCaption = "";
                auto mvProp = [&](const char* key, int& slot) {
                    auto it = ctrl.properties.find(key);
                    if (it != ctrl.properties.end()) slot = (int)it->second.intValue;
                };
                int mvMulti = -999, mvWeek = -999, mvToday = -999;
                mvProp("MultiSelect", mvMulti);
                mvProp("ShowWeekNumbers", mvWeek);
                mvProp("ShowToday", mvToday);
                if (mvMulti != -999 && mvMulti != 0) style |= 0x0002L;
                if (mvWeek  != -999 && mvWeek  != 0) style |= 0x0004L;
                if (mvToday != -999 && mvToday == 0) style |= 0x0010L;
                break;
            }
            case FrmControlType::Slider: {
                // C29-SL-a/h: msctls_trackbar32 的创建样式。折算整个搬进 sliderStyleBits
                // （上面 controlTypeStyleBits 的 Slider 那一格共用它）—— 顶层这条路过去自己
                // 写一份、容器子控件那条路一份都没写，正是账 #83 说的那两张表各写各的。
                // 三条实测口径抄在 vb6forms_slider.c 的 C29-SL-h 那一格与 029 §九：
                //   · Orientation 运行期写样式位**是有效的**（第一发探针拿 channel 矩形判
                //     方向，得出过"只能创建时给"的错结论 —— 那条消息不随方向旋转）。
                //   · TBS_AUTOTICKS(0x0001) 一律挂上：VB6 的 Slider 默认就画刻度。
                //   · TickStyle 四档 ↔ TBS_TOP(0x4)/BOTH(0x8)/NOTICKS(0x10)/那三位全清，
                //     数值读自 OCX 自带的类型库（探针 .build/slprobe/sltlb.cpp），非猜。
                createCaption = "";
                style |= sliderStyleBits(ctrl);
                break;
            }
            case FrmControlType::RichTextBox: {
                // C29-RT-a: RICHEDIT50W 的创建样式。
                //   · ES_MULTILINE 一律挂上：原生这枚控件不是多行就没有滚动条一说（与 TextBox
                //     不同，VB6 的 RichTextBox 本身就是多行的）。
                //   · ScrollBars 从 .frm 立进创建参数，**不走运行期 setter** —— 探针量出来事后
                //     SetWindowLong 只有外观（客户区让位 21px）、滚动量程停在默认 0..100，
                //     内容真实高度 0..504 滚不动。RTL 侧的 VScrollRange/HScrollRange 就是这条
                //     结论的证人，夹具 RT7-RT10 读它。
                //   · 枚举口径按 VB6 文档：0 无 / 1 水平 / 2 垂直 / 3 两者。**与上面 TextBox
                //     那一格相反** —— 那一格把 1 当垂直、2 当水平（既有缺陷，本格不顺手改，
                //     账记在 029 §九）。
                createCaption = "";
                style |= kEsMulti;
                auto rtSb = ctrl.properties.find("ScrollBars");
                if (rtSb != ctrl.properties.end()) {
                    int sb = (int)rtSb->second.intValue;
                    if (sb == 1 || sb == 3) style |= kWsHscroll;
                    if (sb == 2 || sb == 3) style |= kWsVscroll;
                    // ES_DISABLENOSCROLL：不挂它，控件在"内容暂时不需要滚动"时把 bar 连样式位
                    // 一起拆掉（实测空文本下 GWL_STYLE 那两位读回 0），ScrollBars 就没法答。
                    if (sb != 0) style |= 0x00002000L;
                }
                break;
            }
            case FrmControlType::Toolbar:
                // C29-5a: 原生 ToolbarWindow32 的默认观感两条 + **两位绝不能省**:
                //   CCS_NORESIZE(0x4) | CCS_NOPARENTALIGN(0x8) —— 不给这两位，控件会无视
                //   .frm 的 Left/Top/Width/Height 自己去重排父窗顶部那一条 (原生工具栏的
                //   默认行为)，VB6 那种"放在窗体中间的一小条"就做不出来。
                //   TBSTYLE_TOOLTIPS(0x100) = VB6 的 ShowTips 默认 True；TBSTYLE_FLAT(0x800)
                //   是 VB6 6.0 的扁平观感。TextStyle/AllowCustomize 由下面的 Init 那条改。
                style |= 0x0004L | 0x0008L | 0x0100L | 0x0800L;
                createCaption = "";
                break;
            default:
                break;
        }

        // 账 #83(a)（C29-SL-r）: VB6 默认 TabStop = True，而这一位以前两条创建路都没立
        // （读数见 029 的 C29-SL-r-0）。判断只在 controlTabStopStyleBit 一处，容器那趟共用。
        style |= controlTabStopStyleBit(ctrl);

        // 账 #158（按钮那一半）: BUTTON 类不挂 `BS_NOTIFY` 就不把焦点变化作为 `WM_COMMAND`
        // 报给父窗（裸码实测见 `controlButtonNotifyStyleBit` 的定义处）。判断只在那一处，容器那趟共用。
        style |= controlButtonNotifyStyleBit(ctrl);
        // 账 #83(b2): 容器挂 WS_EX_CONTROLPARENT —— 没有这一位，`IsDialogMessage` 不会
        // 走进 Frame/PictureBox/SSTab，里面的控件连 tab 序都进不了（顶层这一条路）。
        exStyle |= controlContainerExStyleBit(ctrl);

        // Bug3-Fix: TextBox/ComboBox 延迟设置初始文本，避免 EN_CHANGE 在句柄赋值前触发
        // C29-RT-a: RichTextBox 同因加入 —— 它也是文本控件，.frm 里的初值走 Text 那一条。
        bool needDelayedText = (ctrl.controlType == FrmControlType::TextBox ||
                                ctrl.controlType == FrmControlType::ComboBox ||
                                ctrl.controlType == FrmControlType::RichTextBox);
        std::string delayedText = (ctrl.controlType == FrmControlType::TextBox ||
                                   ctrl.controlType == FrmControlType::RichTextBox)
                                      ? ctrlText : ctrlCaption;
        if (needDelayedText) createCaption = "";  // 创建时不传文本

        // C29-1a: Line 在 VB6 里没有 Left/Top/Width/Height —— 窗口矩形就是 X1/Y1/X2/Y2
        // (容器缇值) 的包围盒。不这样算, 控件会拿默认尺寸 1000x300 创建出来。
        if (ctrl.controlType == FrmControlType::Line) {
            lineRectFromEndpoints(ctrl, left, top, width, height);
        }

        // 生成vb6_CreateControl调用
        // 生成vb6_CreateControl调用 (P7.6: 数组控件用CtrlArr_SetAt)
        {
            std::string ctrlNameLower = ctrl.controlName;
            std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
            c_.emitLine("{ void* vb6_tmp_hwnd = vb6_CreateControl(");
            c_.indent();
            c_.emitLine("\"" + std::string(win32Class) + "\", \"" + createCaption + "\",");
            c_.emitLine(std::to_string(style) + "L, " + std::to_string(exStyle) + "L,");
            c_.emitLine(std::to_string(left) + ", " + std::to_string(top) + ", "
                       + std::to_string(width) + ", " + std::to_string(height) + ",");
            c_.emitLine(std::to_string(ctrlId) + ", hwnd, hInstance);");
            c_.dedent();
            if (knownControlArrays_.count(ctrlNameLower)) {
                // 控件数组: 存入vb6_CtrlArr
                c_.emitLine("vb6_CtrlArr_SetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", "
                           + std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ", vb6_tmp_hwnd); }");
            } else {
                c_.emitLine("vb6_hwnd_" + cIdent(ctrl.controlName) + " = vb6_tmp_hwnd; }");
            }
        }

        // 设计期 BorderStyle / Alignment (顶层控件)
        {
            std::string dsNameLower = ctrl.controlName;
            std::transform(dsNameLower.begin(), dsNameLower.end(), dsNameLower.begin(), ::tolower);
            std::string dsHwnd = knownControlArrays_.count(dsNameLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            emitDesignerStyleProps(ctrl, dsHwnd);
            emitDesignerStateProps(ctrl, dsHwnd);  // 账 #125: Enabled / Visible / Value
            emitDesignerStringProps(ctrl, dsHwnd);  // 账 #142: ToolTipText / Tag（非空才发）
            emitDesignerFontProps(ctrl, dsHwnd);    // 账 #154: FontName / FontSize（写过才发）
            emitDesignerTabIndexProp(ctrl, dsHwnd, tabIndexFallback);  // 账 #160: 设计期 TabIndex（顶层这条路）

            // Fix 187: 设计期 ForeColor/BackColor — 此前从未发射, .frm 的初始色
            // 全部落空 (picBackColor &H00F8CCB1& 淡蓝被显示成 BTNFACE 灰, 颜色
            // 对话框 OK 后才第一次 Set)。&H8000000C& 这类系统色经低位截断保留
            // 0x80000000 位, RTL 侧按 GetSysColor 解析。
            {
                auto bcIt187 = ctrl.properties.find("BackColor");
                if (bcIt187 != ctrl.properties.end() && bcIt187->second.type == FrmValueType::Integer) {
                    c_.emitLine("vb6_SetControlBackColor((void*)" + dsHwnd + ", "
                                + std::to_string((int)(unsigned int)bcIt187->second.intValue) + ");  /* design BackColor */");
                }
                auto fcIt187 = ctrl.properties.find("ForeColor");
                if (fcIt187 != ctrl.properties.end() && fcIt187->second.type == FrmValueType::Integer) {
                    c_.emitLine("vb6_SetControlForeColor((void*)" + dsHwnd + ", "
                                + std::to_string((int)(unsigned int)fcIt187->second.intValue) + ");  /* design ForeColor */");
                }
            }
        }

        // Bug3-Fix: 延迟设置 TextBox/ComboBox 的初始文本（在句柄赋值之后）
        if (needDelayedText && !delayedText.empty()) {
            std::string ctrlNameLower2 = ctrl.controlName;
            std::transform(ctrlNameLower2.begin(), ctrlNameLower2.end(), ctrlNameLower2.begin(), ::tolower);
            std::string hwndVar = knownControlArrays_.count(ctrlNameLower2) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            // 使用 SendMessage WM_SETTEXT 设置初始文本（句柄已赋值，EN_CHANGE 安全触发）
            c_.emitLine("SendMessageW((HWND)" + hwndVar + ", WM_SETTEXT, 0, (LPARAM)L\"" + delayedText + "\");");
        }

        // Fix 162c-extlist 那条"顶层控件设计期 Value"的旧块**已删**(2026-09-30)。
        // 它和上面第 360 行的 `emitDesignerStateProps(ctrl, dsHwnd)` 是**同一件事的两个出口**,
        // 而两条口径不一致 —— 旧块 `CheckBox || OptionButton` 共用 `vb6_SetCheckValue`:
        //   · OptionButton 的 Value 在 VB6 是 Boolean(账 #128-b 单开了 vb6_SetOptionValue);
        //   · CheckBox 的 Value 是三态(0/1/2), 旧块把 Grayed(2) 也压成 1。
        // 于是 `obOn` 会同时拿到 `vb6_SetOptionValue` 与 `vb6_SetCheckValue` 两条
        // (实测门 #233 的 `cs_emitc_selectivity` 就是钉这条错形状的), 读数取决于谁最后落地。
        // 现在设计期状态属性**只有一处出口** = `emitDesignerStateProps`, 两条创建路
        // (顶层 cgen_form_ctrl_style_apply.inc / 容器 cgen_form_frame_menu.inc) 各调一次。
        // ⚠ 这条纪律与账 #125 的注释同源: 设计期属性按"属性"分出口, 不按"控件类型"分出口。

        // Fix 162c-extlist: Label 设计期 AutoSize=True —— 此前从不消费
        // (extlist 的 Label2/3/4 按设计期 caption "Label2" 的宽度留在原地)。
        if (ctrl.controlType == FrmControlType::Label) {
            auto asIt162c = ctrl.properties.find("AutoSize");
            if (asIt162c != ctrl.properties.end() && asIt162c->second.intValue != 0) {
                std::string asNLower = ctrl.controlName; std::transform(asNLower.begin(), asNLower.end(), asNLower.begin(), ::tolower);
                std::string asHwnd = knownControlArrays_.count(asNLower) ?
                    ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                     std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                    ("vb6_hwnd_" + cIdent(ctrl.controlName));
                c_.emitLine("vb6_SetLabelAutoSize((void*)" + asHwnd + ", -1);  /* design AutoSize */");
            }
        }

        // P20-12: CausesValidation property from .frm (default=True, only set when explicitly False)
        {
            auto cvIt = ctrl.properties.find("CausesValidation");
            if (cvIt != ctrl.properties.end() && cvIt->second.intValue == 0) {
                std::string cvNLower = ctrl.controlName; std::transform(cvNLower.begin(), cvNLower.end(), cvNLower.begin(), ::tolower);
                std::string cvHwnd = knownControlArrays_.count(cvNLower) ?
                    ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)") :
                    ("vb6_hwnd_" + cIdent(ctrl.controlName));
                c_.emitLine("vb6_SetCausesValidation((void*)" + cvHwnd + ", 0);");
            }
        }

        
        // Image.Stretch from .frm (default=False; only set when explicitly True)
        if (ctrl.controlType == FrmControlType::Image) {
            auto stIt = ctrl.properties.find("Stretch");
            if (stIt != ctrl.properties.end() && stIt->second.intValue != 0) {
                std::string stNameLower = ctrl.controlName;
                std::transform(stNameLower.begin(), stNameLower.end(), stNameLower.begin(), ::tolower);
                std::string stHwnd = knownControlArrays_.count(stNameLower) ?
                    ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                     std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                    ("vb6_hwnd_" + cIdent(ctrl.controlName));
                c_.emitLine("vb6_SetImageStretch((void*)" + stHwnd + ", 1);");
            }
        }

        // C29-1a: 设计期外观属性落到控件 (原来这里是一整段内联块; 容器子控件那条路共用同一个助手)
        emitShapeLineProps(ctrl, ctrlHwndExprForInit(ctrl));

        // P20-37 / C29-1b: 三控件的初值。设计期给了 Drive / Path / Pattern 就用它, 没给才
        // 退回当前目录 (FileListBox 的 Pattern 缺省时由 RTL 回 "*.*")。顺序上先填表再设值,
        // SetDrive / SetPath / SetPattern 自己会重刷列表。
        if (ctrl.controlType == FrmControlType::DriveListBox) {
            std::string p37Hwnd = "vb6_hwnd_" + cIdent(ctrl.controlName);
            c_.emitLine("vb6_DriveListBoxRefresh((void*)" + p37Hwnd + ");");
            auto dIt = ctrl.properties.find("Drive");
            if (dIt != ctrl.properties.end() && dIt->second.type == FrmValueType::String) {
                std::string dRaw = dIt->second.rawText;
                if (dRaw.size() >= 2 && dRaw.front() == '"' && dRaw.back() == '"')
                    dRaw = dRaw.substr(1, dRaw.size() - 2);
                if (!dRaw.empty())
                    c_.emitLine("vb6_DriveListBoxSetDrive((void*)" + p37Hwnd
                                + ", vb6_BSTR_FromStr(L\"" + escapeWideCString(dRaw) + "\"));");
            }
        } else if (ctrl.controlType == FrmControlType::DirListBox) {
            std::string p37Var = "_p" + cIdent(ctrl.controlName);
            std::string p37Hwnd2 = "vb6_hwnd_" + cIdent(ctrl.controlName);
            auto pIt = ctrl.properties.find("Path");
            bool hasPath = (pIt != ctrl.properties.end() && pIt->second.type == FrmValueType::String);
            std::string pRaw;
            if (hasPath) {
                pRaw = pIt->second.rawText;
                if (pRaw.size() >= 2 && pRaw.front() == '"' && pRaw.back() == '"')
                    pRaw = pRaw.substr(1, pRaw.size() - 2);
            }
            if (hasPath && !pRaw.empty()) {
                c_.emitLine("vb6_DirListBoxSetPath((void*)" + p37Hwnd2
                            + ", vb6_BSTR_FromStr(L\"" + escapeWideCString(pRaw) + "\"));");
            } else {
                c_.emitLine("{ wchar_t " + p37Var + "[MAX_PATH]; GetCurrentDirectoryW(MAX_PATH, " + p37Var + ");");
                c_.emitLine("  vb6_DirListBoxSetPath((void*)" + p37Hwnd2 + ", " + p37Var + "); }");
            }
        } else if (ctrl.controlType == FrmControlType::FileListBox) {
            std::string p37Var2 = "_p" + cIdent(ctrl.controlName);
            std::string p37Hwnd3 = "vb6_hwnd_" + cIdent(ctrl.controlName);
            auto qIt = ctrl.properties.find("Path");
            bool hasQ = (qIt != ctrl.properties.end() && qIt->second.type == FrmValueType::String);
            std::string qRaw;
            if (hasQ) {
                qRaw = qIt->second.rawText;
                if (qRaw.size() >= 2 && qRaw.front() == '"' && qRaw.back() == '"')
                    qRaw = qRaw.substr(1, qRaw.size() - 2);
            }
            if (hasQ && !qRaw.empty()) {
                c_.emitLine("vb6_FileListBoxSetPath((void*)" + p37Hwnd3
                            + ", vb6_BSTR_FromStr(L\"" + escapeWideCString(qRaw) + "\"));");
            } else {
                c_.emitLine("{ wchar_t " + p37Var2 + "[MAX_PATH]; GetCurrentDirectoryW(MAX_PATH, " + p37Var2 + ");");
                c_.emitLine("  vb6_FileListBoxSetPath((void*)" + p37Hwnd3 + ", " + p37Var2 + "); }");
            }
            auto kIt = ctrl.properties.find("Pattern");
            if (kIt != ctrl.properties.end() && kIt->second.type == FrmValueType::String) {
                std::string kRaw = kIt->second.rawText;
                if (kRaw.size() >= 2 && kRaw.front() == '"' && kRaw.back() == '"')
                    kRaw = kRaw.substr(1, kRaw.size() - 2);
                if (!kRaw.empty())
                    c_.emitLine("vb6_FileListBoxSetPattern((void*)" + p37Hwnd3
                                + ", vb6_BSTR_FromStr(L\"" + escapeWideCString(kRaw) + "\"));");
            }
        }

        // ProgressBar 设计期属性 (Min/Max/Value/Orientation/Scrolling)。
        // 未写进 .frm 的传 -1 表示"用 VB6 默认" (Min=0 / Max=100 / Value=0 / 水平 / 分块)。
        // 注意 VB6 序列化 ProgressBar 时常见 `Value = -10` (控件默认值的内部表示),
        // 这里一并当作未指定处理, 由 RTL 钳位到 0 —— 与 VB6 运行时的实际读数一致。
        if (ctrl.controlType == FrmControlType::ProgressBar) {
            std::string pbNLower = ctrl.controlName;
            std::transform(pbNLower.begin(), pbNLower.end(), pbNLower.begin(), ::tolower);
            std::string pbHwnd = knownControlArrays_.count(pbNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            auto pbIt = ctrl.properties.find("Min");
            int pbgMin = (pbIt != ctrl.properties.end() && pbIt->second.intValue != -1)
                       ? (int)pbIt->second.intValue : -1;
            pbIt = ctrl.properties.find("Max");
            int pbgMax = (pbIt != ctrl.properties.end() && pbIt->second.intValue != -1)
                       ? (int)pbIt->second.intValue : -1;
            pbIt = ctrl.properties.find("Value");
            int pbgVal = (pbIt != ctrl.properties.end() && pbIt->second.intValue != -1
                          && pbIt->second.intValue > -10)
                       ? (int)pbIt->second.intValue : -1;
            pbIt = ctrl.properties.find("Orientation");
            int pbgOri = (pbIt != ctrl.properties.end() && pbIt->second.intValue != -1)
                       ? (int)pbIt->second.intValue : -1;
            pbIt = ctrl.properties.find("Scrolling");
            int pbgScr = (pbIt != ctrl.properties.end() && pbIt->second.intValue != -1)
                       ? (int)pbIt->second.intValue : -1;
            c_.emitLine("vb6_InitProgressBar((void*)" + pbHwnd + ", "
                      + std::to_string(pbgMin) + ", " + std::to_string(pbgMax) + ", "
                      + std::to_string(pbgVal) + ", " + std::to_string(pbgOri) + ", "
                      + std::to_string(pbgScr) + ");");
        }

        // C29-8a: TreeView 设计期初值 (LineStyle / Indentation / CheckBoxes / HotTracking /
        // HideSelection)。哨兵用 **-999** 而不是 ProgressBar 那套 -1: VB6 的布尔属性序列化成
        // `CheckBoxes = -1`，-1 恰好就是 True，拿它当"未写"等于设计期永远勾不上复选框。
        // 全部未写时这条 Init 是个空转 —— 照样发，为的是发码形状稳定 (emitc 断言要看它)。
        if (ctrl.controlType == FrmControlType::TreeView) {
            std::string tvLower = ctrl.controlName;
            std::transform(tvLower.begin(), tvLower.end(), tvLower.begin(), ::tolower);
            std::string tvHwnd = knownControlArrays_.count(tvLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            int tvLineStyle = -999, tvIndent = -999, tvChk = -999, tvHot = -999, tvHide = -999;
            auto tvGet = [&](const char* key, int& slot) {
                auto tvIt = ctrl.properties.find(key);
                if (tvIt != ctrl.properties.end()) slot = (int)tvIt->second.intValue;
            };
            tvGet("LineStyle", tvLineStyle);
            tvGet("Indentation", tvIndent);
            tvGet("CheckBoxes", tvChk);
            tvGet("HotTracking", tvHot);
            tvGet("HideSelection", tvHide);
            c_.emitLine("vb6_TreeView_Init((void*)" + tvHwnd + ", "
                      + std::to_string(tvLineStyle) + ", " + std::to_string(tvIndent) + ", "
                      + std::to_string(tvChk) + ", " + std::to_string(tvHot) + ", "
                      + std::to_string(tvHide) + ");");
        }

        // C29-DT-a: DTPicker 设计期 CustomFormat 下发。Format / CheckBox / UpDown 三条
        // **不进这条 Init** —— 实测样式位事后写会被控件抹回去（见上面创建样式那一段的注释），
        // 所以它们在上面的样式 switch 里就立好了。未写 CustomFormat 时这条是空转（发空串），
        // 照样发 —— 为的是发码形状稳定（emitc 断言要看它）。
        if (ctrl.controlType == FrmControlType::DTPicker) {
            std::string dtLower = ctrl.controlName;
            std::transform(dtLower.begin(), dtLower.end(), dtLower.begin(), ::tolower);
            std::string dtHwnd = knownControlArrays_.count(dtLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            std::string dtCustom;
            auto cfIt = ctrl.properties.find("CustomFormat");
            if (cfIt != ctrl.properties.end()) {
                // .frm 里的 rawText 带外层引号，且 VB 用 "" 表示串内的一个引号。
                // FrmParser::unquote 是 private（实测 C2248），这里就地剥一遍。
                std::string raw = cfIt->second.rawText;
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
                    std::string inner = raw.substr(1, raw.size() - 2);
                    for (size_t k = 0; k < inner.size(); ++k) {
                        if (inner[k] == '"' && k + 1 < inner.size() && inner[k + 1] == '"') continue;
                        dtCustom += inner[k];
                    }
                } else {
                    dtCustom = raw;
                }
            }
            // C 字面量里的转义：反斜杠与双引号都过一遍，免得格式串里写 `"a\"b"` 把发码撑破。
            std::string esc;
            for (char ch : dtCustom) {
                if (ch == '\\' || ch == '"') esc += '\\';
                esc += ch;
            }
            c_.emitLine("vb6_DTP_Init((void*)" + dtHwnd + ", L\"" + esc + "\");");
        }

        // C29-SL-a/b: Slider 设计期值面下发。C29-SL-o 起这段本身搬进 emitSliderDesignTimeInit
        // （账 #83 的 Slider 半边 —— 容器子控件那条创建路以前压根不发这一条，参数序与 -999
        // 哨兵的口径只在这一处，别在这里复制第二份）。
        if (ctrl.controlType == FrmControlType::Slider) {
            std::string slLower = ctrl.controlName;
            std::transform(slLower.begin(), slLower.end(), slLower.begin(), ::tolower);
            std::string slHwnd = knownControlArrays_.count(slLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            emitSliderDesignTimeInit(ctrl, slHwnd);
        }

        // C29-MV-a: MonthView 设计期的多月平铺。原生没有"行/列"这条消息 —— 平铺几个月完全
        // 由**窗口矩形多大**决定（CommCtrl.h 6361 行那段注释就是这么教的），所以 Init 按
        // MonthRows×MonthColumns 把矩形乘出来、交给 MCM_SIZERECTTOMIN 撑到真装得下，再挪一次窗。
        // 1×1（或 .frm 没写）时 RTL 直接返回，发码形状照旧稳定 —— emitc 断言要看这一条存在。
        if (ctrl.controlType == FrmControlType::MonthView) {
            std::string mvLower = ctrl.controlName;
            std::transform(mvLower.begin(), mvLower.end(), mvLower.begin(), ::tolower);
            std::string mvHwnd = knownControlArrays_.count(mvLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            int mvRows = 1, mvCols = 1, mvMax = -999;
            auto mvIt = ctrl.properties.find("MonthRows");
            if (mvIt != ctrl.properties.end()) mvRows = (int)mvIt->second.intValue;
            mvIt = ctrl.properties.find("MonthColumns");
            if (mvIt != ctrl.properties.end()) mvCols = (int)mvIt->second.intValue;
            mvIt = ctrl.properties.find("MaxSelCount");
            if (mvIt != ctrl.properties.end()) mvMax = (int)mvIt->second.intValue;
            if (mvRows < 1) mvRows = 1;
            if (mvCols < 1) mvCols = 1;
            c_.emitLine("vb6_MV_Init((void*)" + mvHwnd + ", "
                        + std::to_string(mvRows) + ", " + std::to_string(mvCols) + ", "
                        + std::to_string(mvMax) + ");");
        }

        // C29-RT-a: RichTextBox 设计期。ScrollBars 已经立在创建参数里（事后写没有量程），
        // BorderStyle 由上面那条通用 emitDesignerStyleProps 管，WordWrap 原生没有 Get 对称项
        // ⇒ RTL 自存窗口属性。-999 = .frm 没写（VB6 布尔的 True 就是 -1，不能拿它当哨兵）。
        // 未写时这条照发（RTL 直接返回）—— 发码形状稳定，emitc 断言要看它。
        if (ctrl.controlType == FrmControlType::RichTextBox) {
            std::string rtLower = ctrl.controlName;
            std::transform(rtLower.begin(), rtLower.end(), rtLower.begin(), ::tolower);
            std::string rtHwnd = knownControlArrays_.count(rtLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            int rtWW = -999, rtRO = -999;
            auto wwIt = ctrl.properties.find("WordWrap");
            if (wwIt != ctrl.properties.end()) rtWW = (wwIt->second.intValue != 0) ? 1 : 0;
            auto roIt = ctrl.properties.find("ReadOnly");
            if (roIt != ctrl.properties.end()) rtRO = (roIt->second.intValue != 0) ? 1 : 0;
            c_.emitLine("vb6_RTB_Init((void*)" + rtHwnd + ", " + std::to_string(rtWW)
                        + ", " + std::to_string(rtRO) + ");");
        }

        // C29-WS-a: Winsock 设计期。这枚控件的一切状态都在 RTL 的实例表里（socket / 状态机 /
        // 接收缓冲），所以这里只做三件事：把窗口认进表、灌 .frm 上写的那几条、把存在的事件
        // 处理器注册成函数指针 —— 网络通知走不了 WM_COMMAND / WM_NOTIFY（那是"往父窗投"那一路），
        // 只能是 Timer / Data / OLEDrop 那种"注册回调"的形状（029 §九 C29-WS 记了为什么）。
        if (ctrl.controlType == FrmControlType::Winsock) {
            std::string wsLower = ctrl.controlName;
            std::transform(wsLower.begin(), wsLower.end(), wsLower.begin(), ::tolower);
            std::string wsHwnd = knownControlArrays_.count(wsLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            auto wsInt = [&](const char* k, int dflt) -> int {
                auto it = ctrl.properties.find(k);
                return (it == ctrl.properties.end()) ? dflt : (int)it->second.intValue;
            };
            auto wsStr = [&](const char* k) -> std::string {
                auto it = ctrl.properties.find(k);
                if (it == ctrl.properties.end() || it->second.rawText.empty()) return "L\"\"";
                std::string raw = it->second.rawText;
                if (raw.size() >= 2 && raw.front() == '\"' && raw.back() == '\"')
                    raw = raw.substr(1, raw.size() - 2);
                return "L\"" + escapeWideCString(raw) + "\"";
            };
            c_.emitLine("vb6_Ws_Create((void*)" + wsHwnd + ");");
            c_.emitLine("vb6_Ws_SetProtocol((void*)" + wsHwnd + ", "
                        + std::to_string(wsInt("Protocol", 0)) + ");");
            c_.emitLine("vb6_Ws_SetRemoteHost((void*)" + wsHwnd + ", " + wsStr("RemoteHost") + ");");
            c_.emitLine("vb6_Ws_SetRemotePort((void*)" + wsHwnd + ", "
                        + std::to_string(wsInt("RemotePort", 0)) + ");");
            // 事件槽：序号是 vb6forms_prop_ctrl.h 里那组 VB6_WS_EV_*，发的是字面量。
            // 处理器形参表**照 VB6 的 ByVal 签名发**（ByVal ⇒ 原样传值；这条在 MV-c 那格
            // 用 x86 才验出来过：形状写错时 x64 读数照旧正确）。
            struct WsEvt { const char* suffix; int kind; const char* proto; };
            static const WsEvt wsEvts[] = {
                { "_Connect",            0, "" },
                { "_Close",              1, "" },
                { "_DataArrival",        2, "int32_t" },
                { "_ConnectionRequest",  3, "int32_t" },
                { "_SendComplete",       4, "" },
                { "_SendProgress",       5, "int32_t" },
                { "_StateChanged",       6, "int16_t" },
                // Error 那一形七参数，形状照生成的处理器原型（--emit-c 实测）：
                //   void f(int16_t, BSTR*, int32_t, BSTR, BSTR, int32_t, int16_t*)
                // ByRef 的两格（Description / CancelDisplay）是指针 —— 写错在 x64 上看不出来，
                // 只有 x86 会当场错值（本线 MV-c 那条纪律）。
                { "_Error",              7, "int16_t, BSTR*, int32_t, BSTR, BSTR, int32_t, int16_t*" }
            };
            for (const WsEvt& ev : wsEvts) {
                auto* sy = symTab_.lookup(ctrl.controlName + ev.suffix);
                if (!sy) continue;
                std::string evFn = cProcName(sy->name, AccessLevel::Private);
                std::string argDecl = std::string(ev.proto);
                c_.emitLine("{ extern void " + evFn + "("
                            + (argDecl.empty() ? std::string("void") : argDecl) + ");");
                c_.emitLine("  vb6_Ws_SetEventHandler((void*)" + wsHwnd + ", "
                            + std::to_string(ev.kind) + ", (void*)" + evFn + "); }");
            }
        }

        // P20-40: StatusBar 设计期 —— Init 一次, 再按 Panels 集合逐项灌进去。
        // 未写进 .frm 的属性走 VB6 默认 (Align=2 靠下 / Style=0 多面板 / 无面板)。
        if (ctrl.controlType == FrmControlType::StatusBar) {
            // 登记裸名: 表达式侧的 statusBarNameOfExpr 是把 `vb6_ComGetObjectProp(vb6_com_X,
            // L"Panels")` 抠成裸名 X 之后比对的, 存槽名或带前缀的名字都匹配不上。
            statusBarVars_.insert(ctrl.controlName);
            std::string sbNLower = ctrl.controlName;
            std::transform(sbNLower.begin(), sbNLower.end(), sbNLower.begin(), ::tolower);
            std::string sbHwnd = knownControlArrays_.count(sbNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));

            // 取字符串字面量: FrmValue::rawText 带引号, 剥掉后转义即可直接拼进 C 源码。
            auto litOf = [](const std::map<std::string, FrmValue>& mp, const std::string& k) {
                auto it = mp.find(k);
                if (it == mp.end() || it->second.type != FrmValueType::String) return std::string();
                std::string raw = it->second.rawText;
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
                    raw = raw.substr(1, raw.size() - 2);
                for (char& ch : raw) if (ch == '"') ch = '\'';
                return raw;
            };
            auto intOf = [](const std::map<std::string, FrmValue>& mp, const std::string& k) -> int {
                auto it = mp.find(k);
                return (it != mp.end()) ? (int)it->second.intValue : -1;
            };

            auto alIt = ctrl.properties.find("Align");
            int sbAlign = (alIt != ctrl.properties.end() && alIt->second.intValue >= 0)
                        ? (int)alIt->second.intValue : 2;
            int sbStyle = (intOf(ctrl.properties, "Style") == 1) ? 1 : 0;
            c_.emitLine("vb6_StatusBar_Init((void*)" + sbHwnd + ", "
                      + std::to_string(sbStyle) + ", " + std::to_string(sbAlign) + ");");

            if (sbStyle == 1) {   // sbrSimple: 只有一格, 内容走 SimpleText
                std::string stx = litOf(ctrl.properties, "SimpleText");
                c_.emitLine("vb6_StatusBar_SetSimpleText((void*)" + sbHwnd + ", "
                          + (stx.empty() ? std::string("L\"\"") : ("L\"" + stx + "\"")) + ");");
            }

            for (const auto& block : ctrl.propertyBlocks) {
                // 两种块名都认: 有裸标题行时取标题, 没有(StatusBar 的 Panels 就是这种)
                // 时解析器兜底成 "Items"。
                if (block.blockName != "Panels" && block.blockName != "Items") continue;
                int pOrd = 0;
                for (const auto& p : block.nestedBlocks) {
                    pOrd++;
                    // 下标没写就按出现顺序 (VB6 的 `Panels(1)/(2)` 一般都会写)。
                    int idx = intOf(p.properties, "Index");
                    if (idx <= 0) idx = pOrd;
                    std::string key = litOf(p.properties, "Key");
                    std::string txt = litOf(p.properties, "Text");
                    c_.emitLine("vb6_StatusBar_AddPanel((void*)" + sbHwnd + ", "
                              + std::to_string(idx) + ", "
                              + (key.empty() ? std::string("NULL") : ("L\"" + key + "\"")) + ", "
                              + "L\"" + txt + "\");");
                    // 子属性只在偏离 VB6 默认值时才发, 省掉一堆无意义的调用。
                    int pv = intOf(p.properties, "Style");
                    if (pv > 0) c_.emitLine("vb6_StatusBar_SetPanelStyle((void*)" + sbHwnd
                                          + ", " + std::to_string(idx) + ", " + std::to_string(pv) + ");");
                    pv = intOf(p.properties, "AutoSize");
                    if (pv > 0) c_.emitLine("vb6_StatusBar_SetPanelAutoSize((void*)" + sbHwnd
                                          + ", " + std::to_string(idx) + ", " + std::to_string(pv) + ");");
                    // Width/MinWidth 的 VB6 默认就是 0, 只在**显式给了正数**时才发;
                    // 用 `!= 0` 会把"没写"误判成 -1 灌进去。
                    pv = intOf(p.properties, "Width");
                    if (pv > 0) c_.emitLine("vb6_StatusBar_SetPanelWidth((void*)" + sbHwnd
                                          + ", " + std::to_string(idx) + ", " + std::to_string(pv) + ");");
                    pv = intOf(p.properties, "MinWidth");
                    if (pv > 0) c_.emitLine("vb6_StatusBar_SetPanelMinWidth((void*)" + sbHwnd
                                          + ", " + std::to_string(idx) + ", " + std::to_string(pv) + ");");
                    std::string tip = litOf(p.properties, "ToolTipText");
                    if (!tip.empty()) c_.emitLine("vb6_StatusBar_SetPanelToolTip((void*)" + sbHwnd
                                          + ", " + std::to_string(idx) + ", L\"" + tip + "\");");
                }
                // **不要 break**: 一个控件里可以有多个同名的集合块 (StatusBar 的 Panels
                // 一项一个块, 块名都是解析器兜底的 "Items"), break 会把第二项之后的
                // 面板整个丢掉, 且不报任何错。
            }
        }

        // C29-5a: Toolbar 设计期 —— Init 灌四条标量属性, 再把 Buttons 集合逐条建进
        // 原生控件。未写进 .frm 的属性传 -999 = "不动" (口径与 C29-8a 一致: VB6 的
        // True 就是 -1, 拿 -1 当哨兵等于 ShowTips 永远关不掉)。
        if (ctrl.controlType == FrmControlType::Toolbar) {
            toolbarVars_.insert(ctrl.controlName);   // 表达式侧认 `tb1.Buttons` 用
            std::string tbNLower = ctrl.controlName;
            std::transform(tbNLower.begin(), tbNLower.end(), tbNLower.begin(), ::tolower);
            std::string tbHwnd = knownControlArrays_.count(tbNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));

            auto tbLit = [](const std::map<std::string, FrmValue>& mp, const std::string& k) {
                auto it = mp.find(k);
                if (it == mp.end() || it->second.type != FrmValueType::String) return std::string();
                std::string raw = it->second.rawText;
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
                    raw = raw.substr(1, raw.size() - 2);
                for (char& ch : raw) if (ch == '"') ch = '\'';
                return raw;
            };
            auto tbInt = [](const std::map<std::string, FrmValue>& mp, const std::string& k) -> int {
                auto it = mp.find(k);
                return (it != mp.end()) ? (int)it->second.intValue : -1;
            };

            int tbTips = -999, tbCust = -999;
            int tbText  = tbInt(ctrl.properties, "TextStyle");
            int tbAlign = tbInt(ctrl.properties, "Align");
            // 两条布尔归一到 VB6 的 -1/0; 没写在 .frm 里就保持 -999 (不动窗口)。
            {
                auto raw = ctrl.properties.find("ShowTips");
                tbTips = (raw == ctrl.properties.end()) ? -999
                       : ((raw->second.intValue != 0) ? -1 : 0);
                raw = ctrl.properties.find("AllowCustomize");
                tbCust = (raw == ctrl.properties.end()) ? -999
                       : ((raw->second.intValue != 0) ? -1 : 0);
            }
            c_.emitLine("vb6_Toolbar_Init((void*)" + tbHwnd + ", "
                      + std::to_string(tbTips) + ", " + std::to_string(tbText < 0 ? -999 : tbText) + ", "
                      + std::to_string(tbCust) + ", " + std::to_string(tbAlign < 0 ? -999 : tbAlign) + ");");

            for (const auto& block : ctrl.propertyBlocks) {
                // 裸标题行 `Buttons` 时块名取标题; 一项一块的形态下解析器兜底成 "Items"。
                if (block.blockName != "Buttons" && block.blockName != "Items") continue;
                int ord = 0;
                for (const auto& b : block.nestedBlocks) {
                    ord++;
                    int idx = tbInt(b.properties, "Index");
                    if (idx <= 0) idx = ord;
                    std::string cap  = tbLit(b.properties, "Text");        // 项的裸标题 = Caption
                    std::string key  = tbLit(b.properties, "Key");
                    std::string tip  = tbLit(b.properties, "ToolTipText");
                    int st  = tbInt(b.properties, "Style");   if (st  < 0) st = 0;
                    int img = tbInt(b.properties, "Image");   if (img < 0) img = -1;
                    int wd  = tbInt(b.properties, "Width");   if (wd  < 0) wd = 0;
                    c_.emitLine("vb6_Toolbar_AddButton((void*)" + tbHwnd + ", "
                              + std::to_string(idx) + ", "
                              + (key.empty() ? std::string("NULL") : ("L\"" + key + "\"")) + ", "
                              + "L\"" + cap + "\", "
                              + std::to_string(st) + ", " + std::to_string(img) + ", "
                              + (tip.empty() ? std::string("NULL") : ("L\"" + tip + "\"")) + ", "
                              + std::to_string(wd) + ");");
                }
                // 不 break: 一项一个块, 块名可能都是兜底的 "Items"。
            }
        }

        // C29-7: ListView 设计期 —— View/GridLines/FullRowSelect/MultiSelect/CheckBoxes/
        // HideColumnHeaders/AllowColumnReorder/LabelEdit 一次 Init 灌进去。
        // 不灌就是"空壳": comctl32 的 SysListView32 默认是**图标视图**, 于是设计期写了
        // View=3 (报表) 也照样看不见列 —— 实测读数 LV13-VIEW=0 而不是 3。
        if (ctrl.controlType == FrmControlType::ListView) {
            std::string lvNLower = ctrl.controlName;
            std::transform(lvNLower.begin(), lvNLower.end(), lvNLower.begin(), ::tolower);
            std::string lvHwnd = knownControlArrays_.count(lvNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            auto lvInt = [](const std::map<std::string, FrmValue>& mp, const std::string& k, int defv) -> int {
                auto it = mp.find(k);
                if (it == mp.end()) return defv;
                return (int)it->second.intValue;
            };
            // VB6 的 True 是 -1; 这些布尔属性按"非 0 即真"归一化后再发。
            // 账 #128-b: 归一化成 **-1** 而不是 1 —— 与其余控件那几族(TreeView/MonthView/Slider)
            // 发的那一档对齐，也让 RTL 里"存进去就是 VB6 的那一个数"这条不变量成立。
            auto lvBool = [&](const std::string& k) -> int {
                return lvInt(ctrl.properties, k, 0) != 0 ? -1 : 0;
            };
            int lvView = lvInt(ctrl.properties, "View", 0);
            if (lvView < 0 || lvView > 3) lvView = 0;
            c_.emitLine("vb6_ListView_Init((void*)" + lvHwnd + ", "
                      + std::to_string(lvView) + ", "
                      + std::to_string(lvBool("GridLines")) + ", "
                      + std::to_string(lvBool("FullRowSelect")) + ", "
                      + std::to_string(lvBool("MultiSelect")) + ", "
                      + std::to_string(lvBool("CheckBoxes")) + ", "
                      + std::to_string(lvBool("HideColumnHeaders")) + ", "
                      + std::to_string(lvBool("AllowColumnReorder")) + ", "
                      + std::to_string(lvBool("LabelEdit")) + ");");
        }

        // C29-Data: Data 控件设计期 —— DatabaseName/RecordSource/Connect 一次 Init 灌进去,
        // Refresh 时才真开 (ODBC 连接是懒的)。随后:
        //   ① 绑定收集: 全窗体 TextBox 的 DataSource/DataField → vb6_Data_Bind
        //   ② Reposition 事件: handler 存在才发函数指针注册
        if (ctrl.controlType == FrmControlType::Data) {
            std::string dtNLower = ctrl.controlName;
            std::transform(dtNLower.begin(), dtNLower.end(), dtNLower.begin(), ::tolower);
            std::string dtHwnd = knownControlArrays_.count(dtNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            auto dtStr = [&](const std::string& k) -> std::string {
                auto it = ctrl.properties.find(k);
                if (it == ctrl.properties.end() || it->second.rawText.empty()) return "L""";
                std::string raw = it->second.rawText;
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
                    raw = raw.substr(1, raw.size() - 2);
                return "L\"" + escapeWideCString(raw) + "\"";
            };
            c_.emitLine("vb6_Data_Init((void*)" + dtHwnd + ", "
                      + dtStr("DatabaseName") + ", "
                      + dtStr("RecordSource") + ", "
                      + dtStr("Connect") + ");");
            // 绑定收集挪到 style_apply 循环外 (本控件段执行时被绑控件的 hwnd 可能还没创建)
            // ② Reposition (函数指针注册 —— 不可见控件没有 WM_NOTIFY 可借)
            if (symTab_.lookup(ctrl.controlName + "_Reposition")) {
                std::string rpFn = cProcName(ctrl.controlName + "_Reposition", AccessLevel::Private);
                c_.emitLine("vb6_Data_SetRepositionHandler((void*)" + dtHwnd + ", " + rpFn + ");");
            }
        }

        // C29-OLE: OLE 容器设计期 —— Class/OLETypeAllowed/SizeMode/DisplayAsIcon/
        // AutoActivate/AutoVerbMenu/BorderStyle/SourceDoc/SourceItem 一次 Init 灌进去。
        // 不灌的话 OLETypeAllowed 会错成 2 (vbOLEEither) 之外任意值、SizeMode 错成 0。
        if (ctrl.controlType == FrmControlType::OLE) {
            std::string ocNLower = ctrl.controlName;
            std::transform(ocNLower.begin(), ocNLower.end(), ocNLower.begin(), ::tolower);
            std::string ocHwnd = knownControlArrays_.count(ocNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));
            auto ocInt = [](const std::map<std::string, FrmValue>& mp, const std::string& k, int defv) -> int {
                auto it = mp.find(k);
                if (it == mp.end()) return defv;
                return (int)it->second.intValue;
            };
            auto ocStr = [&](const std::string& k) -> std::string {
                auto it = ctrl.properties.find(k);
                if (it == ctrl.properties.end() || it->second.rawText.empty()) return "L\"\"";
                // rawText 是含引号的原始文本 ("xxx"); 剥掉外层引号后交给转义助手
                std::string raw = it->second.rawText;
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
                    raw = raw.substr(1, raw.size() - 2);
                return "L\"" + escapeWideCString(raw) + "\"";
            };
            c_.emitLine("vb6_OleCon_Init((void*)" + ocHwnd + ", "
                      + ocStr("Class") + ", "
                      + std::to_string(ocInt(ctrl.properties, "OLETypeAllowed", 2)) + ", "
                      + std::to_string(ocInt(ctrl.properties, "SizeMode", 0)) + ", "
                      + std::to_string(ocInt(ctrl.properties, "DisplayAsIcon", 0) != 0 ? 1 : 0) + ", "
                      + std::to_string(ocInt(ctrl.properties, "AutoActivate", 2)) + ", "
                      + std::to_string(ocInt(ctrl.properties, "AutoVerbMenu", 1) != 0 ? 1 : 0) + ", "
                      + std::to_string(ocInt(ctrl.properties, "BorderStyle", 0)) + ", "
                      + ocStr("SourceDoc") + ", "
                      + ocStr("SourceItem") + ");");
        }

        // P20-42: SSTab 设计期 —— 一次 Init 灌进 Tabs/Tab/TabOrientation/TabStyle/
        // TabsPerRow/WordWrap, 再逐页灌 Caption。
        // 页标题的 .frm 形态支持两种: ①VB6 实际写的 BeginProperty Tabs {…} 块里
        // 逐个 BeginProperty TabN {…} 的 Caption; ②手写的 `TabCaption(0) = "…"` 行。
        if (ctrl.controlType == FrmControlType::SSTab) {
            std::string tsNLower = ctrl.controlName;
            std::transform(tsNLower.begin(), tsNLower.end(), tsNLower.begin(), ::tolower);
            std::string tsHwnd = knownControlArrays_.count(tsNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                 std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(ctrl.controlName));

            auto tsInt = [](const std::map<std::string, FrmValue>& mp, const std::string& k, int defv) -> int {
                auto it = mp.find(k);
                if (it == mp.end()) return defv;
                return (int)it->second.intValue;
            };
            auto tsLit = [](const std::map<std::string, FrmValue>& mp, const std::string& k) -> std::string {
                auto it = mp.find(k);
                if (it == mp.end() || it->second.type != FrmValueType::String) return std::string();
                std::string raw = it->second.rawText;
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
                    raw = raw.substr(1, raw.size() - 2);
                for (char& ch : raw) if (ch == '"') ch = '\'';
                return raw;
            };

            int tsTabs     = tsInt(ctrl.properties, "Tabs", 0);
            int tsCur      = tsInt(ctrl.properties, "Tab", 0);
            int tsOrient   = tsInt(ctrl.properties, "TabOrientation", 0);
            int tsStyle    = tsInt(ctrl.properties, "TabStyle", 0);
            int tsPerRow   = tsInt(ctrl.properties, "TabsPerRow", 0);
            // WordWrap 缺省 **True** (VB6 SSTab/SSTabEx 同): .frm 不写这行时 7 页
            // 折成 TabsPerRow 行; 缺省 0 会让多页折叠成单行+滚动箭头
            // (SSTabEx frmTest: 第二行 Label/Tab5/Tab6 不显示, 实证)。
            int tsWrap     = tsInt(ctrl.properties, "WordWrap", 1);
            if (tsTabs < 0) tsTabs = 0;
            c_.emitLine("vb6_SSTab_Init((void*)" + tsHwnd + ", "
                      + std::to_string(tsTabs) + ", " + std::to_string(tsCur) + ", "
                      + std::to_string(tsOrient) + ", " + std::to_string(tsStyle) + ", "
                      + std::to_string(tsPerRow) + ", " + std::to_string(tsWrap) + ");");

            // ①BeginProperty Tabs {…} → Tab1..TabN 的 Caption (VB6 实际形态)
            for (const auto& tb : ctrl.propertyBlocks) {
                if (tb.blockName != "Tabs") continue;
                int ord = 0;
                for (const auto& page : tb.nestedBlocks) {
                    ord++;
                    // TabN 的 N 是 1 基; 没写就按出现顺序 (VB6 一般都会写)。
                    int idx = tsInt(page.properties, "Index", ord - 1);
                    if (idx < 0) idx = 0;
                    std::string cap = tsLit(page.properties, "Caption");
                    if (cap.empty()) cap = tsLit(page.properties, "Text");
                    c_.emitLine("vb6_SSTab_SetTabCaption((void*)" + tsHwnd + ", "
                              + std::to_string(idx) + ", L\"" + cap + "\");");
                }
            }
            // ②`TabCaption(n) = "…"` 直接写在属性行上 (手写 .frm 的形态)
            for (const auto& kv : ctrl.properties) {
                const std::string& k = kv.first;
                if (k.size() < 12) continue;
                if (_strnicmp(k.c_str(), "TabCaption", 10) != 0) continue;
                size_t lp = k.find('(');
                size_t rp = k.find(')');
                if (lp == std::string::npos || rp == std::string::npos || rp <= lp + 1) continue;
                int idx = atoi(k.substr(lp + 1, rp - lp - 1).c_str());
                std::string raw = kv.second.type == FrmValueType::String ? kv.second.rawText : std::string();
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
                    raw = raw.substr(1, raw.size() - 2);
                for (char& ch : raw) if (ch == '"') ch = '\'';
                c_.emitLine("vb6_SSTab_SetTabCaption((void*)" + tsHwnd + ", "
                          + std::to_string(idx) + ", L\"" + raw + "\");");
            }

            // ③TabPic16/20/24(n) = "frmTest.frx":xxxx —— SSTabEx 的每页标签图标
            // (Task #44 参考截图: Theme/Frame/Other/Cmd 四页有图标, 其余页没有)。
            // 三档尺寸是同一张图 (16/20/24px, .frx 条目与 ImageList 同构: StdPicture
            // GUID + 'lt' magic + 长度 → readPicture 走 isImageList=true 的 28B 头剥离)。
            // 96DPI 优先取 16 档, 缺档顺延 20/24 (RTL 侧 imagelist 固定 16x16)。
            if (frxLoaded && tsTabs > 0) {
                auto tsPicVal = [&](int px, int idx2) -> const FrmValue* {
                    auto it = ctrl.properties.find(
                        "TabPic" + std::to_string(px) + "(" + std::to_string(idx2) + ")");
                    if (it == ctrl.properties.end()) return nullptr;
                    if (it->second.type != FrmValueType::FrxReference) return nullptr;
                    return &it->second;
                };
                for (int idx2 = 0; idx2 < tsTabs; idx2++) {
                    const FrmValue* pv = tsPicVal(16, idx2);
                    if (!pv) pv = tsPicVal(20, idx2);
                    if (!pv) pv = tsPicVal(24, idx2);
                    if (!pv) continue;
                    auto pic = FrxReader::readPicture(pv->frxOffset, true);
                    if (pic.data.empty()) continue;
                    std::string varName = "vb6_frx_tabpic_" + cIdent(ctrl.controlName)
                                        + "_" + std::to_string(idx2);
                    h_.emitLine(bytesToHexArray(pic.data.data(), pic.data.size(), varName));
                    c_.emitLine("vb6_SSTab_SetTabPicture((void*)" + tsHwnd + ", "
                              + std::to_string(idx2) + ", " + varName + ", "
                              + varName + "_size);");
                }
            }
        }

        // P20-44: OLE 拖放目标侧 —— OLEDropMode<>0 时注册为放置目标。
        // 用户写的 <Ctl>_OLEDragOver/_OLEDragDrop **直接注册成 RTL 回调**:
        // C3 给它生成的签名 `void X(void** Data, int32_t*, int16_t*, int16_t*, float*, float*)`
        // 与 RTL 的 Vb6OLEDropCb 完全一致, 不需要 thunk。
        // 块内 extern 声明即可 (处理器的定义在本文件后面), 不用动头文件。
        {
            auto odmIt = ctrl.properties.find("OLEDropMode");
            if (odmIt != ctrl.properties.end() && odmIt->second.intValue != 0) {
                std::string odNLower = ctrl.controlName;
                std::transform(odNLower.begin(), odNLower.end(), odNLower.begin(), ::tolower);
                std::string odHwnd = knownControlArrays_.count(odNLower) ?
                    ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", 0)") :
                    ("vb6_hwnd_" + cIdent(ctrl.controlName));
                c_.emitLine("vb6_OLEDrop_Register((void*)" + odHwnd + ");");
                static const char* odEv43[2] = { "_OLEDragOver", "_OLEDragDrop" };
                for (int odK = 0; odK < 2; odK++) {
                    auto* odSym = symTab_.lookup(ctrl.controlName + odEv43[odK]);
                    if (!odSym) continue;
                    std::string odFn = cProcName(odSym->name, AccessLevel::Private);
                    c_.emitLine("{ extern void " + odFn + "(void**, int32_t*, int16_t*, int16_t*, float*, float*);");
                    c_.emitLine("  vb6_OLEDrop_SetHandler((void*)" + odHwnd + ", "
                              + std::to_string(odK) + ", (void*)" + odFn + "); }");
                }
            }
        }

        // D6 / C29-9: CommonDialog 的设计期四行（VB6 会把用户填的属性原样存进 .frm）。
        // 走的是刚建好的那枚不可见句柄，所以顺序上必须排在 vb6_hwnd_<名> 赋值之后 ——
        // 这里就是那条不变式（同 C29-1a 的 emitShapeLineProps 一个位置）。
        if (ctrl.controlType == FrmControlType::CommonDialog) {
            std::string cdHwnd = "vb6_hwnd_" + cIdent(ctrl.controlName);
            auto cdStr = [&](const char* key, const char* fn) {
                auto it = ctrl.properties.find(key);
                if (it == ctrl.properties.end() || it->second.type != FrmValueType::String) return;
                std::string raw = it->second.rawText;
                if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"')
                    raw = raw.substr(1, raw.size() - 2);
                c_.emitLine(std::string(fn) + "((void*)" + cdHwnd + ", vb6_BSTR_FromStr(L\""
                            + escapeWideCString(raw) + "\"));");
            };
            auto cdInt = [&](const char* key, const char* fn) {
                auto it = ctrl.properties.find(key);
                if (it == ctrl.properties.end()) return;
                c_.emitLine(std::string(fn) + "((void*)" + cdHwnd + ", "
                            + std::to_string((int)it->second.intValue) + ");");
            };
            cdStr("Filter",      "vb6_CdSetFilter");
            cdStr("DialogTitle", "vb6_CdSetDialogTitle");
            cdStr("InitDir",     "vb6_CdSetInitDir");
            cdStr("DefaultExt",  "vb6_CdSetDefaultExt");
            cdStr("FileName",    "vb6_CdSetFileName");
            cdStr("FontName",    "vb6_CdSetFontName");
            cdInt("Flags",       "vb6_CdSetFlags");
            cdInt("CancelError", "vb6_CdSetCancelError");
            cdInt("Color",       "vb6_CdSetColor");
            cdInt("Min",         "vb6_CdSetMin");
            cdInt("Max",         "vb6_CdSetMax");
            cdInt("Copies",      "vb6_CdSetCopies");
            cdInt("FontSize",    "vb6_CdSetFontSize");
        }
        // P24 / Fix 182: .frx 设计期属性（Picture / TextBox.Text / ListBox.List）。
        // 实现已抽到 cgen_form_create_controls.inc 的 emitControlFrxProps, 与容器子控件共用。
        emitControlFrxProps(ctrl);

ctrlId++;
    }

    // C29-Data: 绑定收集独立 pass —— 必须在**所有**控件创建完之后 (被绑的
    // TextBox 等 hwnd 此时才有效; per-ctrl 段里发射时目标 hwnd 还是 NULL,
    // vb6_Data_Bind 静默丢弃, 实测 DT6 绑定读数空)。
    for (const auto& dataCtrl : frmDesc.formControl.children) {
        if (dataCtrl.controlType != FrmControlType::Data) continue;
        std::string dNLower = dataCtrl.controlName;
        std::transform(dNLower.begin(), dNLower.end(), dNLower.begin(), ::tolower);
        std::string dHwnd = knownControlArrays_.count(dNLower) ?
            ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(dataCtrl.controlName) + ", " +
             std::to_string(dataCtrl.index >= 0 ? dataCtrl.index : 0) + ")") :
            ("vb6_hwnd_" + cIdent(dataCtrl.controlName));
        for (const auto& other : frmDesc.formControl.children) {
            auto dsIt = other.properties.find("DataSource");
            auto dfIt = other.properties.find("DataField");
            if (dsIt == other.properties.end() || dfIt == other.properties.end()) continue;
            std::string ds = dsIt->second.rawText;
            if (ds.size() >= 2 && ds.front() == '"' && ds.back() == '"')
                ds = ds.substr(1, ds.size() - 2);
            if (ds != dataCtrl.controlName) continue;
            std::string fldRaw = dfIt->second.rawText;
            if (fldRaw.size() >= 2 && fldRaw.front() == '"' && fldRaw.back() == '"')
                fldRaw = fldRaw.substr(1, fldRaw.size() - 2);
            std::string oNLower = other.controlName;
            std::transform(oNLower.begin(), oNLower.end(), oNLower.begin(), ::tolower);
            std::string oHwnd = knownControlArrays_.count(oNLower) ?
                ("vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(other.controlName) + ", " +
                 std::to_string(other.index >= 0 ? other.index : 0) + ")") :
                ("vb6_hwnd_" + cIdent(other.controlName));
            c_.emitLine("vb6_Data_Bind((void*)" + dHwnd + ", (void*)" + oHwnd
                      + ", L\"" + escapeWideCString(fldRaw) + "\");");
        }
    }

