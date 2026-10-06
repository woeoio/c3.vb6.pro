// cgen_form_frame_menu.inc - emitFormFramework 函数体片段：Frame 子控件 / 控件子类化安装 / Win32 菜单构建
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1773~1937 行，逐行未改。

    // Task #41: 控件数组成员的句柄形态 — 声明侧 (emitControlHandleDecls) 对
    // knownControlArrays_ 命中者只发 static vb6_CtrlArr vb6_arr_<名>, 不发
    // vb6_hwnd_<名>。此前创建/引用一律直写 vb6_hwnd_<名> → C2065
    // (SSTabEx frmTest: SSTabEx1 页容器 Picture2 是数组 → 4 处 C2065)。
    // 创建改 vb6_CtrlArr_SetAt, 引用改 vb6_CtrlArr_GetAt (与
    // cgen_form_create_controls.inc 的 UserControl 数组形态一致)。
    // 定义在函数体层级: P14.4.2 容器子控件与 P18-F 子类化两块共用。
    auto childHwndRead = [&](const FrmControl& c) -> std::string {
        std::string cLower = Symbol::toLower(c.controlName);
        if (knownControlArrays_.count(cLower)) {
            return "vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(c.controlName) + ", "
                 + std::to_string(c.index >= 0 ? c.index : 0) + ")";
        }
        return "vb6_hwnd_" + cIdent(c.controlName);
    };

    // P14.4.2: 容器控件(Frame/PictureBox) - 创建容器的子控件(以容器 HWND 为父窗口)
    // 关键: .frm 中子控件坐标已是"相对容器的缇值", 且 vb6_CreateControl 内部会做缇→像素换算,
    //       因此这里必须直传原始缇值 —— 既不能减去容器偏移, 也不能再乘 15/10.
    {
        int childCtrlId = 200;  // 容器子控件ID起始值(避免与顶层 100-199 冲突)
        auto emitChildControls = [&](auto&& self, const FrmControl& container,
                                     const std::string& containerHwnd) -> void {
            // Task #44 (SSTabEx): `Tab(n).Control(i) = "名(下标)"` 是 VB6 IDE 写进 .frm
            // 的**显式页归属**。SSTabEx 这类 Tabbed 控件把非活动页子控件**统一**停在
            // real-75000 (与页码无关) —— 纯靠 Left 反推页码必错 (frmTest: 全推成 1,
            // 真值 0/2/3/4)。先收集归属表, 键 = 小写 "名" 或 "名(下标)"。
            std::map<std::string, int> tabMember44;
            if (container.controlType == FrmControlType::SSTab) {
                for (const auto& kv44 : container.properties) {
                    const std::string& k44 = kv44.first;
                    if (_strnicmp(k44.c_str(), "Tab(", 4) != 0) continue;
                    size_t cl44 = k44.find(')');
                    if (cl44 == std::string::npos
                        || _strnicmp(k44.c_str() + cl44 + 1, ".Control(", 9) != 0) continue;
                    int pg44 = atoi(k44.c_str() + 4);
                    if (kv44.second.type != FrmValueType::String) continue;
                    std::string v44 = kv44.second.rawText;
                    if (v44.size() >= 2 && v44.front() == '"' && v44.back() == '"')
                        v44 = v44.substr(1, v44.size() - 2);
                    if (v44.empty()) continue;
                    std::transform(v44.begin(), v44.end(), v44.begin(), ::tolower);
                    tabMember44[v44] = pg44;
                }
            }
            // 账 #160: 容器里的 TabIndex 是**这套号自己从 0 编**（VB6 就是这样编号 —— 框架里的
            // 控件不占窗体那一串的号）。没写的那几枚用同一父窗内的声明序号兜底。
            long childTi = -1;
            for (const auto& child : container.children) {
                childTi++;
                if (child.controlType == FrmControlType::Menu) continue;
                const char* childClass = FrmParser::controlTypeToWin32Class(child.controlType);
                if (!childClass) continue;
                int cLeft = 0, cTop = 0, cWidth = 2000, cHeight = 300;
                auto clIt = child.properties.find("Left");
                if (clIt != child.properties.end()) cLeft = (int)clIt->second.intValue;
                auto ctIt = child.properties.find("Top");
                if (ctIt != child.properties.end()) cTop = (int)ctIt->second.intValue;
                auto cwIt = child.properties.find("Width");
                if (cwIt != child.properties.end()) cWidth = (int)cwIt->second.intValue;
                auto chIt = child.properties.find("Height");
                if (chIt != child.properties.end()) cHeight = (int)chIt->second.intValue;
                // ---- P20-42: SSTab 容器 —— 页码只能从 Left 反推 ----
                // VB6 把每页的子控件摆在同一份坐标里, 靠把**非活动页**子控件的 Left
                // **减 75000 缇**推出可见区来隐藏 (微软 KB Q187562 / Q206904 明写)。
                // .frm 里因此没有任何页号字段, 只能反推:
                //       storedLeft = realLeft - 75000 * page
                // 取让 realLeft 落回 [0, 75000) 的那个 page。阈值取 -37500 而不是 0,
                // 是为了不把本来就有小负值 Left 的控件误判到下一页。
                // 反推后**必须把 Left 改回真值**, 否则子控件会被摆到屏幕外 —— 这点
                // 和 VB6 不同: VB6 运行时真的靠挪 Left 来切页, 我们改成登记页号 +
                // ShowWindow, Left 保持设计期值, `Label1.Left` 才读得对。
                int tsPage = 0;
                if (container.controlType == FrmControlType::SSTab && cLeft < -37500) {
                    tsPage = (int)(((long)(-cLeft) + 37500) / 75000);
                    if (tsPage < 0) tsPage = 0;
                    cLeft += 75000 * tsPage;
                }
                // Task #44: 显式 Tab(n).Control(i) 归属优先于 Left 反推。
                // Left 公式对"统一停 -75000"的 SSTabEx 恰好也能还原坐标 (恒得 1 页),
                // 但**页号**必须以 .frm 声明为准。
                if (container.controlType == FrmControlType::SSTab && !tabMember44.empty()) {
                    std::string mk44 = Symbol::toLower(child.controlName);
                    if (child.index >= 0) mk44 += "(" + std::to_string(child.index) + ")";
                    auto mt44 = tabMember44.find(mk44);
                    if (mt44 != tabMember44.end()) tsPage = mt44->second;
                }
                // C29-1a: Line 子控件的矩形 = 四端点的包围盒 (容器缇值)。不这样算会拿
                // 默认的 2000x300 盖在兄弟控件上面 (顶层那条路同一个口径)。
                if (child.controlType == FrmControlType::Line) {
                    lineRectFromEndpoints(child, cLeft, cTop, cWidth, cHeight);
                }
                long childStyle = 0x40000000L | 0x10000000L;  // WS_CHILD | WS_VISIBLE
                childStyle |= controlTypeStyleBits(child);  // 类型样式与顶层控件一致 (SS_NOTIFY/BS_GROUPBOX 等)
                childStyle |= controlTabStopStyleBit(child);  // 账 #83(a): TabStop 默认位与顶层同一处判断
                // 账 #158: 按钮族的焦点通知同样要这一位（两条创建路共用一个出口）
                childStyle |= controlButtonNotifyStyleBit(child);
                // Fix 162d-extlist: Caption 缺省 = "" 与顶层路径 (create_controls.inc
                // Fix 162c) 同口径 —— VB6 只在 Caption 非默认时写盘, 缺省回落空串。
                std::string childCaption;
                auto ccIt = child.properties.find("Caption");
                if (ccIt != child.properties.end() && ccIt->second.type == FrmValueType::String) {
                    std::string raw = ccIt->second.rawText;
                    if (raw.size() >= 2 && raw.front() == 0x22 && raw.back() == 0x22) childCaption = raw.substr(1, raw.size() - 2);
                }
                auto ctIt2 = child.properties.find("Text");
                if (ctIt2 != child.properties.end() && ctIt2->second.type == FrmValueType::String) {
                    std::string raw = ctIt2->second.rawText;
                    if (raw.size() >= 2 && raw.front() == 0x22 && raw.back() == 0x22) childCaption = raw.substr(1, raw.size() - 2);
                }
                if (controlTypeClearsCaption(child)) childCaption = "";
                {
                    std::string childLower41 = Symbol::toLower(child.controlName);
                    // 账 #83(b2): 扩展样式与顶层同出口 —— 容器嵌在容器里时，**这一条路**也得给它
                    // WS_EX_CONTROLPARENT，否则嵌在 Frame 里的 Frame 仍然走不进 tab 序。
                    std::string createArgs =
                        "\"" + std::string(childClass) + "\", \"" + childCaption + "\","
                        + std::to_string(childStyle) + "L, "
                        + std::to_string(controlContainerExStyleBit(child)) + "L,"
                        + std::to_string(cLeft) + ", " + std::to_string(cTop) + ", "
                        + std::to_string(cWidth) + ", " + std::to_string(cHeight) + ","
                        + std::to_string(childCtrlId) + ", " + containerHwnd + ", hInstance";
                    if (knownControlArrays_.count(childLower41)) {
                        // Task #41: 数组成员 — 句柄写进 vb6_arr_<名> (声明侧无
                        // vb6_hwnd_<名> 变量, 直赋值 C2065)
                        c_.emitLine("vb6_CtrlArr_SetAt(&vb6_arr_" + cIdent(child.controlName)
                            + ", " + std::to_string(child.index >= 0 ? child.index : 0)
                            + ", vb6_CreateControl(" + createArgs + "));");
                    } else {
                        c_.emitLine("vb6_hwnd_" + cIdent(child.controlName)
                            + " = vb6_CreateControl(" + createArgs + ");");
                    }
                }
                // SSTab: 登记"这个子控件属于第几页", 切页的显隐由 RTL 负责。
                if (container.controlType == FrmControlType::SSTab) {
                    c_.emitLine("vb6_SSTab_RegisterChild((void*)" + containerHwnd
                              + ", (void*)" + childHwndRead(child)
                              + ", " + std::to_string(tsPage) + ");");
                }
                // C29-1a: 设计期外观属性 —— Frame / PictureBox 里的 Shape/Line 以前一条都不发
                emitShapeLineProps(child, ctrlHwndExprForInit(child));
                // 设计期 BorderStyle / Alignment (容器子控件)
                emitDesignerStyleProps(child, childHwndRead(child));

                emitDesignerStateProps(child, childHwndRead(child));  // 账 #125: 与顶层同口径
                emitDesignerStringProps(child, childHwndRead(child));  // 账 #142: 同上（第二条创建路）
                // Fix 162d-extlist (VbEclipse 批次): 容器子控件此前**从不**走 style_apply
                // 那条路 —— Label2-4.AutoSize=-1 全部落空 (顶层路径 Fix 162c 只覆盖顶层控件)。
                // 这里按同一口径补发 Label 的 AutoSize: 度量文本宽度并 SetWindowPos。
                // 注: Enabled/Visible/Value 由上面那条 emitDesignerStateProps 发 (它认
                // CheckBox 的 Grayed(2), 并且先判 properties 的 type 是不是 Integer);
                // 合并 github/dev 时删掉了本批次原先重复发的那一次同参调用。
                {
                    auto asIt162d = child.properties.find("AutoSize");
                    if (asIt162d != child.properties.end() &&
                        asIt162d->second.type == FrmValueType::Integer &&
                        asIt162d->second.intValue != 0 &&
                        child.controlType == FrmControlType::Label) {
                        c_.emitLine("vb6_SetLabelAutoSize((void*)" + childHwndRead(child)
                                  + ", -1);  /* design AutoSize */");
                    }
                }

                emitDesignerFontProps(child, childHwndRead(child));    // 账 #154: 同上（字体两条也两路都发）
                emitDesignerTabIndexProp(child, childHwndRead(child), childTi);  // 账 #160: 第二条创建路同样要发这一句
                emitDesignerScaleModeProp(child, childHwndRead(child));  // 账 #197: 第二条创建路同样要发这一句
                emitControlFrxProps(child);  // Fix 182: 容器子控件同样加载 .frx 设计期图片/文本/列表
                // C29-SL-o（账 #83 的 Slider 半边）: 设计期值面以前只在顶层那条路发 ——
                // Frame 里的 Slider 光有窗口，Min/Max/Value 那一批整条丢掉。共用同一个出口，
                // 参数序与哨兵口径不写第二份。
                if (child.controlType == FrmControlType::Slider) {
                    emitSliderDesignTimeInit(child, childHwndRead(child));
                }
                // 注册控件名到knownFormControls_(供属性访问)
                {
                    std::string childNameLower = child.controlName;
                    std::transform(childNameLower.begin(), childNameLower.end(), childNameLower.begin(), ::tolower);
                    knownFormControls_[childNameLower] = child.controlType;
                    knownFormControlOriginalNames_[childNameLower] = child.controlName;
                }
                childCtrlId++;
                // 递归处理多层容器(Frame/PictureBox 内再嵌容器)
                if (!child.children.empty() && controlIsContainerType(child)) {
                    self(self, child, childHwndRead(child));   // 账 #167：容器清单收成一处
                }
            }
        };
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (ctrl.children.empty()) continue;
            if (!controlIsContainerType(ctrl)) continue;   // 账 #167
            emitChildControls(emitChildControls, ctrl, childHwndRead(ctrl));
        }
    }

    // P18-F: 安装控件子类化 (在所有控件创建之后)
    {
        // 收集需要子类化的控件名(去重)
        auto needsSubclass = [&](const FrmControl& ctrl) -> bool {
            // C29-SL-n（账 #141）: 这一趟以前自己抄了一份判据，且比发 arm 那一趟少了
            // `_DblClick` 与 `_Paint` ⇒ 只挂这两条之一的控件生成了子类过程却从不被 install。
            // 现在三份（发 arm / 装 / 拆）共用 controlNeedsSubclass 这一处。
            return controlNeedsSubclass(ctrl);
        };
        std::unordered_set<std::string> subEmitted;
        // 顶层控件
        for (const auto& ctrl : frmDesc.formControl.children) {
            std::string ctrlNameLower = ctrl.controlName;
            std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
            if (subEmitted.count(ctrlNameLower)) continue;
            subEmitted.insert(ctrlNameLower);
            if (!needsSubclass(ctrl)) continue;
            std::string subProcName = "vb6_ctrl_subproc_" + cIdent(ctrl.controlName);
            if (knownControlArrays_.count(ctrlNameLower)) {
                // P0-4修复: 控件数组遍历所有元素安装子类化
                c_.emitLine("{ int vb6_si; for (vb6_si = 0; vb6_si < VB6_CTRLARR_MAX; vb6_si++) {");
                c_.emitLine("  void* vb6_sub_hwnd = vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", vb6_si);");
                c_.emitLine("  if (vb6_sub_hwnd) vb6_InstallControlSubclass(vb6_sub_hwnd, (void*)" + subProcName + "); } }");
            } else {
                c_.emitLine("vb6_InstallControlSubclass((void*)vb6_hwnd_" + cIdent(ctrl.controlName) + ", (void*)" + subProcName + ");");
            }
        }
        // P0-5修复: 容器(Frame/PictureBox)内的子控件 (含多层嵌套)
        {
            auto subclassContainerChildren = [&](auto&& self, const FrmControl& container) -> void {
                for (const auto& child : container.children) {
                    std::string childNameLower = child.controlName;
                    std::transform(childNameLower.begin(), childNameLower.end(), childNameLower.begin(), ::tolower);
                    if (!subEmitted.count(childNameLower)) {
                        subEmitted.insert(childNameLower);
                        if (needsSubclass(child)) {
                            std::string subProcName = "vb6_ctrl_subproc_" + cIdent(child.controlName);
                            // Task #41: 数组成员用 GetAt (声明侧无 vb6_hwnd_<名>)
                            c_.emitLine("vb6_InstallControlSubclass((void*)" + childHwndRead(child) + ", (void*)" + subProcName + ");");
                        }
                    }
                    if (!child.children.empty() && controlIsContainerType(child)) {   // 账 #167
                        self(self, child);
                    }
                }
            };
            for (const auto& ctrl : frmDesc.formControl.children) {
                if (!ctrl.children.empty() && controlIsContainerType(ctrl)) {   // 账 #167
                    subclassContainerChildren(subclassContainerChildren, ctrl);
                }
            }
        }
    }

    // Fix 142: 容器控件转发子控件 WM_COMMAND 到窗体 WndProc.
    // 顶层与嵌套 Frame/PictureBox 均安装 (含仅作为容器的控件).
    {
        auto emitForward142 = [&](auto&& self, const FrmControl& ctrl) -> void {
            if (controlIsContainerType(ctrl) && !ctrl.children.empty()) {   // 账 #167
                std::string nmLower = ctrl.controlName;
                std::transform(nmLower.begin(), nmLower.end(), nmLower.begin(), ::tolower);
                if (!knownControlArrays_.count(nmLower)) {
                    c_.emitLine("vb6_ForwardChildCommands((void*)vb6_hwnd_" + cIdent(ctrl.controlName) + ");");
                }
            }
            for (const auto& child : ctrl.children) self(self, child);
        };
        for (const auto& ctrl : frmDesc.formControl.children) {
            emitForward142(emitForward142, ctrl);
        }
    }

    // P7.8: 菜单构建 (VB.Menu控件不创建窗口, 用Win32菜单API)
    {
        int menuId = 1000;  // 菜单项ID起始值 (控件ID用100-999)
        bool hasMenu = false;
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (ctrl.controlType == FrmControlType::Menu) {
                hasMenu = true;
                break;
            }
        }
        if (hasMenu) {
            c_.emitBlank();
            c_.emitLine("// P7.8: Menu construction");
            c_.emitLine("HMENU hMenuBar = CreateMenu();");

            for (const auto& ctrl : frmDesc.formControl.children) {
                if (ctrl.controlType != FrmControlType::Menu) continue;

                // Top-level menu item -> popup menu
                std::string caption = ctrl.controlName;
                auto capIt = ctrl.properties.find("Caption");
                if (capIt != ctrl.properties.end() && capIt->second.type == FrmValueType::String) {
                    std::string raw = capIt->second.rawText;
                    if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
                        caption = raw.substr(1, raw.size() - 2);
                    }
                }

                if (ctrl.children.empty()) {
                    c_.emitLine("AppendMenuW(hMenuBar, MF_STRING, " + std::to_string(menuId) + ", L\"" + escapeWideCString(caption) + "\");");
                    menuId++;
                } else {
                    std::string popupVar = "hPopup_" + cIdent(ctrl.controlName);
                    c_.emitLine("HMENU " + popupVar + " = CreatePopupMenu();");

                    for (const auto& child : ctrl.children) {
                        emitMenuItem(popupVar, child, menuId);
                    }

                    c_.emitLine("AppendMenuW(hMenuBar, MF_POPUP, (UINT_PTR)" + popupVar + ", L\"" + escapeWideCString(caption) + "\");");
                }
            }

            c_.emitLine("SetMenu((HWND)hwnd, hMenuBar);");
            c_.emitBlank();
        }
    }

    c_.emitLine("vb6_formLoading_" + cIdent(formName) + " = 0;  /* allow events after form init */");
    c_.dedent();
    c_.emitLine("}");  // CreateControls
    c_.emitBlank();
