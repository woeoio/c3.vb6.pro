// cgen_form_create_controls.inc - emitFormFramework 函数体片段：CreateControls —— 控件树 CreateWindow
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1272~1475 行，逐行未改。

    // --- 控件创建函数 ---
    c_.emitLine("// === " + formName + " CreateControls ===");
    c_.emitLine("void " + createFn + "(void* hwnd, void* hInstance) {");
    c_.indent();
    c_.emitLine("(void)hwnd; (void)hInstance;");
    c_.emitLine("vb6_formLoading_" + cIdent(formName) + " = 1;  /* block events during form init */");
    c_.emitLine("vb6_ResetControlId();");
    c_.emitBlank();

    // P7.6: 初始化控件数组
    for (const auto& kv : knownControlArrays_) {
        // Find the original control name (need non-lowered version)
        for (const auto& ctrl : frmDesc.formControl.children) {
            std::string ctrlNameLower = ctrl.controlName;
            std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
            if (ctrlNameLower == kv.first) {
                c_.emitLine("vb6_CtrlArr_Init(&vb6_arr_" + cIdent(ctrl.controlName) + ");");
                break;
            }
        }
    }
    c_.emitBlank();
    // P20-35: 注册Shape/Line自定义窗口类
    {
        bool hasShape = false, hasLine = false, hasCd = false, hasTmr = false, hasOle = false, hasData = false;
        bool hasWs = false;   // C29-WS-a
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (ctrl.controlType == FrmControlType::Shape) hasShape = true;
            if (ctrl.controlType == FrmControlType::Line) hasLine = true;
            // D6 / C29-9: CommonDialog 的属性宿主类要在这之前注册好，否则
            // CreateWindowExW 拿不到类 ⇒ 句柄 NULL ⇒ 又回到"静默空转"那个症状。
            if (ctrl.controlType == FrmControlType::CommonDialog) hasCd = true;
            if (ctrl.controlType == FrmControlType::Timer) hasTmr = true;  // C29-T
            // C29-OLE: 同理 —— VB6_OLECONTAINER 类先注册, 控件创建才有窗口。
            if (ctrl.controlType == FrmControlType::OLE) hasOle = true;
            // C29-Data: 同理 —— VB6_DATA 类先注册。
            if (ctrl.controlType == FrmControlType::Data) hasData = true;
            // C29-WS: 同理 —— VB6_WINSOCK 类要先注册。差别只在**这枚的 WndProc 有逻辑**
            // （WSAAsyncSelect 的 FD_* 投在它自己的窗口上），不是又一枚属性袋。
            if (ctrl.controlType == FrmControlType::Winsock) hasWs = true;
        }
        if (hasShape || hasLine) {
            c_.emitLine("vb6_RegisterShapeLineClasses((void*)hInstance);");
        }
        if (hasCd) {
            c_.emitLine("vb6_RegisterCommDialogClass((void*)hInstance);");
        }
        if (hasTmr) {
            c_.emitLine("vb6_RegisterTimerClass((void*)hInstance);");
        }
        if (hasOle) {
            c_.emitLine("vb6_RegisterOleConClass((void*)hInstance);");
        }
        if (hasData) {
            c_.emitLine("vb6_RegisterDataClass((void*)hInstance);");
        }
        if (hasWs) {
            c_.emitLine("vb6_RegisterWinsockClass((void*)hInstance);");
        }
    }
    // 为每个控件生成CreateWindow调用
    ctrlId = 100;

    // Fix 127/128: Property Set 型字体属性 —— 走 Fix 128 的"Get + 原地覆写字段"通道
    // (不调用 Set, 因其实现体末尾的 Refresh 会在 CreateWorks 阶段把图表置空)。
    // 环境变量 C3_FONT_SET=1 可临时打开旧的直接调用路径以便对照。
    const bool kApplyFontSet127 = ([] {
        const char* e = std::getenv("C3_FONT_SET");
        return e && *e && *e != '0';
    })();
    (void)kApplyFontSet127;

    // Fix 119: .frm 的 BeginProperty Font 块 → 构造 vb6_ComIface_Font 对象。
    // VB6 里每个控件实例都持有自己的 Font 对象, 且控件 InitProperties 常以
    // UserControl.Font 为基准 (如 TitleFont.Size = Font.Size + 8); 此前字体块被
    // 整体跳过, 所有控件都退回默认字体 → 字号/文字测量/整体布局全错。
    auto emitFrmFontObj = [&](const FrmPropertyBlock& blk, const std::string& varName) {
        std::string nm = "MS Sans Serif";
        double sz = 8.25;
        long weight = 400, charset = 0;
        long bBold = 0, bItalic = 0, bUnder = 0, bStrike = 0;
        auto it = blk.properties.find("Name");
        if (it != blk.properties.end()) {
            nm = it->second.rawText;
            if (nm.size() >= 2 && nm.front() == '"' && nm.back() == '"')
                nm = nm.substr(1, nm.size() - 2);
        }
        it = blk.properties.find("Size");
        if (it != blk.properties.end())
            sz = (it->second.type == FrmValueType::Float) ? it->second.floatValue
                                                          : (double)it->second.intValue;
        it = blk.properties.find("Weight");
        if (it != blk.properties.end()) weight = (long)it->second.intValue;
        it = blk.properties.find("Charset");
        if (it != blk.properties.end()) charset = (long)it->second.intValue;
        it = blk.properties.find("Bold");
        if (it != blk.properties.end()) bBold = (long)it->second.intValue;
        it = blk.properties.find("Italic");
        if (it != blk.properties.end()) bItalic = (long)it->second.intValue;
        it = blk.properties.find("Underline");
        if (it != blk.properties.end()) bUnder = (long)it->second.intValue;
        it = blk.properties.find("Strikethrough");
        if (it != blk.properties.end()) bStrike = (long)it->second.intValue;
        // 常规字重(400) 时 VB6 的 Bold 由 Weight>=700 推导
        if (weight >= 700) bBold = -1;

        c_.emitLine("{ vb6_ComIface_Font* " + varName
                  + " = (vb6_ComIface_Font*)vb6_UC_NewFont();");
        c_.emitLine("  if (" + varName + ") {");
        c_.emitLine("    " + varName + "->Name = vb6_BSTR_FromStr(L\"" + escapeWideCString(nm) + "\");");
        {
            char szBuf[64];
            snprintf(szBuf, sizeof(szBuf), "%.4g", sz);
            c_.emitLine("    " + varName + "->Size = " + std::string(szBuf) + "f;");
        }
        c_.emitLine("    " + varName + "->Weight = " + std::to_string(weight) + ";");
        c_.emitLine("    " + varName + "->Charset = " + std::to_string(charset) + ";");
        c_.emitLine("    " + varName + "->Bold = " + std::to_string(bBold) + ";");
        c_.emitLine("    " + varName + "->Italic = " + std::to_string(bItalic) + ";");
        c_.emitLine("    " + varName + "->Underline = " + std::to_string(bUnder) + ";");
        c_.emitLine("    " + varName + "->Strikethrough = " + std::to_string(bStrike) + ";");
    };

    // Fix 182: .frx 设计期属性发射（主控件循环与容器子控件共用）。
    // 原先这段内联在 cgen_form_ctrl_style_apply.inc 的 P24 块里, 只有顶层控件走得到;
    // Frame/PictureBox 的子控件由 cgen_form_frame_menu.inc 的 emitChildControls 另起一条精简路径,
    // 从未获取 Picture/Text/List → VBFlexGridDemo 的 Picture1.Picture(MainForm.frx:0000) 一直留白。
    // 放在片段最外层作用域, 后续片段（同一个函数体）都能捕获。
    auto emitControlFrxProps = [&](const FrmControl& ctrl) -> void {
        if (!frxLoaded) return;
            std::string ctrlHwndVar;
            {
                std::string ctrlNameLower = ctrl.controlName;
                std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
                if (knownControlArrays_.count(ctrlNameLower)) {
                    ctrlHwndVar = "vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", " +
                                  std::to_string(ctrl.index >= 0 ? ctrl.index : 0) + ")";
                } else {
                    ctrlHwndVar = "vb6_hwnd_" + cIdent(ctrl.controlName);
                }
            }

            // P24: Control.Picture from .frx (CommandButton/OptionButton/CheckBox/PictureBox/Image)
            if (ctrl.controlType == FrmControlType::CommandButton ||
                ctrl.controlType == FrmControlType::OptionButton ||
                ctrl.controlType == FrmControlType::CheckBox ||
                ctrl.controlType == FrmControlType::PictureBox ||
                ctrl.controlType == FrmControlType::Image) {
                auto picIt = ctrl.properties.find("Picture");
                if (picIt != ctrl.properties.end() && picIt->second.type == FrmValueType::FrxReference) {
                    auto pic = FrxReader::readPicture(picIt->second.frxOffset, false);
                    if (!pic.data.empty()) {
                        std::string varName = "vb6_frx_pic_" + cIdent(ctrl.controlName);
                        if (ctrl.index >= 0) varName += "_" + std::to_string(ctrl.index);
                        h_.emitLine(bytesToHexArray(pic.data.data(), pic.data.size(), varName));
                        bool isGraphicalBtn = (ctrl.controlType == FrmControlType::CommandButton ||
                                               ctrl.controlType == FrmControlType::OptionButton ||
                                               ctrl.controlType == FrmControlType::CheckBox) &&
                                              ctrl.properties.count("Style") &&
                                              ctrl.properties.at("Style").intValue == 1;
                        if (isGraphicalBtn) {
                            // Graphical button: use BM_SETIMAGE with IMAGE_BITMAP
                            c_.emitLine("{ void* vb6_pic = vb6_LoadPictureFromMemory(" + varName + ", " + varName + "_size);");
                            c_.emitLine("  if (vb6_pic) vb6_GraphicalBtn_SetImage((void*)" + ctrlHwndVar + ", vb6_pic); }");
                        } else {
                            // PictureBox/Image: use vb6_SetControlPicture
                            c_.emitLine("{ void* vb6_pic = vb6_LoadPictureFromMemory(" + varName + ", " + varName + "_size);");
                            c_.emitLine("  if (vb6_pic) vb6_SetControlPicture((void*)" + ctrlHwndVar + ", vb6_pic); }");
                        }
                    }
                }
            }

            // P24: TextBox.Text from .frx (multiline text)
            if (ctrl.controlType == FrmControlType::TextBox) {
                auto txtIt = ctrl.properties.find("Text");
                if (txtIt != ctrl.properties.end() && txtIt->second.type == FrmValueType::FrxReference) {
                    auto txt = FrxReader::readText(txtIt->second.frxOffset);
                    if (!txt.text.empty()) {
                        c_.emitLine("{ wchar_t* vb6_wtxt = vb6_Utf8ToWide(\"" + escapeCString(txt.text) + "\");");
                        c_.emitLine("  if (vb6_wtxt) { SetWindowTextW((HWND)" + ctrlHwndVar + ", vb6_wtxt); free(vb6_wtxt); } }");
                    }
                }
            }

            // P24: ListBox.List + ItemData from .frx
            if (ctrl.controlType == FrmControlType::ListBox) {
                auto listIt = ctrl.properties.find("List");
                bool hasListFromFrx = (listIt != ctrl.properties.end() && listIt->second.type == FrmValueType::FrxReference);
                FrxListData listData;
                if (hasListFromFrx) {
                    listData = FrxReader::readStringList(listIt->second.frxOffset);
                    for (size_t i = 0; i < listData.items.size(); i++) {
                        c_.emitLine("{ wchar_t* vb6_witem = vb6_Utf8ToWide(\"" + escapeCString(listData.items[i]) + "\");");
                        c_.emitLine("  if (vb6_witem) { SendMessageW((HWND)" + ctrlHwndVar + ", LB_ADDSTRING, 0, (LPARAM)vb6_witem); free(vb6_witem); } }");
                    }
                }
                auto idataIt = ctrl.properties.find("ItemData");
                if (idataIt != ctrl.properties.end() && idataIt->second.type == FrmValueType::FrxReference) {
                    auto intData = FrxReader::readIntList(idataIt->second.frxOffset);
                    for (size_t i = 0; i < intData.items.size(); i++) {
                        c_.emitLine("SendMessageW((HWND)" + ctrlHwndVar + ", LB_SETITEMDATA, " +
                                    std::to_string((int)i) + ", (LPARAM)" + std::to_string(intData.items[i]) + ");");
                    }
                }
            }
    };

    // 账 #160: 声明序号 —— 设计期没写 `TabIndex` 时的回落值（见 emitDesignerTabIndexProp）。
    // 自增放在循环体第一句：这条循环里有多处 `continue`，放末尾会让后面的控件集体少一号。
    int tabIndexFallback = -1;
    for (const auto& ctrl : frmDesc.formControl.children) {
        tabIndexFallback++;
        // Fix 112: 工程内 UserControl (.ctl) 子控件 — 实例化到宿主子窗口.
        // 此前这类 controlTypeToWin32Class()==nullptr 的非 ActiveX 类型直接 continue,
        // Form2 上的 ucChartBar1 等从未创建 → vb6_ComCall(NULL,...) 被丢弃 → 白板窗体.
        if (findUserControlSpec(ctrl.controlTypeName)) {
            std::string ctlClass = knownUserControlCtrlVars_[Symbol::toLower(ctrl.controlName)];
            // 该 .ctl 模块的宿主描述必须先注册 (无 TU: extern 声明 + 直接调用)
            c_.emitLine("{ extern void vb6_" + cIdent(ctlClass) + "_RegisterHost(void);"
                      + " vb6_" + cIdent(ctlClass) + "_RegisterHost(); }");
            int ucL = 0, ucT = 0, ucW = 1000, ucH = 300;
            auto ucP = ctrl.properties.find("Left");
            if (ucP != ctrl.properties.end()) ucL = (int)ucP->second.intValue;
            ucP = ctrl.properties.find("Top");
            if (ucP != ctrl.properties.end()) ucT = (int)ucP->second.intValue;
            ucP = ctrl.properties.find("Width");
            if (ucP != ctrl.properties.end()) ucW = (int)ucP->second.intValue;
            ucP = ctrl.properties.find("Height");
            if (ucP != ctrl.properties.end()) ucH = (int)ucP->second.intValue;

            std::string ctlLower = Symbol::toLower(ctrl.controlName);
            // Fix 119: 实例字体 (BeginProperty Font) 必须在 HostCreate 之前设定 —
            // 控件的 InitProperties 会以 UserControl.Font 为默认字体基准。
            for (const auto& blk : ctrl.propertyBlocks) {
                if (_stricmp(blk.blockName.c_str(), "Font") != 0) continue;
                emitFrmFontObj(blk, "_vb6_f119");
                c_.emitLine("    vb6_UC_SetPendingFont(_vb6_f119);");
                c_.emitLine("  } }");
            }
            bool ucIsArr = knownControlArrays_.count(ctlLower) != 0;
            std::string hwndExpr;
            if (ucIsArr) {
                int idx = ctrl.index >= 0 ? ctrl.index : 0;
                c_.emitLine("{ void* vb6_uc_h = vb6_UC_HostCreate(\"" + ctlClass + "\", "
                    + std::to_string(ucL) + ", " + std::to_string(ucT) + ", "
                    + std::to_string(ucW) + ", " + std::to_string(ucH)
                    + ", hwnd, hInstance, \"" + cIdent(ctrl.controlName) + "_" + std::to_string(idx)
                    + "\", " + std::to_string(idx) + ");");
                c_.emitLine("  vb6_CtrlArr_SetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", "
                           + std::to_string(idx) + ", vb6_uc_h); }");
                hwndExpr = "vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", "
                         + std::to_string(idx) + ")";
            } else {
                c_.emitLine("vb6_hwnd_" + cIdent(ctrl.controlName) + " = vb6_UC_HostCreate(\""
                    + ctlClass + "\", " + std::to_string(ucL) + ", " + std::to_string(ucT) + ", "
                    + std::to_string(ucW) + ", " + std::to_string(ucH)
                    + ", hwnd, hInstance, \"" + ctrl.controlName + "\", -1);");
                hwndExpr = "vb6_hwnd_" + cIdent(ctrl.controlName);
            }

            // Fix 112/114: 设计期持久化的 UserControl 公共属性
            // (Caption=... / Value=30 / PB_Color1=... / Caption2="..." / ChartStyle=1 ...)
            // — 逐个调用该 .ctl 的 prop_let_ (跳过几何/Tab 等标准属性).
            // Fix 114: 原来用 symTab_.lookupModule(pn) 判断属性是否为 Property Let,
            // 但 lookupModule 只查**当前模块**(Form2)作用域, 而 Property Let 符号属于
            // 被引用 .ctl 模块 → 恒为 nullptr → 所有设计期属性一个都没发射, 导致
            // 三个 ucProgressCircular 只剩默认 50% 蓝环 / LabelPlus 无标题与底色 /
            // 饼图 ChartStyle、LabelsVisible 等全部丢失。改为跨模块查类写方向形参表。
            {
                static const char* kSkip[] = {"Left", "Top", "Width", "Height", "Index",
                    "TabIndex", "TabStop", "Visible", "Enabled", "Name", "_ExtentX",
                    "_ExtentY", "Tag", "ToolTipText", "WhatisThis", "HelpContextID",
                    // BeginProperty 复合块由 propertyBlocks 单独处理, 不应作为简单属性发射
                    "Font", "IconFont", "TitleFont", "Caption1_Font", "Caption2_Font",
                    "Images", "ListImages", "MouseIcon"};
                c_.emitLine("vb6_UC_Enter(" + hwndExpr + ");  /* Fix 112: 换入宿主状态 */");
                // Fix 119: 字体类复合属性块 (Caption1_Font / Caption2_Font / TitleFont /
                // IconFont / Font ...) — 若 .ctl 有同名属性写方向, 构造字体对象后调用之。
                // 例: 三个 ucProgressCircular 的 Caption1_Font=14.25 此前完全没生效。
                for (const auto& blk : ctrl.propertyBlocks) {
                    // Fix 119/125: 字体类块 → 构造字体对象后调用 .ctl 的同名写属性。
                    // · Property Let (ByVal StdFont)  → vb6_ComIface_Font*  直接传
                    // · Property Set (ByRef StdFont)  → vb6_ComIface_Font** 传 &变量
                    //   这类实现体是 `With m_X_Font: .Name = New_Font.Name ...`, 由
                    //   vb6_ComSetProp 落到"字体结构体直写字段"快路径 (Fix 125), 不再崩。
                    std::vector<ParameterInfo> fp119;
                    bool isSet119 = false;
                    if (findClassMemberWriteParams(ctlClass, blk.blockName, false, fp119)
                        && !fp119.empty()) {
                        isSet119 = false;
                    // Fix 130: Property **Set** 型字体属性 (ucProgressCircular.Caption1_Font /
                    // Caption2_Font, LabelPlus.Font) 走 vb6_UC_FontAssign 安全通道:
                    // 用该控件自己的 Property Get 拿到内部字体对象后**直写字段**,
                    // 不调用其 Set 实现体 (内含 Refresh, 会在 CreateControls 阶段
                    // 数据未就绪时刷新并把图表永久置空 —— Fix 127 实测)。
                    // 若不应用, InitProperties 的 `Set m_X_Font = Ambient.Font` 会把
                    // 宿主共享字体当成员用, 后续 COM 读写把它打坏 (实测 Size 变 1,
                    // 仪表 50%/30%/70% 与 Band/Visitors/Volume 缩成一点)。
                    } else if (findClassMemberWriteParams(ctlClass, blk.blockName, true, fp119)
                               && !fp119.empty()) {
                        isSet119 = true;
                    } else {
                        continue;
                    }
                    // 字体类块命名惯例: Font / TitleFont / Caption1_Font / IconFont ...
                    bool fontName119 = false;
                    {
                        std::string bn = Symbol::toLower(blk.blockName);
                        if (bn == "font" || bn.size() >= 4
                            && bn.compare(bn.size() - 4, 4, "font") == 0)
                            fontName119 = true;
                    }
                    if (!fontName119) continue;
                    std::string canonF = canonicalClassMemberName(ctlClass, blk.blockName);
                    if (canonF.empty()) canonF = blk.blockName;
                    std::string setterF = "vb6_" + cIdent(ctlClass)
                                        + (isSet119 ? "_prop_set_" : "_prop_let_") + cIdent(canonF);
                    emitFrmFontObj(blk, "_vb6_f119");
                    if (isSet119) {
                        // Fix 128: Property Set 型**不调用**其 Set 实现体 (内含 Refresh,
                        // 会在数据未就绪时刷新并把图表永久置空)。改用该控件自己的
                        // Property Get 取到内部字体对象, 原地覆写字段。
                        std::string getterF = "vb6_" + cIdent(ctlClass)
                                            + "_prop_get_" + cIdent(canonF);
                        c_.emitLine("    vb6_UC_FontAssign((void*)" + getterF
                                  + "((vb6_cls_" + cIdent(ctlClass) + "*)vb6_UC_InstanceOf("
                                  + hwndExpr + ")), _vb6_f119);");
                    } else {
                        c_.emitLine("    " + setterF + "((vb6_cls_" + cIdent(ctlClass)
                                  + "*)vb6_UC_InstanceOf(" + hwndExpr + "), _vb6_f119);");
                    }
                    c_.emitLine("  } }");
                }
                for (const auto& kv : ctrl.properties) {
                    const std::string& pn = kv.first;
                    bool skip = false;
                    for (const char* sk : kSkip) if (_stricmp(pn.c_str(), sk) == 0) { skip = true; break; }
                    if (skip) continue;
                    // Fix 114: 跨模块查询 .ctl 类的 Property Let 形参表 (driver 已注入
                    // 外部 Class 符号 → Phase B 可命中 memberLetParams).
                    std::vector<ParameterInfo> wp114;
                    if (!findClassMemberWriteParams(ctlClass, pn, false, wp114) || wp114.empty())
                        continue;
                    // 属性名拼写按 .ctl 内声明规范化 (VB6 大小写不敏感), 否则 LNK2019
                    std::string canonPn = canonicalClassMemberName(ctlClass, pn);
                    if (canonPn.empty()) canonPn = pn;
                    std::string setter = "vb6_" + cIdent(ctlClass) + "_prop_let_" + cIdent(canonPn);
                    const ParameterInfo& p114 = wp114.back();
                    std::string meExpr = "(vb6_cls_" + cIdent(ctlClass)
                                       + "*)vb6_UC_InstanceOf(" + hwndExpr + ")";

                    // 解析字符串值 (含 .frx 属性包持久化引用, 如 Caption="Form2.frx":006E)
                    std::string textVal;
                    bool isStr = false;
                    if (kv.second.type == FrmValueType::String) {
                        isStr = true;
                        textVal = kv.second.rawText;
                    } else if (kv.second.type == FrmValueType::FrxReference) {
                        if (frxLoaded) {
                            FrxTextData ds114 = FrxReader::readDocString(kv.second.frxOffset);
                            if (!ds114.text.empty()) { isStr = true; textVal = ds114.text; }
                        }
                        if (!isStr) continue;   // 无法解析则跳过, 避免发射非法字面量
                    }
                    if (isStr) {
                        if (textVal.size() >= 2 && textVal.front() == '"' && textVal.back() == '"')
                            textVal = textVal.substr(1, textVal.size() - 2);
                        std::string lit = "vb6_BSTR_FromStr(L\"" + escapeWideCString(textVal) + "\")";
                        if (p114.isByVal) {
                            c_.emitLine(setter + "(" + meExpr + ", " + lit + ");");
                        } else {
                            // ByRef New_X As String → C 形参为 BSTR*, 需可寻址临时量
                            c_.emitLine("{ BSTR _lp114 = " + lit + "; "
                                      + setter + "(" + meExpr + ", &_lp114); }");
                        }
                    } else {
                        std::string valExpr = kv.second.rawText.empty()
                                ? std::to_string(kv.second.intValue) : kv.second.rawText;
                        c_.emitLine(setter + "(" + meExpr + ", " + valExpr + ");");
                    }
                }

                // czUI fix: UC 事件接线 — thunk/sink 已在文件作用域 (prelude) 生成,
                // 这里按实例挂接 (attach 行由 prelude 存入 ucEventAttach_)。
                {
                    auto itAtt114 = ucEventAttach_.find(Symbol::toLower(ctrl.controlName));
                    if (itAtt114 != ucEventAttach_.end() && !getenv("C3_NO_EVENTS")) {
                        std::string shortCls114 = ctlClass;
                        size_t dotCls114 = shortCls114.rfind('.');
                        if (dotCls114 != std::string::npos) shortCls114 = shortCls114.substr(dotCls114 + 1);
                        c_.emitLine("{ /* UC events */");
                        for (const auto& al : itAtt114->second) c_.emitLine("  " + al);
                        c_.emitLine("  ((vb6_cls_" + cIdent(shortCls114) + "*)vb6_UC_InstanceOf(" + hwndExpr + "))->events = &_sink_" + cIdent(ctrl.controlName) + ";");
                        c_.emitLine("}");
                    }
                }

                // czUI fix: 设计期属性除 prop_let 直赋外, 再经轻量 PropertyBag 重放一次
                // UserControl_ReadProperties — 复现 .ctl 内部"读取后同步"逻辑
                // (czUI.ctl: `If m_Checked Then m_AnimPos = 1!` → toggle 初始 ON;
                //  这类语句只在 ReadProperties 里, prop_let 直赋路径会绕过它)。
                {
                    // ctlClass 可能是 "czFormDemo.czControl" 带命名空间 —
                    // 用设计器描述 + 代码段文本探测 .ctl 是否定义了
                    // UserControl_ReadProperties (CCodeGen 每模块重建, 不能用
                    // 跨模块 state; 类符号只存可调用成员, Private Sub 不在其中)
                    const FrmFile* ucSpec114 = findUserControlSpec(ctlClass);
                    bool hasReadProps114 = false;
                    if (ucSpec114) {
                        std::string lowerCode = Symbol::toLower(ucSpec114->codeSection);
                        hasReadProps114 =
                            lowerCode.find("sub usercontrol_readproperties") != std::string::npos;
                    }
                    static const char* noBagCls = getenv("C3_NO_BAG_CLS");
                    static const char* onlyBagCls = getenv("C3_ONLY_BAG_CLS");
                    bool bagBlocked = (noBagCls && strstr(noBagCls, Symbol::toLower(ctlClass).c_str()))
                        || (onlyBagCls && strstr(onlyBagCls, Symbol::toLower(ctlClass).c_str()) == nullptr);
                    if (hasReadProps114 && !getenv("C3_NO_BAG") && !bagBlocked) {
                        c_.emitLine("extern void vb6_" + cIdent(ctlClass) + "_UC_ReadProps(void* me, void** PropBag);");
                        c_.emitLine("{ void* _bag = vb6_UC_PropBagCreate();");
                        for (const auto& kvb : ctrl.properties) {
                            const std::string& pn114 = kvb.first;
                            bool skip114 = false;
                            for (const char* sk : kSkip) if (_stricmp(pn114.c_str(), sk) == 0) { skip114 = true; break; }
                            if (skip114) continue;
                            std::string litName = "L\"" + escapeWideCString(pn114) + "\"";
                            if (kvb.second.type == FrmValueType::Integer) {
                                c_.emitLine("vb6_UC_BagPutInt(_bag, " + litName + ", " + std::to_string((int)kvb.second.intValue) + ");");
                            } else if (kvb.second.type == FrmValueType::Float) {
                                c_.emitLine("vb6_UC_BagPutDbl(_bag, " + litName + ", " + std::to_string(kvb.second.floatValue) + ");");
                            } else if (kvb.second.type == FrmValueType::String) {
                                std::string tv = kvb.second.rawText;
                                if (tv.size() >= 2 && tv.front() == '"' && tv.back() == '"')
                                    tv = tv.substr(1, tv.size() - 2);
                                c_.emitLine("vb6_UC_BagPutStr(_bag, " + litName + ", L\"" + escapeWideCString(tv) + "\");");
                            } else if (kvb.second.type == FrmValueType::Identifier) {
                                std::string idl = Symbol::toLower(kvb.second.rawText);
                                if (idl == "true")  c_.emitLine("vb6_UC_BagPutBool(_bag, " + litName + ", -1);");
                                else if (idl == "false") c_.emitLine("vb6_UC_BagPutBool(_bag, " + litName + ", 0);");
                            } else if (kvb.second.type == FrmValueType::FrxReference && frxLoaded) {
                                // czUI fix: Caption 等字符串常存于 .frx (如 "Form2.frx":006E) —
                                // 跳过会让 ReadProperty 返回默认空串, 把 prop_let 设置好的
                                // 标题清空 (Charts2020 所有 LabelPlus 标题消失的根因)
                                FrxTextData dsB114 = FrxReader::readDocString(kvb.second.frxOffset);
                                if (!dsB114.text.empty()) {
                                    std::string tvB = dsB114.text;
                                    if (tvB.size() >= 2 && tvB.front() == '"' && tvB.back() == '"')
                                        tvB = tvB.substr(1, tvB.size() - 2);
                                    c_.emitLine("vb6_UC_BagPutStr(_bag, " + litName + ", L\"" + escapeWideCString(tvB) + "\");");
                                }
                            }
                        }
                        c_.emitLine("vb6_" + cIdent(ctlClass) + "_UC_ReadProps((vb6_cls_" + cIdent(ctlClass)
                                  + "*)vb6_UC_InstanceOf(" + hwndExpr + "), &_bag);");
                        c_.emitLine("vb6_UC_PropBagFree(_bag); }");
                    }
                }
            }
            ctrlId++;
            continue;
        }
        const char* win32Class = FrmParser::controlTypeToWin32Class(ctrl.controlType);
        if (!win32Class) {
            // P7.9: WebBrowser控件用vb6_CreateWebView创建
            if (ctrl.controlType == FrmControlType::WebBrowser) {
                int wbLeft = 0, wbTop = 0, wbWidth = 5000, wbHeight = 3000;
                auto pIt = ctrl.properties.find("Left");
                if (pIt != ctrl.properties.end()) wbLeft = (int)pIt->second.intValue;
                pIt = ctrl.properties.find("Top");
                if (pIt != ctrl.properties.end()) wbTop = (int)pIt->second.intValue;
                pIt = ctrl.properties.find("Width");
                if (pIt != ctrl.properties.end()) wbWidth = (int)pIt->second.intValue;
                pIt = ctrl.properties.find("Height");
                if (pIt != ctrl.properties.end()) wbHeight = (int)pIt->second.intValue;
                // 缇转像素: 1缇=1/15像素 (96DPI)
                c_.emitLine("{ void* vb6_tmp_hwnd = vb6_CreateWebView((void*)hwnd, " +
                    std::to_string(wbLeft / 15) + ", " + std::to_string(wbTop / 15) + ", " +
                    std::to_string(wbWidth / 15) + ", " + std::to_string(wbHeight / 15) + ", L\"" +
                    cIdent(ctrl.controlName) + "\");");
                c_.emitLine("vb6_hwnd_" + cIdent(ctrl.controlName) + " = vb6_tmp_hwnd; }");
            }
            // P20-39: ImageList 走**原生复刻** (comctl32 ImageList_* API), 不 CoCreateInstance。
            // 三个理由: ①mscomctl.ocx 未注册 (缺 InprocServer32); ②32 位 inproc OCX 进不了
            // x64 进程; ③ImageList 无窗口, 走 COM 还得养一个假 IDispatch, 而它的全部价值
            // 就是给其它控件递 HIMAGELIST —— 原生句柄一步到位。
            if (ctrl.controlType == FrmControlType::ImageList) {
                std::string ilSlot = "vb6_com_" + cIdent(ctrl.controlName);
                // 登记裸名 (不是 ilSlot): cgen_util_com.cpp 里的 imageListSlotOfExpr 是从
                // `vb6_ComGetObjectProp(vb6_com_X, L"ListImages")` 里抠出 X 来比对的。
                imageListVars_.insert(ctrl.controlName);
                int ilW = 0, ilH = 0;
                auto ilWIt = ctrl.properties.find("ImageWidth");
                if (ilWIt != ctrl.properties.end() && ilWIt->second.intValue > 0)
                    ilW = (int)ilWIt->second.intValue;
                auto ilHIt = ctrl.properties.find("ImageHeight");
                if (ilHIt != ctrl.properties.end() && ilHIt->second.intValue > 0)
                    ilH = (int)ilHIt->second.intValue;
                // 传槽的**地址** (vb6_ImageList_Create 会 malloc 并写回, 槽里放的是
                // 复刻实例指针, 不是 HWND)。
                c_.emitLine("vb6_ImageList_Create((void**)&" + ilSlot + ", "
                          + std::to_string(ilW) + ", " + std::to_string(ilH) + ");");

                // 设计期图片: .frm 的 Images 块里每个 ListImage 带
                // `Picture = "<控件名>.frx":<偏移>`, 字节已由 FrxReader 读出并烤成 hex 数组。
                // 块名两种都得认: 老解析器写的是 `BeginProperty Images {GUID}`, 而 P20-39
                // 的集合块解析器按 VB6 原文取集合名 → `ListImages`。只认 "Images" 会让
                // 设计期图片被整个跳过 (表现为 Count 从 2 掉成 0, 且不报任何错)。
                for (const auto& block : ctrl.propertyBlocks) {
                    if (block.blockName != "Images" && block.blockName != "ListImages") continue;
                    int imgIdx = 0;
                    for (const auto& listImg : block.nestedBlocks) {
                        auto picIt = listImg.properties.find("Picture");
                        if (picIt == listImg.properties.end()
                            || picIt->second.type != FrmValueType::FrxReference) continue;
                        FrxPictureData pic = FrxReader::readPicture(picIt->second.frxOffset, true);
                        if (pic.data.empty()) continue;
                        imgIdx++;
                        std::string varName = "vb6_frx_imglist_" + cIdent(ctrl.controlName)
                                            + "_" + std::to_string(imgIdx);
                        h_.emitLine(bytesToHexArray(pic.data.data(), pic.data.size(), varName));
                        std::string keyLit = "NULL";
                        auto keyIt = listImg.properties.find("Key");
                        if (keyIt != listImg.properties.end()
                            && keyIt->second.type == FrmValueType::String) {
                            std::string raw = keyIt->second.rawText;
                            if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
                                std::string key = raw.substr(1, raw.size() - 2);
                                for (char& ch : key) if (ch == '"') ch = '\'';
                                keyLit = "L\"" + key + "\"";
                            }
                        }
                        c_.emitLine("vb6_ImageList_AddDesignTimeImage((void*)" + ilSlot
                                  + ", " + keyLit + ", " + varName + ", "
                                  + varName + "_size);");
                    }
                    break;
                }

                ctrlId++;
                continue;
            }
            // ActiveX控件: CoCreateInstance
            // 以下 **Toolbar / CommonDialog** 仍是 OCX 占位, 等 P20-40+ 逐个复刻。
            // P20-40: **StatusBar 不在此列** —— 它是 msctls_status32 原生复刻, 必须落到下面
            // 的常规创建路径 (vb6_CreateControl) 才有窗口; 走这里的 CoCreateInstance 会
            // 直接 `continue` 掉, 窗口根本不创建, 属性全落 vb6_ComGetStringProp 静默答错。
            // P24-05: 优先用TypeLib中的CLSID直接生成代码, 避免ProgID不存在或错误
            // VB6 .frm的Begin行用CoClass名(如MSComctlLib.ImageList), 但真实ProgID可能不同(如MSComctlLib.ImageListCtrl)
            // D6 / C29-9: CommonDialog **已从这一组摘出** —— MSComDlg.OCX 只有 32 位，
            // x64 里 CoCreateInstance 失败 ⇒ 整枚控件静默空转。它现在走上面那条普通
            // 控件创建路（自注册的不可见类 VB6_COMMONDIALOG 当属性袋），六个 Show*
            // 直接调 comdlg32。剩下这三类等各自批次到位后同样摘出去。
            // 合并(P20-40 + C29-9): StatusBar 与 CommonDialog **都从这一组摘出**了
            // (前者 msctls_status32 原生复刻、后者 VB6_COMMONDIALOG 属性宿主), 它们
            // 落到下面的常规创建路 vb6_CreateControl 才有窗口。这里只剩无窗口的两类。
            // C29-5a: **Toolbar 已从这一组摘出** —— 它现在走下面那条常规创建路
            // (vb6_CreateControl) 才有真窗口；留在这里会被 continue 掉、句柄压根不声明。
            if (ctrl.controlType == FrmControlType::ImageList) {
                // 从符号表查找ComClass, 获取真实的CLSID或ProgID
                std::string comVarName = "vb6_com_" + cIdent(ctrl.controlName);
                Symbol* comSym = lookupDotted(ctrl.controlTypeName);
                if (!comSym || comSym->kind != SymbolKind::ComClass) {
                    // Try short name (e.g. "ImageList")
                    size_t dot = ctrl.controlTypeName.find('.');
                    if (dot != std::string::npos) {
                        comSym = symTab_.lookup(ctrl.controlTypeName.substr(dot + 1));
                    }
                }
                if (comSym && comSym->kind == SymbolKind::ComClass && !comSym->comClsidStr.empty()) {
                    // Use CLSID directly (most reliable, no runtime ProgID lookup)
                    // Convert "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}" to C init form:
                    // {0xXXXXXXXX,0xXXXX,0xXXXX,{0xXX,0xXX,0xXX,0xXX,0xXX,0xXX,0xXX,0xXX}}
                    std::string raw = comSym->comClsidStr;  // e.g. "{2C247F23-8591-11D1-B16A-00C0F0283628}"
                    std::string cInit;
                    if (raw.size() >= 38 && raw.front() == '{' && raw.back() == '}') {
                        std::string h = raw.substr(1, raw.size()-2); // strip braces
                        // Split by '-'
                        std::vector<std::string> parts;
                        size_t pos = 0;
                        while (pos < h.size()) {
                            size_t dash = h.find('-', pos);
                            if (dash == std::string::npos) { parts.push_back(h.substr(pos)); break; }
                            parts.push_back(h.substr(pos, dash-pos));
                            pos = dash + 1;
                        }
                        if (parts.size() == 5) {
                            cInit = "0x" + parts[0] + ",0x" + parts[1] + ",0x" + parts[2] + ",{0x"
                                + parts[3].substr(0,2) + ",0x" + parts[3].substr(2,2) + ",0x"
                                + parts[4].substr(0,2) + ",0x" + parts[4].substr(2,2) + ",0x"
                                + parts[4].substr(4,2) + ",0x" + parts[4].substr(6,2) + ",0x"
                                + parts[4].substr(8,2) + ",0x" + parts[4].substr(10,2) + "}";
                        }
                    }
                    if (cInit.empty()) cInit = "0,0,0,{0,0,0,0,0,0,0,0}";  // fallback
                    c_.emitLine("{ static const CLSID vb6_clsid_" + cIdent(ctrl.controlName) + " = {" + cInit + "};");
                    c_.emitLine("  CoCreateInstance(&vb6_clsid_" + cIdent(ctrl.controlName) + ", NULL, 1/*CLSCTX_INPROC_SERVER*/, &IID_IDispatch, (void**)&" + comVarName + "); }");
                } else {
                    // Fallback: use ProgID from TypeLib if available, else controlTypeName
                    std::string progId = (comSym && !comSym->comProgId.empty()) ? comSym->comProgId : ctrl.controlTypeName;
                    c_.emitLine("{ CLSID vb6_clsid; CLSIDFromProgID(L\"" + progId + "\", &vb6_clsid);");
                    c_.emitLine("  CoCreateInstance(&vb6_clsid, NULL, 1/*CLSCTX_INPROC_SERVER*/, &IID_IDispatch, (void**)&" + comVarName + "); }");
                }

                // ImageList: set ImageWidth/ImageHeight and load ListImages from .frx
                if (ctrl.controlType == FrmControlType::ImageList && frxLoaded) {
                    std::string comVar = "vb6_com_" + cIdent(ctrl.controlName);
                    // Set ImageWidth/ImageHeight via COM (must be set before adding images)
                    auto iwIt = ctrl.properties.find("ImageWidth");
                    if (iwIt != ctrl.properties.end() && iwIt->second.intValue > 0) {
                        c_.emitLine("vb6_ComSetProp((void*)" + comVar + ", L\"ImageWidth\", vb6_ComPackInt(" + std::to_string((int)iwIt->second.intValue) + "));");
                    }
                    auto ihIt = ctrl.properties.find("ImageHeight");
                    if (ihIt != ctrl.properties.end() && ihIt->second.intValue > 0) {
                        c_.emitLine("vb6_ComSetProp((void*)" + comVar + ", L\"ImageHeight\", vb6_ComPackInt(" + std::to_string((int)ihIt->second.intValue) + "));");
                    }

                    // Find "Images" property block with nested ListImage entries
                    for (const auto& block : ctrl.propertyBlocks) {
                        if (block.blockName == "Images") {
                            int imgIdx = 0;
                            for (const auto& listImg : block.nestedBlocks) {
                                // Each nested block is like: blockName="ListImage1", has Picture/Key/Tag
                                auto picIt = listImg.properties.find("Picture");
                                if (picIt == listImg.properties.end() || picIt->second.type != FrmValueType::FrxReference) continue;

                                // Read image data from .frx (ImageList format: isImageList=true)
                                auto pic = FrxReader::readPicture(picIt->second.frxOffset, true);
                                if (pic.data.empty()) continue;

                                // Generate byte array in .h file
                                imgIdx++;
                                std::string varName = "vb6_frx_imglist_" + cIdent(ctrl.controlName) + "_" + std::to_string(imgIdx);
                                h_.emitLine(bytesToHexArray(pic.data.data(), pic.data.size(), varName));

                                // Get Key property (may be empty string)
                                std::string keyStr;
                                auto keyIt = listImg.properties.find("Key");
                                if (keyIt != listImg.properties.end() && keyIt->second.type == FrmValueType::String) {
                                    std::string raw = keyIt->second.rawText;
                                    if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') keyStr = raw.substr(1, raw.size() - 2);
                                }

                                // ListImages.Add(index, key, picture) via COM
                                // Get ListImages collection, then call Add method
                                c_.emitLine("{ void* vb6_lstImgs = vb6_ComGetObjectProp((void*)" + comVar + ", L\"ListImages\");");
                                c_.emitLine("  if (vb6_lstImgs) {");
                                c_.emitLine("    void* vb6_picCom = vb6_LoadPictureAsCom(" + varName + ", " + varName + "_size);");
                                c_.emitLine("    if (vb6_picCom) {");
                                // args: [index, key, picture] -- VB6 positional parameter order
                                c_.emitLine("      void* vb6_args[] = { vb6_ComPackInt(" + std::to_string(imgIdx) + "), vb6_ComPackBSTR(L\"" + keyStr + "\"), vb6_ComPackObject(vb6_picCom) };");
                                c_.emitLine("      vb6_ComVarFree((void*)vb6_ComCall(vb6_lstImgs, L\"Add\", vb6_args, 3));");
                                c_.emitLine("      vb6_ReleaseObject(&vb6_picCom);");
                                c_.emitLine("    }");
                                c_.emitLine("    vb6_ReleaseObject(&vb6_lstImgs);");
                                c_.emitLine("  } }");
                            }
                            break;  // Only one "Images" block expected
                        }
                    }
                }

                ctrlId++;
                continue;
            }
            // Fix 143: 其他第三方 OCX 控件 (NewTabCtl.NewTab 等) → vb6_OcxHost_Create 真宿主.
            // 之前 Unknown 控件什么都不生成: vb6_hwnd_<name> 恒 NULL, 全部 COM 属性/方法
            // 调用静默无效 → 控件不渲染、集合为空 (NewTab1.Themes 枚举不出).
            // 现路径: CLSID 从 typelib ComClass 符号取; ocx 文件从 vbp Object= 表取;
            // 设计期属性走 IPersistPropertyBag. IDispatch* 存入 vb6_hwnd_<name>,
            // 使 vb6_ComCall/ComSetProp(vb6_hwnd_<name>, ...) 走真实 IDispatch.
            {
                Symbol* comSym143 = lookupDotted(ctrl.controlTypeName);
                if (!comSym143 || comSym143->kind != SymbolKind::ComClass) {
                    size_t dot143 = ctrl.controlTypeName.find('.');
                    if (dot143 != std::string::npos)
                        comSym143 = symTab_.lookup(ctrl.controlTypeName.substr(dot143 + 1));
                }
                if (comSym143 && comSym143->kind == SymbolKind::ComClass
                    && !comSym143->comClsidStr.empty()) {
                    // Fix 143b: 控件实例化 CLSID 优先用 typelib 的 coclass GUID (它来自
                    // OCX 文件本身, 是真实类); vbp Object= 的 GUID 作备用 — 实测工程里
                    // 两者不一致 (vbp 写的是旧 GUID, OCX 的 DllGetClassObject 不认它,
                    // 而 typelib GUID 已注册可用). 两个都交给 RTL 依次尝试.
                    auto normGuid = [](const std::string& g) {
                        std::string r;
                        for (char ch : g) {
                            if (ch == '{' || ch == '}') continue;
                            r += (ch >= 'A' && ch <= 'Z') ? (char)(ch - 'A' + 'a') : ch;
                        }
                        return r;
                    };
                    std::string tlKey143 = normGuid(comSym143->comClsidStr);
                    std::string clsid143 = comSym143->comClsidStr;
                    std::string altClsid143;
                    std::string ocxPath143;
                    if (ocxRefs_ && !ocxRefs_->empty()) {
                        for (const auto& ref143 : *ocxRefs_) {
                            if (ref143.first == tlKey143) {
                                // typelib GUID 命中: 只用它, 路径取该条
                                ocxPath143 = ref143.second;
                                break;
                            }
                        }
                        if (ocxPath143.empty()) {
                            // 不命中 (vbp 写的 GUID 与 typelib coclass GUID 不一致, 见
                            // Fix 143b): vbp 的 GUID 作备用 CLSID, 路径取第一条。
                            //
                            // Fix 160: 备用 CLSID 与路径【都】只在单 OCX 工程回退。
                            // 旧实现只给路径加了 size()==1 守卫, altClsid143 却对任意
                            // 条目数都取 front() —— 多 OCX 工程下正在生成 B 控件时拿到
                            // A.ocx 的 CLSID, 而此时 ocxPath143 为空 (下方 ocxFiles_
                            // 按 tlKey143 也查不到), ocxCreateAny 会退到 CoCreateInstance
                            // 用 A 的 CLSID 激活出 A 的控件, 代码生成侧却按 B 处理,
                            // 静默用错控件。单 OCX 工程 front() 就是自己, 回退语义不变。
                            // Fix 148: front() 前必须判空 — ocxRefs_ 为空时是 UB, 实测段错误.
                            if (ocxRefs_->size() == 1) {
                                altClsid143 = "{" + ocxRefs_->front().first + "}";
                                ocxPath143 = ocxRefs_->front().second;
                            }
                        }
                    }
                    if (ocxPath143.empty() && ocxFiles_) {
                        auto itOcx143 = ocxFiles_->find(tlKey143);
                        if (itOcx143 != ocxFiles_->end()) ocxPath143 = itOcx143->second;
                    }
                    std::string ocxArg143 = ocxPath143.empty()
                        ? std::string("NULL")
                        : "L\"" + escapeWideCString(ocxPath143) + "\"";
                    std::string altArg143 = altClsid143.empty()
                        ? std::string("NULL")
                        : "L\"" + escapeWideCString(altClsid143) + "\"";

                    // 几何 (缇→像素)
                    int l143 = 0, t143 = 0, w143 = 3000, h143 = 1500;
                    auto pIt143 = ctrl.properties.find("Left");
                    if (pIt143 != ctrl.properties.end()) l143 = (int)pIt143->second.intValue;
                    pIt143 = ctrl.properties.find("Top");
                    if (pIt143 != ctrl.properties.end()) t143 = (int)pIt143->second.intValue;
                    pIt143 = ctrl.properties.find("Width");
                    if (pIt143 != ctrl.properties.end()) w143 = (int)pIt143->second.intValue;
                    pIt143 = ctrl.properties.find("Height");
                    if (pIt143 != ctrl.properties.end()) h143 = (int)pIt143->second.intValue;

                    // 设计期属性包: 简单属性 (跳过几何/_Extent/Index 与 .frx 引用)
                    static const char* kBagSkip[] = {"Left", "Top", "Width", "Height", "Index",
                        "_ExtentX", "_ExtentY", "Name", "Tag", "TabIndex", "TabStop",
                        "Visible", "Enabled", "WhatsThisHelp", "HelpContextID"};
                    std::vector<std::string> bagLines143;
                    // Fix 148 诊断: C3_OCX_PROPS_MAX=N 只发前 N 个属性 (二分定位崩因)
                    int propsMax143 = 100000;
                    if (const char* pm = std::getenv("C3_OCX_PROPS_MAX")) { int v = atoi(pm); if (v > 0) propsMax143 = v; }
                    for (const auto& kv143 : ctrl.properties) {
                        if ((int)bagLines143.size() >= propsMax143) break;
                        const std::string& pn143 = kv143.first;
                        bool skip143 = false;
                        for (const char* sk : kBagSkip)
                            if (_stricmp(pn143.c_str(), sk) == 0) { skip143 = true; break; }
                        if (skip143) continue;
                        const FrmValue& fv143 = kv143.second;
                        if (fv143.type == FrmValueType::Integer) {
                            bagLines143.push_back("{L\"" + escapeWideCString(pn143) + "\", 3/*VT_I4*/, "
                                + std::to_string(fv143.intValue) + ", 0, NULL, NULL, 0.0f, 0, 0, 0}");
                        } else if (fv143.type == FrmValueType::Float) {
                            bagLines143.push_back("{L\"" + escapeWideCString(pn143) + "\", 4/*VT_R4*/, 0, "
                                + std::to_string(fv143.floatValue) + "f, NULL, NULL, 0.0f, 0, 0, 0}");
                        } else if (fv143.type == FrmValueType::String) {
                            std::string sv143 = fv143.rawText;
                            if (sv143.size() >= 2 && sv143.front() == '"' && sv143.back() == '"')
                                sv143 = sv143.substr(1, sv143.size() - 2);
                            bagLines143.push_back("{L\"" + escapeWideCString(pn143) + "\", 8/*VT_BSTR*/, 0, 0, L\""
                                + escapeWideCString(sv143) + "\", NULL, 0.0f, 0, 0, 0}");
                        }
                        /* Identifier / FrxReference: 跳过 (枚举常量/.frx 二进制, 后续支持) */
                    }
                    /* Fix 149: BeginProperty 字体块 (Font / IconFont(n)) → 设计期 IFont.
                     * stdole 的 IFontDisp 类型库 GUID. VB6 控件的 ReadProperties 会
                     * 直接解引用字体对象 (ctlNewTab.ctl:
                     * `If mTabData(c).IconFontName <> mTabData(c).IconFont.Name`),
                     * 只给默认 Nothing 会空指针解引用 → 异常逃逸 (0xC000041D). */
                    {
                        static const char* kFontGuid143 = "{0BE35203-8F91-11CE-9DE3-00AA004BB851}";
                        for (const auto& blk143 : ctrl.propertyBlocks) {
                            bool isFont143 = false;
                            if (!blk143.blockGuid.empty()) {
                                std::string g143 = blk143.blockGuid;
                                for (auto& ch : g143) ch = (char)toupper((unsigned char)ch);
                                std::string want143 = kFontGuid143;
                                for (auto& ch : want143) ch = (char)toupper((unsigned char)ch);
                                if (g143 == want143) isFont143 = true;
                            }
                            if (!isFont143 && _stricmp(blk143.blockName.c_str(), "Font") == 0) isFont143 = true;
                            if (!isFont143) continue;

                            std::string fName143 = "Tahoma";
                            double fSize143 = 9.0;
                            long fWeight143 = 400, fCharset143 = 0, fFlags143 = 0;
                            auto bit143 = blk143.properties.find("Name");
                            if (bit143 != blk143.properties.end() && bit143->second.type == FrmValueType::String) {
                                std::string s143 = bit143->second.rawText;
                                if (s143.size() >= 2 && s143.front() == '"' && s143.back() == '"')
                                    s143 = s143.substr(1, s143.size() - 2);
                                if (!s143.empty()) fName143 = s143;
                            }
                            bit143 = blk143.properties.find("Size");
                            if (bit143 != blk143.properties.end()) {
                                if (bit143->second.type == FrmValueType::Float) fSize143 = bit143->second.floatValue;
                                else fSize143 = (double)bit143->second.intValue;
                                if (fSize143 <= 0) fSize143 = 9.0;
                            }
                            bit143 = blk143.properties.find("Weight");
                            if (bit143 != blk143.properties.end() && bit143->second.intValue > 0)
                                fWeight143 = (long)bit143->second.intValue;
                            bit143 = blk143.properties.find("Charset");
                            if (bit143 != blk143.properties.end()) fCharset143 = (long)bit143->second.intValue;
                            bit143 = blk143.properties.find("Italic");
                            if (bit143 != blk143.properties.end() && bit143->second.intValue) fFlags143 |= 1;
                            bit143 = blk143.properties.find("Underline");
                            if (bit143 != blk143.properties.end() && bit143->second.intValue) fFlags143 |= 2;
                            bit143 = blk143.properties.find("Strikethrough");
                            if (bit143 != blk143.properties.end() && bit143->second.intValue) fFlags143 |= 4;

                            bagLines143.push_back("{L\"" + escapeWideCString(blk143.blockName)
                                + "\", 9/*VT_DISPATCH(IFont)*/, 0, 0, NULL, L\""
                                + escapeWideCString(fName143) + "\", "
                                + std::to_string(fSize143) + "f, "
                                + std::to_string(fWeight143) + ", "
                                + std::to_string(fCharset143) + ", "
                                + std::to_string(fFlags143) + "}");
                        }
                    }
                    std::string bagArr143 = "NULL";
                    std::string bagCnt143 = "0";
                    if (!bagLines143.empty()) {
                        // 控件数组: 每实例属性不同 → 数组名带索引后缀防重定义
                        bool ocxIsArr143 = knownControlArrays_.count(Symbol::toLower(ctrl.controlName)) != 0;
                        std::string arrName143 = "vb6_ocxprops_" + cIdent(ctrl.controlName);
                        if (ocxIsArr143)
                            arrName143 += "_" + std::to_string(ctrl.index >= 0 ? ctrl.index : 0);
                        c_.emitLine("static const Vb6OcxProp " + arrName143 + "[] = {");
                        c_.indent();
                        for (auto& bl : bagLines143) c_.emitLine(bl + ",");
                        c_.dedent();
                        c_.emitLine("};");
                        bagArr143 = arrName143;
                        bagCnt143 = std::to_string(bagLines143.size());
                    }
                    std::string call143 = "vb6_OcxHost_Create((void*)hwnd, L\"" + clsid143
                              + "\", " + altArg143 + ", " + ocxArg143 + ", "
                              + std::to_string(l143 / 15) + ", " + std::to_string(t143 / 15) + ", "
                              + std::to_string(w143 / 15) + ", " + std::to_string(h143 / 15)
                              + ", L\"" + escapeWideCString(ctrl.controlName) + "\", "
                              + bagArr143 + ", " + bagCnt143 + ")";
                    if (knownControlArrays_.count(Symbol::toLower(ctrl.controlName))) {
                        int idx143 = ctrl.index >= 0 ? ctrl.index : 0;
                        c_.emitLine("{ void* vb6_ocx_h = " + call143 + ";");
                        c_.emitLine("  vb6_CtrlArr_SetAt(&vb6_arr_" + cIdent(ctrl.controlName) + ", "
                                  + std::to_string(idx143) + ", vb6_ocx_h); }  /* Fix 143: OCX 真宿主 (数组) */");
                    } else {
                        c_.emitLine("vb6_hwnd_" + cIdent(ctrl.controlName) + " = " + call143
                                  + ";  /* Fix 143: OCX 真宿主 */");
                    }
                    ctrlId++;
                    continue;
                }
            }
            // 不可见控件 (Timer等) 跳过
            ctrlId++;
            continue;
        }

        // 提取控件属性
        int left = 0, top = 0, width = 1000, height = 300;
        // Fix 162c-extlist: Caption 缺省 = "" 而不是控件名 —— VB6 的 .frm 只在
        // Caption 非默认值时才写盘, Frame/Label/按钮 的持久化默认值都是空串
        // (控件名只是 IDE 插入时的初始值)。此前回落控件名, extlist 的 Frame2
        // 没写 Caption, 运行时标题栏凭空多出 "Frame2" 字样, 与 VB6 参考图不符。
        std::string ctrlCaption;  // 默认 Caption="" (Fix 162c)
        std::string ctrlText;  // TextBox的Text属性

        auto propIt = ctrl.properties.find("Left");
        if (propIt != ctrl.properties.end()) left = (int)propIt->second.intValue;
        propIt = ctrl.properties.find("Top");
        if (propIt != ctrl.properties.end()) top = (int)propIt->second.intValue;
        propIt = ctrl.properties.find("Width");
        if (propIt != ctrl.properties.end()) width = (int)propIt->second.intValue;
        propIt = ctrl.properties.find("Height");
        if (propIt != ctrl.properties.end()) height = (int)propIt->second.intValue;

        propIt = ctrl.properties.find("Caption");
        if (propIt != ctrl.properties.end() && propIt->second.type == FrmValueType::String) {
            std::string raw = propIt->second.rawText;
            if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
                ctrlCaption = raw.substr(1, raw.size() - 2);
            }
        }
        propIt = ctrl.properties.find("Text");
        if (propIt != ctrl.properties.end() && propIt->second.type == FrmValueType::String) {
            std::string raw = propIt->second.rawText;
            if (raw.size() >= 2 && raw.front() == '"' && raw.back() == '"') {
                ctrlText = raw.substr(1, raw.size() - 2);
            }
        }

