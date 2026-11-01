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

    // Fix <vbeclipse> 2026-10-12 (账 #257 第十八层): 容器(CoolBar 等工程内 UserControl)内嵌套的
    // UserControl —— 此前 emitChildControls 仅认 controlTypeToWin32Class()!=NULL 的原生控件,
    // .ctl 类 childClass==NULL 在下面 `if (!childClass) continue` 被跳过 ⇒ 从不创建 ⇒
    // vb6_hwnd_<X> 恒 NULL ⇒ Form_Load 里 `<X>.<属性>` 解引用 NULL 崩 (av read target=0x64)。
    // 这里补齐与顶层 vb6_UC_HostCreate 完全一致的发射路径, 父窗口换成容器 hwnd。
    auto emitUserControlChild = [&](const FrmControl& uc, const std::string& parentHwnd) -> void {
        std::string ctlClass = knownUserControlCtrlVars_[Symbol::toLower(uc.controlName)];
        c_.emitLine("{ extern void vb6_" + cIdent(ctlClass) + "_RegisterHost(void);"
                  + " vb6_" + cIdent(ctlClass) + "_RegisterHost(); }");
        int ucL = 0, ucT = 0, ucW = 1000, ucH = 300;
        auto ucP = uc.properties.find("Left");
        if (ucP != uc.properties.end()) ucL = (int)ucP->second.intValue;
        ucP = uc.properties.find("Top");
        if (ucP != uc.properties.end()) ucT = (int)ucP->second.intValue;
        ucP = uc.properties.find("Width");
        if (ucP != uc.properties.end()) ucW = (int)ucP->second.intValue;
        ucP = uc.properties.find("Height");
        if (ucP != uc.properties.end()) ucH = (int)ucP->second.intValue;

        std::string ctlLower = Symbol::toLower(uc.controlName);
        auto emitUcFrmFontObj = [&](const FrmPropertyBlock& blk, const std::string& varName) {
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
            if (weight >= 700) bBold = -1;
            c_.emitLine("{ vb6_ComIface_Font* " + varName
                      + " = (vb6_ComIface_Font*)vb6_UC_NewFont();");
            c_.emitLine("  if (" + varName + ") {");
            c_.emitLine("    " + varName + "->Name = vb6_BSTR_FromStr(L\"" + escapeWideCString(nm) + "\");");
            c_.emitLine("    " + varName + "->Size = " + floatSingleLiteral(sz, 4) + ";");
            c_.emitLine("    " + varName + "->Weight = " + std::to_string(weight) + ";");
            c_.emitLine("    " + varName + "->Charset = " + std::to_string(charset) + ";");
            c_.emitLine("    " + varName + "->Bold = " + std::to_string(bBold) + ";");
            c_.emitLine("    " + varName + "->Italic = " + std::to_string(bItalic) + ";");
            c_.emitLine("    " + varName + "->Underline = " + std::to_string(bUnder) + ";");
            c_.emitLine("    " + varName + "->Strikethrough = " + std::to_string(bStrike) + ";");
        };
        for (const auto& blk : uc.propertyBlocks) {
            if (_stricmp(blk.blockName.c_str(), "Font") != 0) continue;
            emitUcFrmFontObj(blk, "_vb6_f119");
            c_.emitLine("    vb6_UC_SetPendingFont(_vb6_f119);");
            c_.emitLine("  } }");
        }
        bool ucIsArr = knownControlArrays_.count(ctlLower) != 0;
        bool ucSkipInitProps = false;
        // 账 #305 第二百五十八刀: 与顶层那条创建路同一问 (invisible_at_runtime.hpp)。
        const FrmFile* spUc258 = findUserControlSpec(ctlClass);
        const bool ucWindowless258 = ucInvisibleAtRuntime(spUc258);
        if (spUc258)
            ucSkipInitProps = Symbol::toLower(spUc258->codeSection)
                                  .find("sub usercontrol_readproperties") != std::string::npos;
        std::string hwndExpr;
        if (ucIsArr) {
            int idx = uc.index >= 0 ? uc.index : 0;
            if (ucSkipInitProps) c_.emitLine("vb6_UC_SetSkipInitProps(1);");
            if (ucWindowless258) c_.emitLine("vb6_UC_SetWindowless(1);");
            c_.emitLine("{ void* vb6_uc_h = vb6_UC_HostCreate(\"" + ctlClass + "\", "
                + std::to_string(ucL) + ", " + std::to_string(ucT) + ", "
                + std::to_string(ucW) + ", " + std::to_string(ucH)
                + ", " + parentHwnd + ", hInstance, \"" + cIdent(uc.controlName) + "_" + std::to_string(idx)
                + "\", " + std::to_string(idx) + ");");
            c_.emitLine("  vb6_CtrlArr_SetAt(&vb6_arr_" + cIdent(uc.controlName) + ", "
                       + std::to_string(idx) + ", vb6_uc_h); }");
            if (ucWindowless258) c_.emitLine("vb6_UC_SetWindowless(0);");
            if (ucSkipInitProps) c_.emitLine("vb6_UC_SetSkipInitProps(0);");
            hwndExpr = "vb6_CtrlArr_GetAt(&vb6_arr_" + cIdent(uc.controlName) + ", "
                     + std::to_string(idx) + ")";
        } else {
            std::string slot220 = (uc.controlType == FrmControlType::ImageList
                                   ? "vb6_com_" : "vb6_hwnd_")
                                + cIdent(uc.controlName);
            if (ucSkipInitProps) c_.emitLine("vb6_UC_SetSkipInitProps(1);");
            if (ucWindowless258) c_.emitLine("vb6_UC_SetWindowless(1);");
            c_.emitLine(slot220 + " = vb6_UC_HostCreate(\""
                + ctlClass + "\", " + std::to_string(ucL) + ", " + std::to_string(ucT) + ", "
                + std::to_string(ucW) + ", " + std::to_string(ucH)
                + ", " + parentHwnd + ", hInstance, \"" + uc.controlName + "\", -1);");
            if (ucWindowless258) c_.emitLine("vb6_UC_SetWindowless(0);");
            if (ucSkipInitProps) c_.emitLine("vb6_UC_SetSkipInitProps(0);");
            hwndExpr = slot220;
        }
        {
            static const char* kSkip[] = {"Left", "Top", "Width", "Height", "Index",
                "TabIndex", "TabStop", "Visible", "Enabled", "Name", "_ExtentX",
                "_ExtentY", "Tag", "ToolTipText", "WhatisThis", "HelpContextID",
                "Font", "IconFont", "TitleFont", "Caption1_Font", "Caption2_Font",
                "Images", "ListImages", "MouseIcon"};
            c_.emitLine("vb6_UC_Enter(" + hwndExpr + ");");
            std::string ucRpCodeLower;
            if (const FrmFile* spRp = findUserControlSpec(ctlClass))
                ucRpCodeLower = Symbol::toLower(spRp->codeSection);
            for (const auto& blk : uc.propertyBlocks) {
                std::vector<ParameterInfo> fp119;
                bool isSet119 = false;
                if (findClassMemberWriteParams(ctlClass, blk.blockName, false, fp119)
                    && !fp119.empty()) {
                    isSet119 = false;
                } else if (findClassMemberWriteParams(ctlClass, blk.blockName, true, fp119)
                           && !fp119.empty()) {
                    isSet119 = true;
                } else {
                    continue;
                }
                bool fontName119 = false;
                {
                    std::string bn = Symbol::toLower(blk.blockName);
                    if (bn == "font" || (bn.size() >= 4
                        && bn.compare(bn.size() - 4, 4, "font") == 0))
                        fontName119 = true;
                }
                if (!fontName119) continue;
                std::string canonF = canonicalClassMemberName(ctlClass, blk.blockName);
                if (canonF.empty()) canonF = blk.blockName;
                std::string setterF = "vb6_" + cIdent(ctlClass) + "_"
                                    + propFnSegment(isSet119 ? "prop_set_" : "prop_let_", canonF);
                emitUcFrmFontObj(blk, "_vb6_f119");
                if (isSet119) {
                    std::string getterF = "vb6_" + cIdent(ctlClass) + "_"
                                        + propFnSegment("prop_get_", canonF);
                    c_.emitLine("    vb6_UC_FontAssign((void*)" + getterF
                              + "((vb6_cls_" + cIdent(ctlClass) + "*)vb6_UC_InstanceOf("
                              + hwndExpr + ")), _vb6_f119);");
                } else {
                    c_.emitLine("    " + setterF + "((vb6_cls_" + cIdent(ctlClass)
                              + "*)vb6_UC_InstanceOf(" + hwndExpr + "), _vb6_f119);");
                }
                c_.emitLine("  } }");
            }
            for (const auto& kv : uc.properties) {
                const std::string& pn = kv.first;
                bool skip = false;
                for (const char* sk : kSkip) if (_stricmp(pn.c_str(), sk) == 0) { skip = true; break; }
                if (skip) continue;
                if (!ucRpCodeLower.empty()
                    && ucRpCodeLower.find(".readproperty(\"" + Symbol::toLower(pn) + "\"")
                           != std::string::npos) {
                    continue;
                }
                std::vector<ParameterInfo> wp114;
                if (!findClassMemberWriteParams(ctlClass, pn, false, wp114) || wp114.empty())
                    continue;
                std::string canonPn = canonicalClassMemberName(ctlClass, pn);
                if (canonPn.empty()) canonPn = pn;
                std::string setter = "vb6_" + cIdent(ctlClass) + "_"
                                   + propFnSegment("prop_let_", canonPn);
                const ParameterInfo& p114 = wp114.back();
                const std::string p114RefLower = Symbol::toLower(p114.typeRefName);
                const bool p114RealVariant =
                    p114.type == Vb6Type::Variant
                    && (p114.typeRefName.empty() || p114RefLower == "variant");
                std::string meExpr = "(vb6_cls_" + cIdent(ctlClass)
                                   + "*)vb6_UC_InstanceOf(" + hwndExpr + ")";
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
                    if (!isStr) continue;
                }
                if (isStr) {
                    if (textVal.size() >= 2 && textVal.front() == '"' && textVal.back() == '"')
                        textVal = textVal.substr(1, textVal.size() - 2);
                    std::string lit = "vb6_BSTR_FromStr(L\"" + escapeWideCString(textVal) + "\")";
                    if (p114RealVariant && p114.isByVal) {
                        lit = "vb6_VariantFromValue(" + lit + ")";
                    }
                    if (p114.isByVal) {
                        c_.emitLine(setter + "(" + meExpr + ", " + lit + ");");
                    } else {
                        c_.emitLine("{ BSTR _lp114 = " + lit + "; "
                                  + setter + "(" + meExpr + ", &_lp114); }");
                    }
                } else {
                    std::string valExpr = kv.second.rawText.empty()
                            ? std::to_string(kv.second.intValue) : kv.second.rawText;
                    if (p114RealVariant && p114.isByVal) {
                        valExpr = "vb6_VariantFromValue(" + valExpr + ")";
                    }
                    c_.emitLine(setter + "(" + meExpr + ", " + valExpr + ");");
                }
            }
            {
                auto itAtt114 = ucEventAttach_.find(ctrlElemKey(uc.controlName, uc.index));
                if (itAtt114 != ucEventAttach_.end() && !getenv("C3_NO_EVENTS")) {
                    std::string shortCls114 = ctlClass;
                    size_t dotCls114 = shortCls114.rfind('.');
                    if (dotCls114 != std::string::npos) shortCls114 = shortCls114.substr(dotCls114 + 1);
                    c_.emitLine("{ /* UC events */");
                    for (const auto& al : itAtt114->second.lines) c_.emitLine("  " + al);
                    c_.emitLine("  ((vb6_cls_" + cIdent(shortCls114) + "*)vb6_UC_InstanceOf(" + hwndExpr + "))->events = &" + itAtt114->second.sinkVar + ";");
                    c_.emitLine("}");
                }
            }
            {
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
                    // Fix: .frx 属性包块 (UserControl 持久化 PropertyBag) —— 交原始字节,
                    // 运行期由 RTL 的 PropertyBag.Contents 解析 (见 uc_propbag.c)。
                    auto emitFrxBagBlob = [&](const std::string& litName,
                                              const std::vector<uint8_t>& bytes) {
                        c_.emitLine("vb6_UC_BagPutBin(_bag, " + litName + ",");
                        std::string line = "\"";
                        for (unsigned char b : bytes) {
                            if (b == '"')        line += "\\\"";
                            else if (b == '\\')  line += "\\\\";
                            else if (b >= 0x20 && b < 0x7F) line += (char)b;
                            else {
                                line += '\\';
                                line += (char)('0' + ((b >> 6) & 7));
                                line += (char)('0' + ((b >> 3) & 7));
                                line += (char)('0' + (b & 7));
                            }
                            if (line.size() >= 900) { c_.emitLine(line + "\""); line = "\""; }
                        }
                        c_.emitLine(line + "\",");
                        c_.emitLine(std::to_string(bytes.size()) + ");");
                    };
                    for (const auto& kvb : uc.properties) {
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
                            FrxPropBagBlob pbB114 = FrxReader::readPropBagBlob(kvb.second.frxOffset);
                            if (pbB114.found) {
                                emitFrxBagBlob(litName, pbB114.bytes);
                            } else {
                            FrxTextData dsB114 = FrxReader::readDocString(kvb.second.frxOffset);
                            if (!dsB114.text.empty()) {
                                std::string tvB = dsB114.text;
                                if (tvB.size() >= 2 && tvB.front() == '"' && tvB.back() == '"')
                                    tvB = tvB.substr(1, tvB.size() - 2);
                                c_.emitLine("vb6_UC_BagPutStr(_bag, " + litName + ", L\"" + escapeWideCString(tvB) + "\");");
                            }
                            }
                        }
                    }
                    c_.emitLine("vb6_" + cIdent(ctlClass) + "_UC_ReadProps((vb6_cls_" + cIdent(ctlClass)
                          + "*)vb6_UC_InstanceOf(" + hwndExpr + "), &_bag);");
                    c_.emitLine("vb6_UC_PropBagFree(_bag); }");
                }
            }
            if (ucSkipInitProps) c_.emitLine("vb6_UC_ActivateHost(" + hwndExpr + ");");
        }
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
                // Fix <vbeclipse> 2026-10-12 (账 #257 第十八层): 嵌套的工程内 UserControl
                // (CoolBar 的 band 里的 ImageCombo 等) —— 先走这条, 否则下面
                // `if (!childClass) continue` 会把它跳过 (childClass==NULL)。
                if (findUserControlSpec(child.controlTypeName)) {
                    emitUserControlChild(child, containerHwnd);
                    c_.emitLine("vb6_HostObj_Register((void*)" + childHwndRead(child) + ", \""
                                + child.controlName + "\", \"" + child.controlTypeName + "\", 0, "
                                + std::to_string(child.index) + ");  /* VB identity */");
                    {
                        std::string childNameLower = child.controlName;
                        std::transform(childNameLower.begin(), childNameLower.end(), childNameLower.begin(), ::tolower);
                        knownFormControls_[childNameLower] = child.controlType;
                        knownFormControlOriginalNames_[childNameLower] = child.controlName;
                    }
                    emitControlFrxProps(child);
                    childCtrlId++;
                    if (!child.children.empty()) {
                        if (controlIsContainerType(child) || findUserControlSpec(child.controlTypeName))
                            self(self, child, childHwndRead(child));
                    }
                    continue;
                }
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
                // 账 #252: 第二条创建路同样要记这一次 VB 身份（容器里的控件以前一条也没有）。
                // 句柄表达式走既有出口 ctrlHwndExprForInit —— 与下面几档设计期属性同一个口径，
                // 控件数组的元素句柄在 vb6_arr_<名> 里，拼 vb6_hwnd_<名>_0 会 C2065。
                c_.emitLine("vb6_HostObj_Register((void*)" + ctrlHwndExprForInit(child) + ", \""
                            + child.controlName + "\", \"" + child.controlTypeName + "\", 0, "
                            + std::to_string(child.index) + ");  /* VB identity */");
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
                    // Fix <c3-menu3d>: 与顶层路径同口径 —— 容器里的 PictureBox/Image
                    // 设计期 AutoSize=True 同样要发 (RT 的 FromCom AutoSize 块读它)。
                    if (asIt162d != child.properties.end() &&
                        asIt162d->second.type == FrmValueType::Integer &&
                        asIt162d->second.intValue != 0 &&
                        (child.controlType == FrmControlType::PictureBox ||
                         child.controlType == FrmControlType::Image)) {
                        c_.emitLine("vb6_SetPictureAutoSize((void*)" + childHwndRead(child)
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
                if (!child.children.empty() && (controlIsContainerType(child)
                        || findUserControlSpec(child.controlTypeName))) {
                    self(self, child, childHwndRead(child));   // 账 #167：容器清单收成一处
                }
            }
        };
        for (const auto& ctrl : frmDesc.formControl.children) {
            if (ctrl.children.empty()) continue;
            // Fix <vbeclipse> 2026-10-12 (账 #257 第十八层): 工程内 UserControl 容器
            // (CoolBar 等) 也走这条, 其内部嵌套的 UserControl 才能被 emiUserControlChild 创建。
            if (!controlIsContainerType(ctrl) && !findUserControlSpec(ctrl.controlTypeName)) continue;   // 账 #167
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
