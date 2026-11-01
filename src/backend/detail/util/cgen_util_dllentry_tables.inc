// cgen_util_dllentry_tables.inc - generateDllEntry 函数体片段：方法描述表、IID/事件常量表与 coclass 描述表
// 由 src/backend/cgen_util.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 398~623 行，逐行未改。

    // 2. IDispatch方法描述表
    for (auto& cc : coClasses) {
        std::string clsId = cIdent(cc.moduleName);
        std::string methodsVar = "g_vb6_disp_" + clsId + "Methods";

        // Fix 099: 有 Public 字段 (无方法) 时也要生成表 —— 字段被暴露为 Property Get/Let
        if (cc.publicMethodNames.empty() && cc.publicFields.empty()) {
            continue;  // No public methods/fields - skip method table
        }

        // M29: dispid取自methodSym->comDispid (来自driver TypeLib全局连续分配), nextDispid仅作fallback
        std::unordered_map<std::string, int> dispIdMap;
        int nextDispid = 1;

        // 账 #305 §B219(第二百二十刀): 表行要引用一枚**先于表**定义好的声明档数组(文件作用域，
        // 表体内不能塞定义)，所以方法那一趟先把数组行与表行各自攒下来，再按「数组 → 表头 → 行」落。
        std::vector<std::string> argVtArrayLines;
        std::vector<std::string> methodRowLines;

        for (size_t mi = 0; mi < cc.publicMethodNames.size(); mi++) {
            std::string& methodName = cc.publicMethodNames[mi];
            Symbol* methodSym = cc.publicMethodSyms[mi];
            std::string prefix = cc.moduleName + "_";
            std::string bareName = methodName;
            if (Symbol::toLower(bareName).substr(0, prefix.size()) == Symbol::toLower(prefix)) {
                bareName = bareName.substr(prefix.size());
            }
            std::string invokeSuffixMD;
            if (methodSym->kind == SymbolKind::PropertyGet) invokeSuffixMD = "_get";
            else if (methodSym->kind == SymbolKind::PropertyLet) invokeSuffixMD = "_let";
            else if (methodSym->kind == SymbolKind::PropertySet) invokeSuffixMD = "_set";
            std::string invokeName = "vb6_disp_" + clsId + "_" + cIdent(bareName) + invokeSuffixMD + "_invoke";

            int invkind = 1;
            if (methodSym->kind == SymbolKind::PropertyGet) {
                invkind = 2;
            } else if (methodSym->kind == SymbolKind::PropertyLet) {
                invkind = 4;
            } else if (methodSym->kind == SymbolKind::PropertySet) {
                invkind = 8;
            }

                        // COM convention: Property Get/Let/Set share the same DISPID
            std::string lowerBareName = Symbol::toLower(bareName);
            // M29: dispid优先取自methodSym->comDispid (driver TypeLib阶段回写), 保证dll_entry.c与TypeLib dispid一致
            //      COM convention: Property Get/Let/Set共享同一dispid (driver已确保同名Property变体共享同一comDispid)
            int dispid = methodSym->comDispid;
            if (dispid == 0) {
                // 安全fallback: 理论上不应发生 (driver TypeLib阶段已经回写comDispid) 仅防意外崩溃
                dispid = nextDispid++;
                dispIdMap[lowerBareName] = dispid;
            } else if (dispIdMap.find(lowerBareName) == dispIdMap.end()) {
                dispIdMap[lowerBareName] = dispid;
            }
std::string wideName = "L\"" + bareName + "\"";

            // 账 #305 §B219: 这一枚方法的形参声明档表(服务端进门那一趟按它换算实参)。
            // 档位答案只住 cDispatchArgVt 那一处，这里只负责把它落成一行静态数组。
            // 空形参表 = 不发数组、表行交 NULL(与 §B219 之前「没人问」同一形状)。
            std::string argVtRef = "NULL";
            {
                const size_t nPrm = methodSym->params.size();
                if (nPrm > 0) {
                    std::string init;
                    for (size_t pi = 0; pi < nPrm; pi++) {
                        if (!init.empty()) init += ", ";
                        init += cDispatchArgVt(methodSym->params[pi]);
                    }
                    // 名字直接从 invokeName 派生 —— invokeName 本来就是全模块唯一的真符号，
                    // 再拼一遍类名/后缀就是第二套命名口径。
                    argVtRef = "vb6_avt_" + invokeName;
                    argVtArrayLines.push_back("static const VARTYPE " + argVtRef +
                                              "[" + std::to_string((int)nPrm) + "] = { " + init + " };");
                }
            }

            methodRowLines.push_back("{ " + wideName + ", " + std::to_string((int32_t)dispid) +
                        ", " + std::to_string(invkind) + ", " + invokeName + ", " + argVtRef + ", " +
                        (argVtRef == "NULL" ? std::string("0")
                             : std::to_string((int)methodSym->params.size())) + " },");
        }

        for (auto& arrayLine : argVtArrayLines) entry.emitLine(arrayLine);
        entry.emitLine("static const vb6_DispMethodDesc " + methodsVar + "[] = {");
        entry.indent();
        for (auto& rowLine : methodRowLines) entry.emitLine(rowLine);

        // Fix 099: Public 字段条目 —— 每字段两条 (invkind 2=PropertyGet / 4=PropertyPut),
        // 共享同一 dispid (COM 约定: 同名 Property 的 Get/Let 同 dispid, TypeLib 侧亦然).
        for (size_t fi = 0; fi < cc.publicFields.size(); fi++) {
            const std::string& fname = cc.publicFields[fi];
            int32_t fdispid = (fi < cc.publicFieldDispids.size()) ? cc.publicFieldDispids[fi] : 0;
            if (fdispid == 0) {
                // 防御: TypeLib 阶段未回写 (理论上不应发生). 用高位段避开方法区
                // 的连续 dispid (TypeLib 从 1 开始分配), 否则 Invoke 按 dispid
                // 线性匹配会命中错误的表项.
                fdispid = 9000 + (int32_t)fi;
            }
            std::string fInvoke = "vb6_disp_" + clsId + "_" + cIdent(fname);
            // 账 #305 §B219: 字段那一腿的桥接把**整槽 vb6_VARIANT** 交给 _field_let_(它自己按
            // 字段档装箱)，认档那一问早就住在 cgen_com.cpp 的发码里 ⇒ 表行交 NULL, 0 = 服务端不问。
            entry.emitLine("{ L\"" + fname + "\", " + std::to_string(fdispid) +
                           ", 2, " + fInvoke + "_get_invoke, NULL, 0 },");
            entry.emitLine("{ L\"" + fname + "\", " + std::to_string(fdispid) +
                           ", 4, " + fInvoke + "_let_invoke, NULL, 0 },");
        }

        entry.dedent();
        entry.emitLine("};");
        entry.emitBlank();
    }

    // generateIid moved above (before collection loop)

    // 收集所有IID字符串 (interfaceName -> IID字符串)，避免重复
    std::unordered_map<std::string, std::string> iidMap;
    for (auto& cc : coClasses) {
        for (auto& ifaceName : cc.implementsNames) {
            if (iidMap.find(ifaceName) != iidMap.end()) continue;
            std::string lower = Symbol::toLower(ifaceName);
            // ai/022 B13c: 新式接口（Pass E 的接口 IID 表里有名）取表里的值 —— 与类自己的
            // vtable QI、类型库三处同值；表里没有（legacy `Implements` 的类模块）才落到
            // prelude 那枚 legacy 派生，存量工程的产物逐字节不变。
            // （B13b 在这里还有一条"看它是不是块 `[Default]` 那个"的分支：B13e 起对外默认接口
            // 不再从出口取，那条分支没有对象可管了，跟着一起撤 —— 见 collect.inc 同批注释。）
            if (iidreg_) {
                auto mapIt = iidreg_->find(lower);
                if (mapIt != iidreg_->end() && !mapIt->second.empty()) {
                    iidMap[ifaceName] = mapIt->second;
                    continue;
                }
            }
            iidMap[ifaceName] = generateIid("iface:" + progId + "." + ifaceName);
        }
    }

    // 生成静态IID常量 (仅在有Implements接口时生成)
    if (!iidMap.empty()) {
        entry.emitLine("// P12.1: Implements interface IIDs");
        for (auto& [ifaceName, iidStr] : iidMap) {
            std::string iidVar = "IID_vb6iface_" + cIdent(ifaceName);
            // 解析IID字符串并生成静态IID常量
            // 格式: {0xAABBCCDD,0xEEFF,0x1122,{0x33,0x44,0x55,0x66,0x77,0x88,0x99,0xAA}}
            uint32_t d1; uint16_t d2, d3; uint8_t d4[8];
            sscanf(iidStr.c_str(), "{%08X-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}",
                   &d1, &d2, &d3, &d4[0], &d4[1], &d4[2], &d4[3], &d4[4], &d4[5], &d4[6], &d4[7]);
            entry.emitLine("static const IID " + iidVar + " = {0x" + 
                StringFormatHex8(d1) + ",0x" + StringFormatHex4(d2) + ",0x" + StringFormatHex4(d3) +
                ",{0x" + StringFormatHex2(d4[0]) + ",0x" + StringFormatHex2(d4[1]) +
                ",0x" + StringFormatHex2(d4[2]) + ",0x" + StringFormatHex2(d4[3]) +
                ",0x" + StringFormatHex2(d4[4]) + ",0x" + StringFormatHex2(d4[5]) +
                ",0x" + StringFormatHex2(d4[6]) + ",0x" + StringFormatHex2(d4[7]) + "}};");
        }
        entry.emitBlank();

        // 生成每个coclass的ifaceIids数组
        for (auto& cc : coClasses) {
            if (cc.implementsNames.empty()) continue;
            std::string clsId = cIdent(cc.moduleName);
            std::string iidsVar = "g_vb6_ifaceIids_" + clsId;
            entry.emitLine("static const IID* const " + iidsVar + "[] = {");
            entry.indent();
            for (auto& ifaceName : cc.implementsNames) {
                entry.emitLine("&IID_vb6iface_" + cIdent(ifaceName) + ",");
            }
            entry.dedent();
            entry.emitLine("};");
            entry.emitBlank();
        }
    }
    // P6.6.4: 生成事件描述表和source IID常量
    bool hasAnyEvents = false;

    // P6.6.6: Default dispinterface IID constants (用于早绑定QI匹配)
    {
        bool hasDefaultIid = false;
        for (auto& cc : coClasses) {
            if (!cc.defaultIfaceIid.empty()) { hasDefaultIid = true; break; }
        }
        if (hasDefaultIid) {
            entry.emitLine("// P6.6.6: Default dispinterface IID constants");
            for (auto& cc : coClasses) {
                if (cc.defaultIfaceIid.empty()) continue;
                std::string clsId = cIdent(cc.moduleName);
                std::string iidVar = "IID_vb6def_" + clsId;
                uint32_t d1; uint16_t d2, d3; uint8_t d4[8];
                sscanf(cc.defaultIfaceIid.c_str(), "{%08X-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}",
                       &d1, &d2, &d3, &d4[0], &d4[1], &d4[2], &d4[3], &d4[4], &d4[5], &d4[6], &d4[7]);
                entry.emitLine("static const IID " + iidVar + " = {0x" +
                    StringFormatHex8(d1) + ",0x" + StringFormatHex4(d2) + ",0x" + StringFormatHex4(d3) +
                    ",{0x" + StringFormatHex2(d4[0]) + ",0x" + StringFormatHex2(d4[1]) +
                    ",0x" + StringFormatHex2(d4[2]) + ",0x" + StringFormatHex2(d4[3]) +
                    ",0x" + StringFormatHex2(d4[4]) + ",0x" + StringFormatHex2(d4[5]) +
                    ",0x" + StringFormatHex2(d4[6]) + ",0x" + StringFormatHex2(d4[7]) + "}};");
            }
            entry.emitBlank();
        }
    }
    for (auto& cc : coClasses) {
        if (!cc.eventNames.empty()) { hasAnyEvents = true; break; }
    }
    if (hasAnyEvents) {
        entry.emitLine("// P6.6.4: Source dispinterface IID constants");
        for (auto& cc : coClasses) {
            if (cc.sourceIfaceIid.empty()) continue;
            std::string clsId = cIdent(cc.moduleName);
            std::string iidVar = "IID_vb6src_" + clsId;
            uint32_t d1; uint16_t d2, d3; uint8_t d4[8];
            sscanf(cc.sourceIfaceIid.c_str(), "{%08X-%04hX-%04hX-%02hhX%02hhX-%02hhX%02hhX%02hhX%02hhX%02hhX%02hhX}",
                   &d1, &d2, &d3, &d4[0], &d4[1], &d4[2], &d4[3], &d4[4], &d4[5], &d4[6], &d4[7]);
            entry.emitLine("static const IID " + iidVar + " = {0x" +
                StringFormatHex8(d1) + ",0x" + StringFormatHex4(d2) + ",0x" + StringFormatHex4(d3) +
                ",{0x" + StringFormatHex2(d4[0]) + ",0x" + StringFormatHex2(d4[1]) +
                ",0x" + StringFormatHex2(d4[2]) + ",0x" + StringFormatHex2(d4[3]) +
                ",0x" + StringFormatHex2(d4[4]) + ",0x" + StringFormatHex2(d4[5]) +
                ",0x" + StringFormatHex2(d4[6]) + ",0x" + StringFormatHex2(d4[7]) + "}};");
        }
        entry.emitBlank();

        entry.emitLine("// P6.6.4: Event description tables");
        for (auto& cc : coClasses) {
            if (cc.eventNames.empty()) continue;
            std::string clsId = cIdent(cc.moduleName);
            std::string eventsVar = "g_vb6_events_" + clsId;
            entry.emitLine("static const vb6_EventDesc " + eventsVar + "[] = {");
            entry.indent();
            for (size_t ei = 0; ei < cc.eventNames.size(); ei++) {
                int32_t dispid = (int32_t)(ei + 1);  // 事件DISPID从1递增
                entry.emitLine("{ L\"" + cc.eventNames[ei] + "\", " + std::to_string(dispid) + " },");
            }
            entry.dedent();
            entry.emitLine("};");
            entry.emitBlank();
        }
    }

    // 3. 全局coclass描述表
    entry.emitLine("const vb6_CoClassDesc g_vb6_coclasses[] = {");
    entry.indent();
    // 一行 = 一个 ProgID 档。B13b 起同一份契约可以占两档：类模块名档（存量口径）
    // 与 CoClass 块名档（手写块 + [ComCreatable(True)]），CLSID 相同、工厂相同。
    auto emitCoclassRow = [&](const CoClassInfo& cc, const std::string& rowProgId) {
        std::string clsId = cIdent(cc.moduleName);
        std::string clsStruct = "vb6_cls_" + clsId;
        entry.emitLine("{");
        entry.indent();
        entry.emitLine("\"" + rowProgId + "\",  /* progId */");
        entry.emitLine("\"" + cc.clsidStr + "\",  /* clsidStr */");
        entry.emitLine("\"" + cc.moduleName + "\",  /* classVariable */");
        entry.emitLine("(void*(*)(void))" + clsStruct + "_New,  /* factoryFunc */");
        entry.emitLine("(void(*)(void*))" + clsStruct + "_Destroy,  /* destroyFunc */");
        entry.emitLine("NULL,  /* dispatchVtable */");
        // Fix 099: Public 字段占 2 条 (Get + Let), methodCount 要把它们算进去 —
        // GetIDsOfNames / Invoke 都是遍历 methods[0..methodCount) 匹配的.
        if (cc.publicMethodNames.empty() && cc.publicFields.empty()) {
            entry.emitLine("0,  /* methodCount */");
            entry.emitLine("NULL,  /* methods */");
        } else {
            size_t totalMethodEntries = cc.publicMethodNames.size() + cc.publicFields.size() * 2;
            entry.emitLine(std::to_string(totalMethodEntries) + ",  /* methodCount */");
            entry.emitLine("g_vb6_disp_" + clsId + "Methods,  /* methods */");
        }
        // P12.1: Implements interface IIDs
        if (!cc.implementsNames.empty()) {
            entry.emitLine(std::to_string(cc.implementsNames.size()) + ",  /* ifaceCount */");
            entry.emitLine("g_vb6_ifaceIids_" + clsId + ",  /* ifaceIids */");
        } else {
            entry.emitLine("0,  /* ifaceCount */");
            entry.emitLine("NULL,  /* ifaceIids */");
        }
        // P6.6.6: 默认接口IID (早绑定QI用)
        if (!cc.defaultIfaceIid.empty()) {
            entry.emitLine("&IID_vb6def_" + clsId + ",  /* defaultIfaceIid */");
        } else {
            entry.emitLine("NULL,  /* defaultIfaceIid */");
        }
        // P6.6.4: 事件源信息
        if (!cc.sourceIfaceIid.empty()) {
            entry.emitLine("\"" + cc.sourceIfaceIid + "\",  /* sourceIfaceIid */");
            entry.emitLine(std::to_string(cc.eventNames.size()) + ",  /* eventCount */");
            entry.emitLine("g_vb6_events_" + clsId + ",  /* events */");
        } else {
            entry.emitLine("NULL,  /* sourceIfaceIid */");
            entry.emitLine("0,  /* eventCount */");
            entry.emitLine("NULL,  /* events */");
        }
        // ai/022 B16: 新式接口那一面（薄指针 + 包装器放手桩）。两个指针同生同灭，
        // 都只在"实现新式接口"的类上非空；其余类为 NULL = QI 与销毁走原来的胖路。
        if (cc.hasNewStyleIface) {
            entry.emitLine("vb6_iv_thin_" + clsId + ",  /* ifaceThinPtr */");
            entry.emitLine("vb6_iv_claim_" + clsId + ",  /* instanceClaimRelease */");
        } else {
            entry.emitLine("NULL,  /* ifaceThinPtr (无新式接口: QI 走胖路) */");
            entry.emitLine("NULL,  /* instanceClaimRelease */");
        }
        entry.dedent();
        entry.emitLine("},");
    };
    for (auto& cc : coClasses) {
        emitCoclassRow(cc, cc.progId);
        if (!cc.groupProgId.empty()) {
            emitCoclassRow(cc, cc.groupProgId);
            extraGroupRows++;
        }
    }
    entry.dedent();
    entry.emitLine("};");
    entry.emitBlank();
