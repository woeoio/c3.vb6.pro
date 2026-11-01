// driver_codegen_dll_sync.inc - runCodeGeneration 函数体片段：P6.6.6/M29 IID/CLSID/DISPID 符号同步（位于 if (options.isDll) 块内）
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 699~753 行，逐行未改。

        // P6.6.6: Sync comDefaultIfaceIid/comClsidStr from allSymTabs to lastAnalyzer's symbol table
        // (TypeLib builder writes to per-module Symbol objects, but cgen uses lastAnalyzer's merged table)
        for (auto& [key, sym] : lastAnalyzer->symbolTable().moduleScope()->symbols()) {
            if (sym->kind == SymbolKind::Class && !sym->isInterface && sym->comDefaultIfaceIid.empty()) {
                // Search all symbol tables for a matching class with IID info
                for (auto* st : allSymTabs) {
                    auto* srcSym = st->lookupModule(sym->name);
                    if (srcSym && srcSym->kind == SymbolKind::Class && !srcSym->comDefaultIfaceIid.empty()) {
                        sym->comDefaultIfaceIid = srcSym->comDefaultIfaceIid;
                        sym->comDefaultIfaceName = srcSym->comDefaultIfaceName;
                        sym->comDefaultMemberName = srcSym->comDefaultMemberName;
                        sym->comDefaultMemberRealName = srcSym->comDefaultMemberRealName;
                    }
                    if (srcSym && srcSym->kind == SymbolKind::Class && !srcSym->comSourceIfaceIid.empty() && sym->comSourceIfaceIid.empty()) {
                        sym->comSourceIfaceIid = srcSym->comSourceIfaceIid;
                        sym->comSourceIfaceName = srcSym->comSourceIfaceName;
                    }
                }
            }
        }

        // 账 #305 §B220 (第二百二十七刀): CLSID 这一格原先住在上面的「默认接口 IID 还是空」那道闸里 ——
        // 已经有默认成员的那两枚（View 的属性 / PopupMenu 的默认方法）整块被跳过，于是 cgen 那一趟读到
        // 的是另一枚同名 Class 符号（实测两个地址不同，且按名 lookupModule 取不到 tlb 写的那一枚），
        // 它只好自己 mint 一枚 ⇒ 与类型库广告的不是同一枚（DllGetClassObject 答 0x80040111）。
        // 这里**逐表逐符号**扫一遍按名补齐（不用 lookupModule：同名两份时它只回其中一份），
        // 只填空值、值只取自已经非空的那枚，绝不产生第二种 mint。与下面 M29 的 DISPID 同规。
        {
            std::unordered_map<std::string, std::string> knownClsidByName;
            for (auto* st : allSymTabs) {
                if (!st || !st->moduleScope()) continue;
                for (auto& kv2 : st->moduleScope()->symbols()) {
                    Symbol* s2 = kv2.second.get();
                    if (!s2 || s2->kind != SymbolKind::Class || s2->isInterface) continue;
                    if (!s2->comClsidStr.empty()) knownClsidByName[s2->name] = s2->comClsidStr;
                }
            }
            if (!knownClsidByName.empty()) {
                for (auto* st : allSymTabs) {
                    if (!st || !st->moduleScope()) continue;
                    for (auto& kv2 : st->moduleScope()->symbols()) {
                        Symbol* s2 = kv2.second.get();
                        if (!s2 || s2->kind != SymbolKind::Class || s2->isInterface) continue;
                        if (!s2->comClsidStr.empty()) continue;
                        auto it2 = knownClsidByName.find(s2->name);
                        if (it2 != knownClsidByName.end()) s2->comClsidStr = it2->second;
                    }
                }
            }
        }

        // M29: Sync comDispid from per-module method symbols to lastAnalyzer's merged table
        // (TypeLib builder writes comDispid to per-module Symbol* in allSymTabs, but cgen's
        //  generateDllEntry reads via lastAnalyzer.symbolTable().lookupModule[ByKind], which
        //  may return a DIFFERENT Symbol* than what driver TypeLib phase wrote to. Same pattern
        //  as the comDefaultIfaceIid sync block above.)
        for (auto& [mKey, mClsSym] : lastAnalyzer->symbolTable().moduleScope()->symbols()) {
            if (!mClsSym || mClsSym->kind != SymbolKind::Class || mClsSym->isInterface) continue;
            if (mClsSym->isExternal) continue;  // 跳过外部注入副本
            for (auto& memberName : mClsSym->memberNames) {
                // Sync each possible method kind independently (Sub/Function are exclusive;
                // Property Get/Let/Set share name but are distinct Symbol objects)
                for (auto mKind : {SymbolKind::Sub, SymbolKind::Function,
                                   SymbolKind::PropertyGet, SymbolKind::PropertyLet,
                                   SymbolKind::PropertySet}) {
                    // Find any Symbol* with comDispid set (the "source of truth" written by driver TypeLib phase)
                    int knownDispid = 0;
                    for (auto* st : allSymTabs) {
                        Symbol* s = st->lookupModuleByKind(memberName, mKind);
                        if (s && !s->isExternal && s->comDispid != 0) { knownDispid = s->comDispid; break; }
                    }
                    if (knownDispid == 0) continue;  // method not present, or driver didn't set dispid
                    // Propagate to ALL matching Symbol* across all symbol tables (only fill zeros)
                    // (cgen's generateDllEntry may read via a DIFFERENT Symbol* than what driver wrote to,
                    //  because lastAnalyzer's merged table can have its own copy. Update every copy.)
                    for (auto* st : allSymTabs) {
                        Symbol* s = st->lookupModuleByKind(memberName, mKind);
                        if (s && !s->isExternal && s->comDispid == 0) s->comDispid = knownDispid;
                    }
                }
            }
        }
