// driver_codegen_dll_entry.inc - runCodeGeneration 函数体片段：generateDllEntry 与 dll_entry.c 写出 + 返回（位于 if (options.isDll) 块内）
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 754~772 行，逐行未改。

        // P6.6.6: generateDllEntry() runs after TypeLib building so IIDs/CLSIDs are available
        std::string dllEntryCode = dllCgen.generateDllEntry(options.dllProgId, allSymTabs);

        std::string dllEntryPath = outputDir + "/dll_entry.c";
        {
            std::ofstream ofs = ofstreamUtf8(dllEntryPath, std::ios::out | std::ios::trunc);
            if (!ofs) {
                std::cerr << "C3: 无法写入文件: " << dllEntryPath << std::endl;
                return false;
            }
            ofs << dllEntryCode;
        }

        if (options.verbose) {
            std::cout << "C3: 生成 " << dllEntryPath << " (" << dllEntryCode.size() << " bytes)" << std::endl;
        }
    }

    // ExeComBridge 01: EXE 工程平行分支 — 生成 coclass 表 (com_entry.c), 与 DLL 的
    //   dll_entry.c 同构但跳过 DLL 导出函数段 (includeDllExports=false).
    //   阶段 1 仅让基础设施生成到位, 不接入任何打包点 (阶段 3), 不改变现有程序行为.
    //   progId 只影响 desc 字段与 CLSID/IID 的确定性哈希种子 (EXE 不注册, 无 ProgID
    //   反查): 优先 projectBaseName_ (driver_compile.cpp 赋值 = vbp 的 ExeName32
    //   stem 或 vbp 文件名 stem), 单文件编译时空, 退 options.dllProgId (EXE 模式
    //   driver_compile.cpp:116 不兜底赋值, 仅 CLI --progid 显式指定时有值),
    //   再退固定 "VB6EXE".
    //   注: classVoidFieldMap 等共享变量声明在 runCodeGeneration 的平铺片段区
    //   (voidfield_scan / typedfield_scan / variant_funcs / dup_module_vars 各
    //   .inc), 不在 DLL 的 if 块内, EXE 模式同样已初始化, 此处直接可用.
    //   注: DLL 分支的 P6.8 CLSID 注入与 dll_sync 的 IID/DISPID 同步在 if(isDll)
    //   块内, EXE 不执行 → comClsidStr/comDispid 走 generateDllEntry 的
    //   确定性哈希 fallback (collect.inc generateClsid/generateIid), 自洽.
    if (!options.isDll && !analyzers_.empty()) {
        auto& lastAnalyzer = analyzers_.back();
        CCodeGen comEntryCgen(*diag_, lastAnalyzer->symbolTable(), lastAnalyzer->typeSystem(),
                              &classVoidFieldMap, &classTypedFieldMap, &variantReturnFuncs, options.verbose);
        comEntryCgen.setFormModuleNames(formModuleNames);  // Fix 086
        comEntryCgen.setModulePublicConsts(&modulePublicConsts);  // Fix 086
        comEntryCgen.setDesignerFiles(&frmFiles_);  // Fix 112
        comEntryCgen.setStaticModuleVarNames(&dupModuleVarNames);  // Fix 093b: 撞名模块级变量转 static
        comEntryCgen.setInterfaceImplementers(&ifaceImplementersMap);  // P6.4+: 接口唯一实现类映射
        comEntryCgen.setInterfaceAllImplementers(&ifaceAllImplementersMap);  // Fix <vbeclipse>: 接口全部实现类 (桩分派用)
        comEntryCgen.setDefaultInstanceClasses(&defaultInstanceClasses);  // P6.4+: 类模块默认实例
        std::vector<SymbolTable*> allSymTabs;
        for (auto& analyzer : analyzers_) {
            allSymTabs.push_back(&analyzer->symbolTable());
        }
        std::string exeProgId = !projectBaseName_.empty() ? projectBaseName_
                                : (options.dllProgId.empty() ? std::string("VB6EXE") : options.dllProgId);
        std::string comEntryCode = comEntryCgen.generateDllEntry(exeProgId, allSymTabs, /*includeDllExports=*/false);

        std::string comEntryPath = outputDir + "/com_entry.c";
        {
            std::ofstream ofs = ofstreamUtf8(comEntryPath, std::ios::out | std::ios::trunc);
            if (!ofs) {
                std::cerr << "C3: 无法写入文件: " << comEntryPath << std::endl;
                return false;
            }
            ofs << comEntryCode;
        }

        if (options.verbose) {
            std::cout << "C3: 生成 " << comEntryPath << " (" << comEntryCode.size() << " bytes)" << std::endl;
        }
    }

    return !diag_->hasErrors();
