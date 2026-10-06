// driver_codegen_module_loop.inc - runCodeGeneration 函数体片段：逐模块代码生成主循环
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 371~462 行，逐行未改。


    // Fix 142: 确定启动对象 (决定哪个模块生成进程入口点).
    // VBP 未写 Startup= 时按 VB6 语义取工程中的第一个窗体模块作为启动窗体;
    // 全无窗体则留空 (此时只有含 Public Sub Main 的模块生成 main).
    std::string startupObject = startupObject_;
    if (startupObject.empty() && modules_.size() > 1) {
        for (size_t k = 0; k < modules_.size(); k++) {
            if (frmFiles_.count(modules_[k]->moduleName)) {
                startupObject = modules_[k]->moduleName;
                break;
            }
        }
    }

    // Fix 143: OCX 文件表 (vbp Object= 收集, CLSID → ocx 绝对路径)
    const std::map<std::string, std::string>& ocxFiles = ocxFiles_;
    const std::vector<std::pair<std::string, std::string>>& ocxRefs = ocxRefs_;

    // Fix 160: ComLib= 免注册组件表 {ProgID, CLSID, coclass名, 相对exe路径}
    const std::vector<std::array<std::string, 4>>& comLibRefs = comLibRefs_;

    // Fix <vbeclipse> 2026-10-06: Object= OCX 免注册组件表 {ProgID, CLSID, coclass名, 相对exe路径}
    const std::vector<std::array<std::string, 4>>& ocxLibRefs = ocxLibRefs_;

    for (size_t i = 0; i < modules_.size(); i++) {
        auto& module = modules_[i];
        auto& analyzer = analyzers_[i];

        // 泛型模板类 (tB, G4): 本体不发码 (成员 As T 无法编译); 特化副本作为
        // 独立 Module 已在 runGenericsPrepass 注入 modules_, 各自发码.
        if (!module->classTypeParams.empty()) continue;

        // 确定输出基名
        std::string baseName;
        if (modules_.size() == 1 && !options.outputFile.empty()) {
            // 单文件: 用输出文件名作为基名
            std::filesystem::path p(utf8ToPath(options.outputFile));
            baseName = pathToUtf8(p.stem());
        } else {
            // Fix 013: 多文件: 用 module.moduleName (VB_Name) 作为基名
            // 确保输出文件名与 #include 指令、跨模块 sourceModule 一致
            baseName = module->moduleName;
        }

        // 收集跨模块include需求（从符号表获取外部模块名）
        auto externalModules = analyzer->symbolTable().getExternalModuleNames();

        // Fix 075: 多模块项目中，所有模块的 isMultiModule_ 必须为 true，
        // 否则函数定义名不带模块前缀而跨模块调用带模块前缀，导致 LNK2019。
        // 将项目中所有其他模块名也加入 externalModules，确保 isMultiModule_=true。
        if (modules_.size() > 1) {
            for (size_t j = 0; j < modules_.size(); j++) {
                if (j != i) {
                    // 泛型模板类 (G4): 不发码 → 不进 include 依赖 (否则悬空 #include)
                    if (!modules_[j]->classTypeParams.empty()) continue;
                    externalModules.insert(modules_[j]->moduleName);
                }
            }
        }

        // 调用C代码生成器
        // Fix 023: 传入 void* 字段表供 cgen_expr.cpp fallback 路径查询
        CCodeGen cgen(*diag_, analyzer->symbolTable(), analyzer->typeSystem(),
                      &classVoidFieldMap, &classTypedFieldMap, &variantReturnFuncs, options.verbose);
        cgen.setFormModuleNames(formModuleNames);  // Fix 086: 跨模块窗体默认实例引用
        cgen.setStartupObject(startupObject);      // Fix 142: 启动对象 (入口点归属模块)
        cgen.setModulePublicConsts(&modulePublicConsts);  // Fix 086: 跨模块常量内联
        cgen.setDesignerFiles(&frmFiles_);  // Fix 112: 工程内 UserControl 设计器描述
        cgen.setOcxFiles(&ocxFiles);        // Fix 143: vbp Object= 的 ocx 路径表
        cgen.setOcxRefs(&ocxRefs);          // Fix 143b: 实例化 CLSID 以 vbp Object= 为准
        cgen.setComLibRefs(&comLibRefs);    // Fix 160: ComLib= 免注册组件表 (入口点烘焙)
        cgen.setOcxLibRefs(&ocxLibRefs);    // Fix <vbeclipse>: Object= OCX 免注册表 (入口点烘焙)
        cgen.setStaticModuleVarNames(&dupModuleVarNames);  // Fix 093b: 撞名模块级变量转 static
        cgen.setInterfaceImplementers(&ifaceImplementersMap);  // P6.4+: 接口唯一实现类映射
        cgen.setInterfaceAllImplementers(&ifaceAllImplementersMap);  // Fix <vbeclipse>: 接口全部实现类 (桩分派用)
        cgen.setClassChainRegistry(&classes_);  // tB 类继承 (B07b): stage 2.8 只读链登记表
        cgen.setProjectClassNames(&projectClassNames);  // 工程级类名表 (不受同名成员遮蔽)
        cgen.setInterfaceRegistry(&ifaces_);  // tB Interface 契约 (B04): stage 2.7 只读登记表
        cgen.setInterfaceIidMap(&ifaceIds_);  // tB 接口 IID 唯一出口 (B13c): 同一个 Pass E 建的表
        cgen.setViaRegistry(&vias_);          // tB 委托式实现 (B10): stage 2.7 Pass D 裁决表
        cgen.setDefaultInstanceClasses(&defaultInstanceClasses);  // P6.4+: 类模块默认实例
        cgen.setTrimIncludes(options.trimIncludes);  // opt4: 裁剪未实际引用的跨模块include
        cgen.setTargetArch(options.arch);            // ai/vb-asm-extension-spec: Asm 块仅 x64
        if (options.trimIncludes) {
            // opt4: 基于AST实际引用收集外部模块, 供 include 裁剪
            std::unordered_set<std::string> refdModules;
            ExternalRefCollector collector(analyzer->symbolTable(), externalModules, refdModules);
            traverseAST(*module, collector);
            cgen.setTrimModules(refdModules);
        }

        // 传入模块基名和外部模块列表
        // P6.6: 传递ActiveX DLL模式信息
        // P7: 传递窗体描述 (仅.frm模块有效)
        const FrmFormDesc* frmDescPtr = nullptr;
        auto frmIt = frmFiles_.find(module->moduleName);
        if (frmIt != frmFiles_.end()) {
            frmDescPtr = &frmIt->second.form;
        }
        bool ok = cgen.generate(*module, baseName, externalModules,
                                options.isDll, options.dllProgId, frmDescPtr);
        if (!ok) return false;

        // ai/vb-asm-extension-spec: 回收本模块的 Asm 过程 (体=单个 Asm 块), 交 runLinker 生成 .asm
        for (auto& ap : cgen.asmProcs()) {
            asmProcs_.push_back({baseName, ap});
        }

        // 写 .h 文件
        std::string hPath = outputDir + "/" + baseName + ".h";
        {
            std::ofstream ofs = ofstreamUtf8(hPath, std::ios::out | std::ios::trunc);
            if (!ofs) {
                std::cerr << "C3: 无法写入文件: " << hPath << std::endl;
                return false;
            }
            ofs << cgen.headerCode();
        }

        // 写 .c 文件
        std::string cPath = outputDir + "/" + baseName + ".c";
        {
            std::ofstream ofs = ofstreamUtf8(cPath, std::ios::out | std::ios::trunc);
            if (!ofs) {
                std::cerr << "C3: 无法写入文件: " << cPath << std::endl;
                return false;
            }
            ofs << cgen.sourceCode();
        }

        if (options.verbose) {
            std::cout << "C3: 生成 " << hPath << " (" << cgen.headerCode().size() << " bytes)" << std::endl;
            std::cout << "C3: 生成 " << cPath << " (" << cgen.sourceCode().size() << " bytes)" << std::endl;
        }

        // --emit-c 模式: 输出C代码后结束
        if (options.emitC) {
            std::cout << cgen.headerCode() << std::endl;
            std::cout << cgen.sourceCode() << std::endl;
        }
    }
