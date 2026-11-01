// driver_codegen_module_loop.inc - runCodeGeneration 函数体片段：逐模块代码生成主循环
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 371~462 行，逐行未改。


    // 账 §B97: 外部模块名单的**唯一排法** = modules_ 的下标序。⚠ 它**不是** .vbp 里
    // Module=/Class=/Form= 的字面先后 —— runParse 末尾有一步**刻意**的稳定前移（类模块排在
    // 标准模块之前，理由见 driver_frontend.cpp:327-339：标准模块里 `Dim WithEvents x As 某类`
    // 要那枚类符号先进表），所以这里的「序」= 前移之后的序（同类之间、窗体与标准模块之间
    // 仍保持 .vbp 的先后）。要点是它**有序且确定**，而不是照抄文件行序。
    // 这份名单往下决定三处输出的行序：
    // 聚合 .c 的 #include 段、.h 的 crossmod #include 段、入口点里 vb6_mod_<X>_init()
    // 的调用序（= 各模块全局的初始化先后）。以前它是 unordered_set，由符号表那枚
    // unordered_map 的遍历造出 ⇒ 「启动序」随编译器二进制漂（实测 tests/cc_id/Id.vbp：
    // HEAD 交 IdMain,VbpImpl,CircleImpl,IdIfaces,IdBlocks，本刀交 IdMain,CircleImpl,VbpImpl,
    // IdIfaces,IdBlocks）。口径：**要保的是确定性**，不是照抄 VB6 ——
    // VB6 的模块级变量只在首次引用时才建，本来就没有这条 eager 序，它是 C3 的实现选择。
    std::unordered_set<std::string> projectModNames;
    for (size_t j = 0; j < modules_.size(); j++) {
        if (modules_[j]) projectModNames.insert(modules_[j]->moduleName);
    }
    auto extModulesInModuleVectorOrder = [&](size_t self) {
        std::vector<std::string> out;
        for (size_t j = 0; j < modules_.size(); j++) {
            if (!modules_[j] || j == self) continue;
            // 泛型模板类 (G4): 不发码 → 不进 include 依赖 (否则悬空 #include)
            if (!modules_[j]->classTypeParams.empty()) continue;
            out.push_back(modules_[j]->moduleName);
        }
        return out;
    };

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

        // 收集跨模块include需求（账 §B97: 名单按工程声明序排，见文件头那两条）。
        // Fix 075 的语义原样保留：多模块工程里「所有其他模块」都必须在名单里，
        // 否则函数定义名不带模块前缀而跨模块调用带前缀 ⇒ LNK2019。
        auto externalModules = extModulesInModuleVectorOrder(i);
        // 兜底：符号表报了**不在工程表里**的名字（实测本项目内是空集）也留着，
        // 成员集合与改动前一致；getExternalModuleNames() 自身已升序，追加段因此也确定。
        for (const auto& nm : analyzer->symbolTable().getExternalModuleNames()) {
            if (!projectModNames.count(nm)) externalModules.push_back(nm);
        }

        // 账 #305 第318刀 (账 #594): 与 cProcName 同一条条件 —— externalModules 为空
        //   （本工程只有这一枚发码模块）且不是类模块时，本模块的函数发**裸名** vb6_<proc>；
        //   预扫描只登记了 vb6_<模块>_<proc> / vb6_pv_<proc> 两形，裸名这一支从没进过
        //   variantReturnFuncs ⇒ 单模块工程的 Variant 函数被 cExprIsVariant 当裸标量：
        //   `Select Case TagVal()` 发 `BSTR v = vb6_TagVal();`（cl C2440），
        //   `l = VariantFn()` 少一次解包（§B148 同形）。
        if (externalModules.empty() && !module->isClassModule) {
            auto bareIt = variantBareProcNames.find(module->moduleName);
            if (bareIt != variantBareProcNames.end()) {
                for (const auto& bareProc : bareIt->second) {
                    variantReturnFuncs.insert("vb6_" + bareProc);
                }
            }
        }

        // 调用C代码生成器
        // Fix 023: 传入 void* 字段表供 cgen_expr.cpp fallback 路径查询
        CCodeGen cgen(*diag_, analyzer->symbolTable(), analyzer->typeSystem(),
                      &classVoidFieldMap, &classTypedFieldMap, &variantReturnFuncs, options.verbose);
        cgen.setFormModuleNames(formModuleNames);  // Fix 086: 跨模块窗体默认实例引用
        // 账 #305 第320刀 (账 #592): App.Title 真值 = .vbp 的 Title（空则工程 Name）。
        // 单文件编译（无 .vbp）时为空 ⇒ 入口点一行不发，读取侧退 EXEName（旧口径）。
        cgen.setProjectTitle(!verTitle_.empty() ? verTitle_ : vbpProjectName_);
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
        cgen.setLegacyIfaceShapes(&legacyIfaceShapeMap);  // 账 #305 新账 #464: legacy 接口槽表的形状归接口自己定
        cgen.setClassChainRegistry(&classes_);  // tB 类继承 (B07b): stage 2.8 只读链登记表
        cgen.setProjectClassNames(&projectClassNames);  // 工程级类名表 (不受同名成员遮蔽)
        cgen.setProjectCoclassNames(&projectCoclassNames);  // 账 #305 §B187: 装箱腿问"这类在 coclass 表里有没有一行"
        cgen.setProjectImplementedNames(&projectImplementedNames);  // 账 #305 §B189 第二档: 接口视图那一族不许按声明档折 TypeName
        cgen.setInterfaceRegistry(&ifaces_);  // tB Interface 契约 (B04): stage 2.7 只读登记表
        cgen.setInterfaceIidMap(&ifaceIds_);  // tB 接口 IID 唯一出口 (B13c): 同一个 Pass E 建的表
        cgen.setViaRegistry(&vias_);          // tB 委托式实现 (B10): stage 2.7 Pass D 裁决表
        cgen.setDefaultInstanceClasses(&defaultInstanceClasses);  // P6.4+: 类模块默认实例
        cgen.setClassFieldKindMap(&classFieldKindMap);  // 账 #305 §B168/§B169: 工程级「类 → 标量字段 → 声明档」表
        cgen.setClassMemberArrayMap(&classMemberArrayMap);  // 账 #476 §B190: 「类 → 动态数组成员 → 声明语法」表 (只运语法, 档位在出口里映射)
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
        c3PhaseMark("gen " + baseName);   // ai/030 续: 逐模块发码计时 (含串料化)
        c3StatDump("mod " + baseName);    // ai/030 续: 本模块的热函数累计 (emitLine/infer/emitExpr)

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
        c3PhaseMark("io " + baseName);    // ai/030 续: .h/.c 落盘计时

        if (options.verbose) {
            std::cout << "C3: 生成 " << hPath << " (" << cgen.headerCode().size() << " bytes)" << std::endl;
            std::cout << "C3: 生成 " << cPath << " (" << cgen.sourceCode().size() << " bytes)" << std::endl;
        }

        // --emit-c 模式: 输出C代码后结束
        // 账 #267: 这份产物要和上面那两份 .h/.c **同一口径**（内部 UTF-8 原样），所以绕开 std::cout
        // 上那层「管道/文件就按控制台代码页转字节」的 streambuf —— 它是为给人看的诊断装的，
        // 用在产物上就变成「同一个问题两个答案，其中一个随机器变」：本机 936 交 GBK、CI 65001 交 UTF-8。
        // 先 flush 诊断，再直接 fwrite（stdout 是文本流，'\n' 照旧翻成 CRLF，与 ofstream 写的那两份一致）。
        if (options.emitC) {
            std::cout.flush();
            std::string dump = cgen.headerCode();
            dump += "\n";
            dump += cgen.sourceCode();
            dump += "\n";
            writeStdoutRaw(dump);
            std::fflush(stdout);
        }
    }
