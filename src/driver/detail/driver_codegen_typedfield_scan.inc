// driver_codegen_typedfield_scan.inc - runCodeGeneration 函数体片段：Fix 037b typed 字段表 + Fix 086 窗体模块名/Public 常量表
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 188~271 行，逐行未改。

    // Fix 037b: 构建 classTypedFieldMap — 非 void* 的 typed 对象字段类型表.
    // 扫描每个类模块/窗体模块的字段声明, 对 isVoidFieldType 返回 false 但
    // 类型为 Class/ComClass/ComInterface 的字段, 记录其 VB6 类型名.
    // 用于 IndexOrCallExpr 中 obj.typedField(idx) 模式: typedField 是对象数据字段
    // (非函数), VB6 语义为调用其默认 Item 属性.
    std::unordered_map<std::string, std::unordered_map<std::string, std::string>> classTypedFieldMap;
    // Fix 086: 工程内窗体模块名 (小写) — 供跨模块窗体默认实例引用
    // (FLogs.Visible / Unload FLogs → vb6_form_hwnd_FLogs()) 解析。
    // 账 #278 §B110: 建造点搬进 Driver::collectFormModuleNames() —— 语义层那份"裸名位放行"的
    // 名单与这里必须同源, 两边各扫一遍 modules_ 就是两份权威 (改一边漏一边 = 又一个"产物对、诊断错")。
    std::unordered_set<std::string> formModuleNames = collectFormModuleNames();
    // Fix 086: 收集各模块 Public 常量表 (小写模块名 → 小写常量名 → 值).
    // 当本地同名过程符号 (如 cSerialPort.Sub SetDTR) 阻止了跨模块常量注入时,
    // Module.Const (modSerialPortAPI.SetDTR) 引用仍能内联数值.
    std::unordered_map<std::string, std::unordered_map<std::string, long long>> modulePublicConsts;
    for (size_t i = 0; i < modules_.size(); i++) {
        const std::string& modName = modules_[i]->moduleName;
        if (modName.empty()) continue;
        std::string modLower = Symbol::toLower(modName);
        auto* modScope = analyzers_[i]->symbolTable().moduleScope();
        if (!modScope) continue;
        for (const auto& [ckey, csym] : modScope->symbols()) {
            if (!csym || csym->kind != SymbolKind::Constant || !csym->hasConstValue) continue;
            if (csym->access != AccessLevel::Public) continue;
            if (csym->constType != Vb6Type::Long && csym->constType != Vb6Type::Integer
                && csym->constType != Vb6Type::Boolean && csym->constType != Vb6Type::Byte
                && csym->constType != Vb6Type::Error) continue;
            modulePublicConsts[modLower][Symbol::toLower(csym->name)] = csym->constIntValue;
        }
    }
    for (size_t i = 0; i < modules_.size(); i++) {
        const auto& module = modules_[i];
        if (!module->isClassModule && !module->isFormModule) continue;
        const std::string& className = module->moduleName;
        std::unordered_map<std::string, std::string> typedFields;
        for (const auto& decl : module->declarations) {
            if (decl->kind != ASTNodeKind::VariableDecl) continue;
            const auto& var = static_cast<const VariableDecl&>(*decl);
            if (var.isDynamicArray || !var.dimensions.empty()) continue;
            if (!var.asType || var.asType->kind != ASTNodeKind::SimpleTypeRef) continue;
            std::string typeName = static_cast<SimpleTypeRef*>(var.asType.get())->name;
            // 在符号表中查找类型, 区分项目类 vs COM 接口
            std::string lookupName = typeName;
            size_t dotPos = typeName.find('.');
            if (dotPos != std::string::npos) {
                std::string shortName = typeName.substr(dotPos + 1);
                auto* dotSym = analyzers_[i]->symbolTable().lookupModule(shortName);
                if (dotSym) lookupName = shortName;
            }
            auto* clsSym = analyzers_[i]->symbolTable().lookupModule(lookupName);
            std::string oLower = var.name;
            std::transform(oLower.begin(), oLower.end(), oLower.begin(), [](unsigned char c) { return (char)std::tolower(c); });
            std::string mLower = "m_" + oLower;
            // vbeclipse: 类型名被本模块同名成员遮蔽时 (字段自身 Variable / 同名 Sub /
            // 同名 PropertyGet) lookupModule 拿到的不是 Class, 或干脆查不到 → 该字段
            // 既不进 typedFields 也不被正确分类, With 退回 Unknown/COM 后期绑定.
            // 工程级类名表是唯一不受遮蔽的判据 (说明见
            // driver_codegen_voidfield_scan.inc 的 projectClassNames 注释).
            if ((!clsSym || clsSym->kind != SymbolKind::Class) &&
                isProjectClassName(lookupName)) {
                typedFields[oLower] = projectClassNameOf(lookupName);
                typedFields[mLower] = projectClassNameOf(lookupName);
                continue;
            }
            if (!clsSym) continue;
            if (clsSym->kind == SymbolKind::Class) {
                // Fix 090a: As New 项目类字段 (C 存储 void*) 也入 typed 表, 供
                // inferClassTypeOfExpr 推断 (me.FileStream → cToolsStream). 此前被
                // isVoidFieldType 跳过仅入 classVoidFieldMap_ 走 COM dispatch →
                // me.FileStream.ReadLine() 等有括号调用的 Optional 参数无法补齐
                // (callee 对象为 MemberAccessExpr(me,FileStream), Fix 042b 推断断链
                // → calleeParams 空 → C2198 参数太少).
                typedFields[oLower] = clsSym->name;
                typedFields[mLower] = clsSym->name;
                continue;
            }
            // 项目类外: void* 字段已在 classVoidFieldMap_ 走 COM dispatch
            if (isVoidFieldType(var, analyzers_[i]->typeSystem(),
                               analyzers_[i]->symbolTable())) continue;
            if (clsSym->kind == SymbolKind::ComClass ||
                clsSym->kind == SymbolKind::ComInterface) {
                typedFields[oLower] = "COM:" + typeName;  // COM 标记
                typedFields[mLower] = "COM:" + typeName;
            }
            // UDT/Enum 等非对象类型跳过
        }
        if (!typedFields.empty()) {
            classTypedFieldMap[className] = std::move(typedFields);
        }
    }

