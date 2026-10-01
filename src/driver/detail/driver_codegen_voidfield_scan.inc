// driver_codegen_voidfield_scan.inc - runCodeGeneration 函数体片段：Fix 023 className → void* 字段集预扫描
// 由 src/driver/driver.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 100~187 行，逐行未改。

    // 工程级类名表 (小写类名 → 原大小写类名).
    //
    // 为什么必须有它: 每模块一张 SymbolTable, 而类名是工程级唯一的; 本模块作用域里
    // 任何同名成员 (字段自身 Variable / 同名 Sub / 同名 PropertyGet) 都会**遮蔽**
    // 跨模块类名 —— vbeclipse 实测:
    //   ucSplitBar.ctl `Private WithEvents SplitBar As SplitBar` → 类型名 "SplitBar"
    //     在本模块作用域里被字段自身的 Variable 符号占位, isVoidFieldType 落到
    //     "未知类型 → void*" 兜底 → 该字段被登记进 classVoidFieldMap_ → With 走
    //     后期绑定 COM dispatch, `.SplitterMouseDown hWnd, m_Rect, x, y` 的 RECT
    //     结构体实参被 vb6_ComPackInt 打包 → C2440 (ucSplitBar.c 175).
    //   ucCaption.ctl `Private WithEvents m_PopupMenu As PopupMenu` → 同名 Sub
    //     "PopupMenu" 占位 (SymbolKind::Sub), 后端 classVariantMembers_ 误注册 →
    //     vb6_VariantToObjectVal(me->m_PopupMenu) → C2440 (ucCaption.c 269).
    // 判"X 是不是工程类"只能问工程级的模块清单, 不能问任何单模块符号表.
    // (不能用 ClassChainRegistry: driver_classchain.cpp:355 的 D24 护栏在工程
    //  无任何 Inherits 时整表早退为空, vbeclipse 正是这种工程.)
    std::unordered_map<std::string, std::string> projectClassNames;
    for (const auto& mod : modules_) {
        if (!mod || !mod->isClassModule || mod->moduleName.empty()) continue;
        if (mod->isInterfaceModule) continue;  // 接口宿主块没有实例
        projectClassNames[Symbol::toLower(mod->moduleName)] = mod->moduleName;
    }
    auto isProjectClassName = [&projectClassNames](const std::string& name) {
        return !name.empty() && projectClassNames.count(Symbol::toLower(name)) > 0;
    };
    auto projectClassNameOf = [&projectClassNames](const std::string& name) {
        auto it = projectClassNames.find(Symbol::toLower(name));
        return it == projectClassNames.end() ? std::string() : it->second;
    };

    // Fix 023: 预扫描所有模块的 AST, 构建 className → void* 字段集 映射.
    // 这个映射在 cgen_expr.cpp 的 knownClassVars_ fallback 路径用于判断
    // obj->member 是否返回 void* COM 指针 (例: cDataBase.Rs As New ADODB.Recordset,
    // Rs 字段在 C 结构体中是 void*). 若是, 在 emit 的注释中加入 "voidptr" 标记,
    // 由外层 MemberAccessExpr 识别并切换为 COM dispatch (vb6_ComCall/ComGet*Prop).
    // 不依赖模块编译顺序 — driver 在主 codegen 循环前一次性扫描所有模块.
    std::unordered_map<std::string, std::set<std::string>> classVoidFieldMap;
    // void* 判定规则与 cgen_base.cpp::mapTypeRef 保持一致:
    //   - var.isNew (As New ClassName) → void*
    //   - vb6 内置对象类型 (Collection/ErrObject/...) → void*
    //   - 已知 ADODB 等 COM 对象类型 (Connection/Recordset/...) → void*
    //   - 兜底: 无法解析为 Class/ComClass/UDT/Enum/内置类型 的名称 → void*
    auto isVoidFieldType = [&isProjectClassName](const VariableDecl& var,
                              const TypeSystem& typeSys,
                              const SymbolTable& symTab) -> bool {
        if (var.isNew) return true;
        if (!var.asType || var.asType->kind != ASTNodeKind::SimpleTypeRef) return false;
        std::string typeName = static_cast<SimpleTypeRef*>(var.asType.get())->name;
        if (typeName == "Any" || typeName == "any") return true;
        if (typeName.size() > 4 && typeName.compare(0, 4, "VBA.") == 0) typeName = typeName.substr(4);
        // vbeclipse: 工程类字段的 C 类型也是 void* (类结构体指针), 但**不是**后期绑定
        // COM 对象 —— 方法调用必须走直接分发 (vb6_SplitBar_SplitterMouseDown(...)).
        // 本模块符号表查不到类名 (同名成员遮蔽, 见上方 projectClassNames 注释) 时会
        // 落到下面的"未知类型 → void*"兜底, 把项目类字段误登记进 classVoidFieldMap_,
        // With 遂走 COM dispatch. 先问工程级类名表.
        if (isProjectClassName(typeName)) return false;
        static const std::unordered_set<std::string> vb6BuiltinObj = {
            "Collection", "Forms", "ErrObject", "App", "Screen", "Printer", "Clipboard"
        };
        if (vb6BuiltinObj.count(typeName)) return true;
        // Fix 179: As Object 字段也是 void* (mapTypeRef 同映射), 必须进 void 表
        // (原先漏列, resolveTypeName 可解析走了非 void* 分支, 179a 目标判定失效)
        if (typeName == "Object" || typeName == "object") return true;
        Vb6Type t = typeSys.resolveTypeName(typeName);
        if (t != Vb6Type::Unknown) return false;  // 基础类型可解析 → 非 void*
        std::string lookupName = typeName;
        size_t dotPos = typeName.find('.');
        if (dotPos != std::string::npos) {
            std::string shortName = typeName.substr(dotPos + 1);
            auto* dotSym = symTab.lookupModule(shortName);
            if (dotSym) lookupName = shortName;
        }
        auto* clsSym = symTab.lookupModule(lookupName);
        if (clsSym) {
            if (clsSym->kind == SymbolKind::Class) return false;        // 真实类实例指针
            if (clsSym->kind == SymbolKind::ComClass ||
                clsSym->kind == SymbolKind::ComInterface) return false; // 类型化 COM 接口指针
        }
        auto* udtSym = symTab.lookup(lookupName);
        if (udtSym) {
            if (udtSym->kind == SymbolKind::UserDefinedType) return false;
            if (udtSym->kind == SymbolKind::EnumType) return false;
        }
        if (lookupName == "LongPtr" || lookupName == "Longptr") return false;
        if (lookupName.size() >= 2 && lookupName.compare(0, 2, "Vb") == 0) return false;
        if (lookupName.size() >= 4 && lookupName.compare(0, 4, "OLE_") == 0) return false;
        if (lookupName.size() >= 4 &&
            lookupName.compare(lookupName.size() - 4, 4, "Enum") == 0) return false;
        static const std::unordered_set<std::string> comObjTypes = {
            "Connection", "Recordset", "Command", "Parameter",
            "Field", "Fields", "Error", "Errors", "Property",
            "Properties", "Stream"
        };
        if (comObjTypes.count(lookupName)) return true;
        static const std::unordered_set<std::string> vb6EnumAliases = {
            "CompareMethod", "TriState", "FirstDayOfWeek", "FirstWeekOfYear",
            "MsgBoxResult", "MsgBoxStyle", "FileAttribute", "DateFormat",
            "Calendar", "DateTimeFormat", "CallType", "VariantType",
            "VarType", "QueryDef", "EditModeEnum", "FieldAttributeEnum"
        };
        if (vb6EnumAliases.count(lookupName)) return false;
        return true;  // 兜底: 未知类型 → void*
    };
    for (size_t i = 0; i < modules_.size(); i++) {
        const auto& module = modules_[i];
        // 类模块和窗体模块都有结构体字段; 标准模块无 struct 不需要扫描
        if (!module->isClassModule && !module->isFormModule) continue;
        const std::string& className = module->moduleName;
        std::set<std::string> voidFields;
        for (const auto& decl : module->declarations) {
            if (decl->kind != ASTNodeKind::VariableDecl) continue;
            const auto& var = static_cast<const VariableDecl&>(*decl);
            if (var.isDynamicArray || !var.dimensions.empty()) continue;  // 数组字段不是 void*
            if (!isVoidFieldType(var, analyzers_[i]->typeSystem(),
                                 analyzers_[i]->symbolTable())) continue;
            std::string oLower = var.name;
            std::transform(oLower.begin(), oLower.end(), oLower.begin(),
                           [](unsigned char c) { return (char)std::tolower(c); });
            std::string mLower = "m_" + oLower;
            voidFields.insert(oLower);
            voidFields.insert(mLower);
        }
        if (!voidFields.empty()) {
            classVoidFieldMap[className] = std::move(voidFields);
        }
    }

