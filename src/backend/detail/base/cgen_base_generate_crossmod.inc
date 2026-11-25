// cgen_base_generate_crossmod.inc - 跨模块注册与定义顺序 (在 cgen_base.cpp 的 generate 内被 #include)

    // Fix 010r-11: 外部类实例变量 → knownClassVars_
    // (ToolsStr As New cToolsStr → ToolsStr.HasStr 分发为 vb6_cToolsStr_HasStr)
    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
        if (sym->isExternal && sym->kind == SymbolKind::Variable && !sym->variableTypeName.empty()) {
            std::string varLower = sym->name;
            std::transform(varLower.begin(), varLower.end(), varLower.begin(), ::tolower);
            // 账 #305 第三百三十一刀: As <工程窗体类型> 的外部变量 (Public fMainForm As
            // frmMain 声明在标准模块, 消费方 frmMain/frmClassView/frmSolutionView 都要
            // 认得) —— 先于 Class/Object 各支登记 knownFormTypeVars_, 成员访问走
            // form_builtin 外部窗体支; 否则被下面的 Object 支收进 knownObjectVars_,
            // 发 vb6_ComCall(fMainForm, L"ActiveForm") (C2224 / 运行期双输)。
            if (knownFormModuleNames_.count(Symbol::toLower(sym->variableTypeName))) {
                knownFormTypeVars_[varLower] = sym->variableTypeName;
                continue;
            }
            auto* clsSym = lookupModuleDotted(sym->variableTypeName);
            if (clsSym && clsSym->kind == SymbolKind::Class && !clsSym->isInterface) {
                knownClassVars_[varLower] = clsSym->name;
                moduleClassVars_[varLower] = clsSym->name;
                // Fix 183: 外部 As New 变量同样需惰性实例化 — 标志随 Symbol 跨模块复制,
                // 注册进 moduleNewVars_ 后过程入口带入 knownNewVars_, 生成 if (!X) X = ... 守卫.
                if (sym->isNewVar) {
                    moduleNewVars_[varLower] = cIdent(clsSym->name);
                }
            }
            // Fix 089: 外部 Object 模块级变量 → 对应 COM 跟踪集合,
            // 否则 consuming 侧 knownObjectVars_ 为空 → 生成 void* 上的 .member → C2224.
            else if (sym->type == Vb6Type::Object) {
                if (clsSym && clsSym->kind == SymbolKind::ComClass) {
                    // Fix <vbeclipse>-2: 本工程有同名类模块 ⇒ 该外部变量其实是工程类实例
                    // (VB6: 工程内定义优先; 类型库自动加载注进来的内建 coclass
                    // isExternal=false). 走原生 knownClassVars_, 否则把原生
                    // vb6_cls_<Name>* 当 IDispatch 解 vtable → 0xC0000005.
                    const std::string projClsXm = projectClassNameOf(sym->variableTypeName);
                    if (!projClsXm.empty()) {
                        knownClassVars_[varLower] = projClsXm;
                        moduleClassVars_[varLower] = projClsXm;
                        if (sym->isNewVar) moduleNewVars_[varLower] = cIdent(projClsXm);
                    } else {
                        knownTypedComVars_[varLower] = clsSym;
                    }
                } else if (clsSym && clsSym->kind == SymbolKind::Class && clsSym->isInterface) {
                    knownIfaceVars_[varLower] = clsSym->name;
                } else {
                    // 账 #305 §11.108 (第一百二十六刀): 类模块的 Public 字段会被 stage 3.5
                    // 跨模块注入成 `isExternal=true, sourceModule=<类名>` 的「模块级变量」符号
                    // (实测 .temp/k125: FBox.cls 的 `Public A/B As Object` 在 FMain7 与 FC 两个
                    // 消费模块里都成了 a/b)。它们**不是**全局变量 —— 裸名在消费模块里根本不可写,
                    // 只有 `obj.A` 这种成员访问才成立 —— 让它们进 knownObjectVars_ 就是 §11.104
                    // 那条「名字表里混着别的类的成员名」的污染源。标准模块的 Public Object 变量
                    // (Fix 089) 仍走下面这一行。
                    auto* srcMod = sym->sourceModule.empty() ? nullptr : symTab_.lookup(sym->sourceModule);
                    const bool srcIsClassModule = srcMod && srcMod->kind == SymbolKind::Class;
                    if (!srcIsClassModule) {
                        knownObjectVars_.insert(varLower);
                    }
                }
            }
            else if (sym->isArray && sym->type == Vb6Type::UserDefinedType && !sym->variableTypeName.empty()) {
                // Cross-module Public UDT *array* (`Public Tracers(0 To XRES-1) As TTracePoint`).
                // Element access asks arrayUdtElemTypes_ by bare name; without an entry here
                // it fell back to mapSaElemCType(UserDefinedType) = vb6_VARIANT, i.e. a 16-byte
                // stride into an array whose elements are sizeof(UDT) apart. modRenderer read
                // modRaycaster's Tracers through VB6_SA_AT(vb6_VARIANT, ...) while the real
                // element was 48 bytes -> garbage trace points (grey stripe walls).
                arrayUdtElemTypes_[varLower] = "vb6_type_" + cIdent(sym->variableTypeName);
                if (sym->dimCount > 0) arrayDimCounts_[varLower] = sym->dimCount;
            }
            else if (sym->type == Vb6Type::UserDefinedType && !sym->variableTypeName.empty()) {
                // Cross-module Public UDT variable: register its C struct type so
                // `Player.ammo(i)` is emitted as an array index, not a function call.
                moduleUdtMembers_[varLower] = "vb6_type_" + cIdent(sym->variableTypeName);
            }
        }
    }

    // Fix 010r-12: Enum/UDT 定义须先于跨模块 #include, 否则循环引用时定义不可见.
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::EnumDecl) {
            visit(static_cast<EnumDecl&>(*decl));
        }
    }

    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::TypeDecl) {
            visit(static_cast<TypeDecl&>(*decl));
        }
    }

    // P8.7: 跨模块 #include 策略 — 类模块(或被类模块引用)走 .h, 标准→标准走 .c (避免 .h 循环)
    for (const auto& extMod : externalModules) {
        auto* extSym = symTab_.lookup(extMod);
        bool extIsClass = extSym && extSym->kind == SymbolKind::Class;
        if (extIsClass || isClassModule_) {
            h_.emitLine("#include \"" + extMod + ".h\"");
        }
    }
    h_.emitBlank();

