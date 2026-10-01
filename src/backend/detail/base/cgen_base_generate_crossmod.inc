// cgen_base_generate_crossmod.inc - 跨模块注册与定义顺序 (在 cgen_base.cpp 的 generate 内被 #include)

    // Fix 010r-11: 外部类实例变量 → knownClassVars_
    // (ToolsStr As New cToolsStr → ToolsStr.HasStr 分发为 vb6_cToolsStr_HasStr)
    for (const auto& [key, sym] : symTab_.moduleScope()->symbols()) {
        if (sym->isExternal && sym->kind == SymbolKind::Variable && !sym->variableTypeName.empty()) {
            std::string varLower = sym->name;
            std::transform(varLower.begin(), varLower.end(), varLower.begin(), ::tolower);
            auto* clsSym = lookupModuleDotted(sym->variableTypeName);
            if (clsSym && clsSym->kind == SymbolKind::Class && !clsSym->isInterface) {
                knownClassVars_[varLower] = clsSym->name;
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
                        if (sym->isNewVar) moduleNewVars_[varLower] = cIdent(projClsXm);
                    } else {
                        knownTypedComVars_[varLower] = clsSym;
                    }
                } else if (clsSym && clsSym->kind == SymbolKind::Class && clsSym->isInterface) {
                    knownIfaceVars_[varLower] = clsSym->name;
                } else {
                    knownObjectVars_.insert(varLower);
                }
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

