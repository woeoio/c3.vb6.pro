// cgen_base_generate_prologue.inc - generate 函数体片段：入口初始化 + Fix 090q 返回 UDT 预扫
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 91~131 行，逐行未改。

    currentModule_ = &module;
    baseName_ = baseName;
    moduleName_ = module.moduleName;  // Fix 013: 模块名 = VB_Name (规范名称), 与 clsSym->name 一致
    isMultiModule_ = !externalModules.empty();  // 有外部依赖 = 多模块项目
    externalModules_ = externalModules;  // M22: 保存外部模块集合
    isClassModule_ = module.isClassModule;
    isFormModule_ = module.isFormModule;
    // Fix 110: .ctl/.pag 的设计期子控件模块 (非 .frm 但有设计器描述).
    // 这些模块的可见内容由代码绘制, 但 `Timer1.Interval` / `Picture1.ScaleWidth`
    // 等设计期子控件成员仍需句柄符号与控件名映射.
    isDesignerModule_ = (frmDesc != nullptr) && !module.isFormModule;
    isDll_ = isDll;  // P6.6
    dllProgId_ = dllProgId;  // P6.6
    emittedSymbols_.clear();
    labelCounter_ = 0;
    tempCounter_ = 0;

    // <vbeclipse>: 读本模块的 Option Compare (此前 module.options 只有语义层在看,
    // 代码生成层从不读 ⇒ `Option Compare Text` 编出来的比较仍是二进制, 且
    // cgen_localdecl.cpp 里那条 g_vb6_optionCompareText 赋值从来没机会发出)。
    optionCompareText_ = false;
    for (auto& opt : module.options) {
        if (opt->optionKind == OptionKind::CompareText) optionCompareText_ = true;
        else if (opt->optionKind == OptionKind::CompareBinary) optionCompareText_ = false;
    }

    // Fix 090q: 预扫模块内返回 UDT 的函数/方法 (小写名 → C 返回类型), 供
    // Declare As Any ByRef 打包识别「实参是返回 UDT 的函数调用」→ 生成临时
    // UDT 复合字面量取址 (void*)&(vb6_type_X){...}, 避免强转 struct 值 → C2440.
    // 不能依赖 emitFunctionDecl 逐函数注册 — 类方法无 VB 顺序约束, 调用点
    // (函数体) 可能先于被调函数的声明输出 (cZipArchive pvVfsSetEof 在
    // pvToFileTime 定义之前调用 pvToFileTime(...) 作 SetFileTime 的 As Any 实参).
    funcUdtRetCType_.clear();
    for (auto& decl090q : module.declarations) {
        ASTNode* retTy090q = nullptr;
        std::string fName090q;
        if (decl090q->kind == ASTNodeKind::FunctionDecl) {
            auto& f090q = static_cast<FunctionDecl&>(*decl090q);
            retTy090q = f090q.returnType.get();
            fName090q = f090q.name;
        } else if (decl090q->kind == ASTNodeKind::PropertyDecl) {
            auto& p090q = static_cast<PropertyDecl&>(*decl090q);
            if (p090q.propKind == ProcKind::PropertyGet) {
                retTy090q = p090q.returnType.get();
                fName090q = p090q.name;
            }
        }
        if (!retTy090q || retTy090q->kind != ASTNodeKind::SimpleTypeRef
            || fName090q.empty()) continue;
        std::string retType090q = mapTypeRef(retTy090q);
        if (retType090q.rfind("vb6_type_", 0) == 0) {
            funcUdtRetCType_[Symbol::toLower(fName090q)] = retType090q;
        }
    }
