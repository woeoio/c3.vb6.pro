// cgen_base_generate_decl_pass.inc - generate 函数体片段：第一遍 = 声明
// 由 src/backend/cgen_base.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 498~596 行，逐行未改。

    // === 第一遍: 声明 (前向声明 → .h, 定义 → .c) ===
    // (Enum和Type定义已在类结构体前生成, 见上方Fix 010)

    // 3. Const
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::ConstDecl) {
            visit(static_cast<ConstDecl&>(*decl));
        }
    }

    // 4. 类模块变量: 只注册tracking set, 不生成声明(已在结构体中)
    //    标准模块/窗体模块: 生成变量声明并注册tracking set
    trackOnly_ = isClassModule_;
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::VariableDecl) {
            visit(static_cast<VariableDecl&>(*decl));
        }
    }
    trackOnly_ = false;

    // P13.23: External COM WithEvents need vb6com.h for vb6_CreateEventSink/ComAdvise/ComUnadvise
    // (knownWithEventsVars_ is populated during VariableDecl visits above)
    if (!knownWithEventsVars_.empty()) {
        for (auto& [varLower, srcClassName] : knownWithEventsVars_) {
            auto* srcClsSym = symTab_.lookup(srcClassName);
            if (srcClsSym && srcClsSym->kind == SymbolKind::ComClass && srcClsSym->comHasSourceIface) {
                c_.emitLine("#include \"vb6com.h\"");
                break;
            }
        }
    }

    // === P7: 窗体模块额外代码 ===
    // 窗体模块需要: WndProc声明、控件句柄变量、控件创建函数
    if (module.isFormModule && frmDesc) {
        emitFormFramework(*frmDesc, module);
    } else if (frmDesc) {
        // Fix 110: .ctl/.pag — 只发射设计期子控件符号 (Timer/PictureBox 等).
        // 不发射窗体窗口框架: UserControl/PropertyPage 对外是类, 其可见内容由
        // 代码绘制到 UserControl.hDC; 但 `Timer1.Interval = 100` /
        // `Picture1.ScaleWidth` 这类设计期子控件成员必须能解析.
        emitDesignerControlDecls(*frmDesc, module.docKind);
    }

    // 5. Declare (外部函数声明)
    for (auto& decl : module.declarations) {
        if (decl->kind == ASTNodeKind::DeclareDecl) {
            visit(static_cast<DeclareDecl&>(*decl));
        }
    }

    // 5.4 Delegate (tB 扩展): typedef + 桩前向声明; 定义在 epilogue 发射
    emitDelegateDecls(module);

    // 6. Event声明 (类模块中)
    if (isClassModule_) {
        for (auto& decl : module.declarations) {
            if (decl->kind == ASTNodeKind::EventDecl) {
                visit(static_cast<EventDecl&>(*decl));
            }
        }
    }

    // 7. 过程前向声明 → .h
    // P6.4: 接口类的方法不生成前向声明 (方法是抽象的, 由实现类提供)
    bool currentClassIsInterface = false;
    if (isClassModule_) {
        auto* clsSym = symTab_.lookupModule(module.moduleName);
        if (clsSym && clsSym->kind == SymbolKind::Class && clsSym->isInterface) {
            currentClassIsInterface = true;
        }
    }

    // 6.5 Inherits (tB B07b): 祖先成员的转发桩声明 → .h (接口宿主无继承, 与 currentClassIsInterface 互斥)
    if (isClassModule_ && !currentClassIsInterface) emitInheritedProcDecls(module);
    if (!currentClassIsInterface) emitChainInitDecl(module);  // tB Inherits B09: 构造链入口
    for (auto& decl : module.declarations) {
        // Fix <vbeclipse>: 接口类**要**发方法声明. VB6 里 `Public m_Scheme As New IScheme`
        // (modPublic.bas:12) 会实例化接口本身 —— 接口的 `Public Property Get X()` 声明
        // (IScheme.cls, 无实现体) 就是它的**默认实现**, 调用方 (ucTab/ucButton) 生成的
        // `vb6_IScheme_prop_get_BackColor(m_Scheme)` 需要看得见原型. P6.4 当年"接口抽象
        // 不发声明"的前提是接口**不可实例化**; `As New` 打破该前提 → 之前无声明(靠隐式
        // 声明蒙混过 C 编译) 且 .c 里也无定义 → 链接期 LNK2001 ×26.
        // 泛型模板 (tB): 本体不发前向声明 (特化副本各自以扁名发; 模板名从不被调用)
        if ((decl->kind == ASTNodeKind::SubDecl &&
             !static_cast<SubDecl&>(*decl).typeParams.empty()) ||
            (decl->kind == ASTNodeKind::FunctionDecl &&
             !static_cast<FunctionDecl&>(*decl).typeParams.empty()) ||
            (decl->kind == ASTNodeKind::PropertyDecl &&
             !static_cast<PropertyDecl&>(*decl).typeParams.empty())) {
            continue;
        }
        if (decl->kind == ASTNodeKind::SubDecl) {
            auto& sub = static_cast<SubDecl&>(*decl);
            std::string sig = makeProcSignature(sub);
            // Fix 055: Form事件处理函数不能为static, 因为wndproc用extern引用它们
            bool isFormEventProc = module.isFormModule && sub.name.find("Form_") == 0;
            // Fix 089e: Friend/Public 成员跨模块可调用 → 非 static 声明
            // tB B08a: Protected 也非 static —— 派生模块 TU 的转发桩要能看见它
            if (sub.access == AccessLevel::Public || sub.access == AccessLevel::Friend ||
                sub.access == AccessLevel::Protected || isFormEventProc) {
                h_.emitLine(sig + ";");
            } else {
                h_.emitLine("static " + sig + ";");
            }
        } else if (decl->kind == ASTNodeKind::FunctionDecl) {
            auto& func = static_cast<FunctionDecl&>(*decl);
            std::string sig = makeProcSignature(func);
            // Fix 055: Form事件处理函数不能为static, 因为wndproc用extern引用它们
            bool isFormEventFunc = module.isFormModule && func.name.find("Form_") == 0;
            // Fix 089e: Friend/Public 成员跨模块可调用 → 非 static 声明
            if (func.access == AccessLevel::Public || func.access == AccessLevel::Friend ||
                func.access == AccessLevel::Protected || isFormEventFunc) {  // tB B08a
                h_.emitLine(sig + ";");
            } else {
                h_.emitLine("static " + sig + ";");
            }
        } else if (decl->kind == ASTNodeKind::PropertyDecl) {
            auto& prop = static_cast<PropertyDecl&>(*decl);
            std::string sig = makePropertySignature(prop);
            // Fix 089e: Friend Property (Friend Property Get/Let/Set) 跨模块可调用
            if (prop.access == AccessLevel::Public || prop.access == AccessLevel::Friend ||
                prop.access == AccessLevel::Protected) {  // tB B08a
                h_.emitLine(sig + ";");
            } else {
                h_.emitLine("static " + sig + ";");
            }
        }
    }

    // 7.5 AddressOf 回调桩 (账 #184): 原型跟着过程原型同批发射 (类型依赖同一条顺序),
    //     桩定义后置到 epilogue —— 与委托桩同一套拆分/转发口径。
    emitAddressOfThunkDecls(module);

