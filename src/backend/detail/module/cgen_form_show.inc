// cgen_form_show.inc - emitFormFramework 函数体片段：Show 函数
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1938~1996 行，逐行未改。

    // --- Show函数 (P7.7: 支持MDIForm/MDIChild) ---
    // Fix 146: Owner — 非启动窗体以"当前活动窗体"为 Owner.
    // Owner 为 NULL 时模态窗体无法禁用调用者: 主窗体仍可点击, 于是能在模态
    // 循环打开期间重入打开第二个模态窗体 (嵌套 → 崩溃). VB6 的 Form.Show
    // 语义就是把调用窗体设为 Owner.
    std::string ownerExpr = "NULL";
    if (!startupObject_.empty()) {
        std::string so146 = Symbol::toLower(startupObject_);
        std::string fn146 = Symbol::toLower(formName);
        std::string md146 = Symbol::toLower(moduleName_);
        if (so146 != fn146 && so146 != md146 && so146 != "sub main") {
            ownerExpr = "vb6_Forms_GetActive()";
        }
    }
    c_.emitLine("// === " + formName + " Load (create hidden, fire Form_Load) ===");
    c_.emitLine("void vb6_form_load_" + cIdent(formName) + "(void) {");
    c_.indent();
    c_.emitLine("if (vb6_hwnd_" + cIdent(formName) + ") return;  /* 已 Load */");
    c_.emitLine("void* hInst = vb6_GetAppInstance();");
    c_.emitLine("void* hMDIClient = NULL;");
    if (isMDIForm) {
        c_.emitLine("vb6_RegisterMDIFormClass(\"" + clsName + "\", (void*)" + wndProc + ", hInst, 0);");
        c_.emitLine("vb6_hwnd_" + cIdent(formName) + " = vb6_CreateMDIFormWindow(");
        c_.indent();
        c_.emitLine("\"" + clsName + "\", \"" + caption + "\",");
        if (startupPos == 2) {
            c_.emitLine("-1, -1,  /* M22-Issue2: CenterScreen, RTL handles pixel calc */");
        } else if (startupPos == 3) {
            c_.emitLine("CW_USEDEFAULT, CW_USEDEFAULT,");
        } else {
            c_.emitLine("0, 0,");
        }
        c_.emitLine(std::to_string(clientWidth) + ", " + std::to_string(clientHeight) + ",");
        c_.emitLine("hInst);");
        c_.dedent();
    } else if (isMDIChild) {
        if (formBackColor >= 0) {
            c_.emitLine("vb6_RegisterFormClassBg(\"" + clsName + "\", (void*)" + wndProc + ", hInst, 0, " + std::to_string(formBackColor) + ");");
        } else {
            c_.emitLine("vb6_RegisterFormClass(\"" + clsName + "\", (void*)" + wndProc + ", hInst, 0);");
        }
        c_.emitLine("vb6_hwnd_" + cIdent(formName) + " = vb6_CreateMDIChildWindow(");
        c_.indent();
        c_.emitLine("\"" + clsName + "\", \"" + caption + "\",");
        if (startupPos == 2) {
            c_.emitLine("-1, -1,  /* M22-Issue2: CenterScreen, RTL handles pixel calc */");
        } else if (startupPos == 3) {
            c_.emitLine("CW_USEDEFAULT, CW_USEDEFAULT,");
        } else {
            c_.emitLine("0, 0,");
        }
        c_.emitLine(std::to_string(clientWidth) + ", " + std::to_string(clientHeight) + ",");
        c_.emitLine("hMDIClient, hInst);");
        c_.dedent();
    } else {
        if (formBackColor >= 0) {
            c_.emitLine("vb6_RegisterFormClassBg(\"" + clsName + "\", (void*)" + wndProc + ", hInst, 0, " + std::to_string(formBackColor) + ");");
        } else {
            c_.emitLine("vb6_RegisterFormClass(\"" + clsName + "\", (void*)" + wndProc + ", hInst, 0);");
        }
        c_.emitLine("vb6_hwnd_" + cIdent(formName) + " = vb6_CreateFormWindowB(");
        c_.indent();
        c_.emitLine("\"" + clsName + "\", \"" + caption + "\",");
        if (startupPos == 2) {
            c_.emitLine("-1, -1,  /* M22-Issue2: CenterScreen, RTL handles pixel calc */");
        } else if (startupPos == 3) {
            c_.emitLine("CW_USEDEFAULT, CW_USEDEFAULT,");
        } else {
            c_.emitLine("0, 0,");
        }
        c_.emitLine(std::to_string(clientWidth) + ", " + std::to_string(clientHeight) + ",");
        c_.emitLine("hInst, " + ownerExpr + ", " + std::to_string(formBorderStyle) + ");");
        c_.dedent();
    }
    // Fix <vbeclipse> rev29: 把本窗体的 Form_Resize 入口注册到 HWND 上, 让 RTL
    // 能在**任意时刻直接调用**它 (`vb6_InvokeFormResize`), 不必等 WM_SIZE。
    //
    // 为什么必须这样 (探针实测, play78 --arch x86):
    //   停靠视图窗体 (frmViewViews 等) 由 ucFolder.ViewArea_Resize 用
    //   `vb6_ComCall(视图窗体, L"Move", …)` 摆到最终尺寸 —— 尺寸**确实变了**
    //   (探针: rect 0x0 → 318x291 → 202x364 …), 但它**一生只收到 1 次 WM_SIZE
    //   且那次是建窗时的 0x0** ⇒ rev28 那条 `case VB6_FORM_FR_MSG` 永远等不到
    //   ⇒ `Form_Resize` 跑 0 次 ⇒ `frmViewViews.frm:66`
    //   `tvwViews.Move 0, 0, ScaleWidth, ScaleHeight` 永不执行
    //   (树控件停在设计期 251x211, 比它 188x484 的窗体还宽 = 溢出)。
    //   rev29 试过在 `vb6_ControlMove` 里 PostMessage 排队, 结果**堆损坏 0xC0000374**:
    //   排队那一刻根本**不在主消息循环里** (还在 Form_Load 的同步调用栈中), 消息
    //   进了队列却没人 Dispatch ⇒ 该手法在启动期不成立, 已回退。
    //
    // 所以改成**直接调用**: 存函数指针 → RTL 在 Move 收口处同步调。
    // 同步调为什么安全: `vb6_ControlMove` 此刻已经 `SetWindowPos` 完, 目标窗体的
    // **最终**尺寸已定型, Form_Resize 读到的就是终值 (不是中间态); 而且
    // Form_Resize 内部自己 Move 子控件是**在同一个 hwnd 上**的第二次 Move,
    // 由 `sizeChanged` 判据保证只在真变化时再进一层 ⇒ 收敛。
    {
        std::string resizeFn29 = cProcName("Form_Resize", AccessLevel::Private);
        if (symTab_.lookup(formName + "_Resize") || symTab_.lookup("Form_Resize")) {
            c_.emitLine("{ extern void " + resizeFn29 + "(); vb6_RegisterFormResize(vb6_hwnd_"
                        + cIdent(formName) + ", (void*)" + resizeFn29 + "); }  /* Fix <vbeclipse> rev29 */");
        }
    }
    c_.emitLine("vb6_LoadForm(vb6_hwnd_" + cIdent(formName) + ");");
    c_.dedent();
    c_.emitLine("}");  // Load
    c_.emitBlank();
    c_.emitLine("// === " + formName + " Show ===");
    c_.emitLine("void " + showFn + "(void* hMDIClient, int modal) {");
    c_.indent();
    c_.emitLine("(void)hMDIClient;");
    c_.emitLine("vb6_form_load_" + cIdent(formName) + "();");
    c_.emitLine("vb6_ShowForm(vb6_hwnd_" + cIdent(formName) + ", modal);");
    c_.dedent();
    c_.emitLine("}");  // Show
    c_.emitBlank();