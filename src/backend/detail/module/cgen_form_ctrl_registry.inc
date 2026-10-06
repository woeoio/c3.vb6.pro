// cgen_form_ctrl_registry.inc - emitFormFramework 函数体片段：控件名映射 / 控件数组检测 与 .h 声明
// 由 src/backend/module/cgen_form.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 145~261 行，逐行未改。

    // --- P7.5+P7.6: 填充已知控件名映射 + 检测控件数组 ---
    {
        std::string formNameLower = formName;
        std::transform(formNameLower.begin(), formNameLower.end(), formNameLower.begin(), ::tolower);
        knownFormName_ = formNameLower;
        knownFormControls_[formNameLower] = FrmControlType::Form;
        knownFormControlOriginalNames_[formNameLower] = formName;

        // P7.6: 先扫描控件数组 (同名控件出现多次 = 数组)
        // Fix 086: 递归注册嵌套控件 (Menu子项如 Window > WindowTopMost).
        // 此前仅注册顶层控件, 嵌套菜单成员访问 (.Checked 等) 误判为
        // Module.X → 生成 vb6_<菜单名>_<属性> 未声明标识符 (C2065).
        std::unordered_set<std::string> ctrlArraySeen;
        std::function<void(const FrmControl&)> registerCtrlRec;
        registerCtrlRec = [&](const FrmControl& ctrl) {
            std::string ctrlNameLower = ctrl.controlName;
            std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
            if (!ctrlNameLower.empty()) {
                const bool duplicateName = !ctrlArraySeen.insert(ctrlNameLower).second;
                if (ctrl.index >= 0 || duplicateName) {
                    // 单元控件 Index=7 也必须生成数组句柄及带 Index 的事件分发.
                    knownControlArrays_[ctrlNameLower] = true;
                }
                knownFormControls_[ctrlNameLower] = ctrl.controlType;
                knownFormControlOriginalNames_[ctrlNameLower] = ctrl.controlName;
                // Fix 112: 工程内 UserControl 子控件 (Begin Proyecto1.ucChartBar ucChartBar1)
                // — 记录 控件名 → .ctl 模块名, 使 `ucChartBar1.AddSerie(...)` 走类成员
                // 直接调用 (vb6_ucChartBar_AddSerie) 而非按 HWND 走 COM 晚绑定.
                if (const FrmFile* ucSpec = findUserControlSpec(ctrl.controlTypeName)) {
                    // 保留 VB_Name 原大小写: 生成 vb6_cls_ucChartBar (与类结构体一致)
                    if (!ucSpec->form.formName.empty()) {
                        knownUserControlCtrlVars_[ctrlNameLower] = ucSpec->form.formName;
                        // 同时登记为类实例变量: 使 IndexOrCallExpr 的可选参数补全
                        // (Fix 090c/Fix 042b) 能按该类形参表展开.
                        knownClassVars_[ctrlNameLower] = ucSpec->form.formName;
                    }
                }
            }
            for (const auto& child : ctrl.children) {
                registerCtrlRec(child);
            }
        };
        for (const auto& ctrl : frmDesc.formControl.children) {
            registerCtrlRec(ctrl);
        }

        // P7.6: 构建控件ID→Index映射 (WM_COMMAND事件分发用)
        {
            int tempId = 100;
            for (const auto& ctrl : frmDesc.formControl.children) {
                std::string ctrlNameLower = ctrl.controlName;
                std::transform(ctrlNameLower.begin(), ctrlNameLower.end(), ctrlNameLower.begin(), ::tolower);
                if (knownControlArrays_.count(ctrlNameLower) && ctrl.index >= 0) {
                    auto& idxMap = controlIdToIndexMap_[ctrlNameLower];
                    int idSlot = tempId - 100;
                    while ((int)idxMap.size() <= idSlot) idxMap.push_back(-1);
                    idxMap[idSlot] = ctrl.index;
                }
                tempId++;
            }
        }
    }

    // --- .h文件: 声明 ---
    h_.emitLine("// P7: Win32 Form - " + formName);
    h_.emitBlank();

    // Fix 010c: 窗体/控件句柄变量移到.c文件, 避免跨窗体.h文件中static变量重定义(C2374)
    // 不同窗体可能有同名控件(如Text1), static变量在.h中被多次包含会导致重定义

    // 窗体句柄变量
    c_.emitLine("static void* vb6_hwnd_" + cIdent(formName) + " = NULL;");
    // Fix 086: 跨模块窗体默认实例访问器 — 其他模块 (cLogs 等) 引用
    // FLogs.Visible / Unload FLogs 时通过 vb6_form_hwnd_<Form>() 获取窗体HWND
    // Fix <vbeclipse>: VB6 的窗体默认实例在**被当对象引用时隐式 Load** (建窗+Form_Load,
    // 不显示)。frmMain 的 `.AddView VIEW_VIEWS, frmViewViews` 与 `frmViewViews.tvwViews`
    // 都靠这条隐式 Load 才拿到有效句柄; 旧访问器只 `return vb6_hwnd_` → 从没 Show 过的
    // 视图窗体恒为 NULL → ucFolder 收到 NULL → 停靠面板全空。这里改成首次引用即 Load。
    // (Unload/Is Nothing 语义偏差可接受: 本工程无 `frmX Is Nothing`, 且 Load+Unload 幂等。)
    h_.emitLine("void* vb6_form_hwnd_" + cIdent(formName) + "(void);");
    h_.emitLine("void vb6_form_load_" + cIdent(formName) + "(void);");
    c_.emitLine("void* vb6_form_hwnd_" + cIdent(formName) + "(void) {");
    c_.emitLine("    if (!vb6_hwnd_" + cIdent(formName) + ") vb6_form_load_" + cIdent(formName) + "();");
    c_.emitLine("    return vb6_hwnd_" + cIdent(formName) + ";");
    c_.emitLine("}");
    // Fix 093a: 跨模块窗体默认实例 Show — 其他模块 `Frm.Show` 转发到本函数
    h_.emitLine("void vb6_form_show_" + cIdent(formName) + "(void* hMDIClient, int modal);");

    // 控件句柄变量 (P7.6: 数组控件使用vb6_CtrlArr, 非数组使用void*)
    // Fix 110: 递归到嵌套容器 (Frame 内的 Picture/TextBox), 并为所有设计期
    // 控件 (含工程内 UserControl 实例 / 第三方 ActiveX / 未知类型) 发射符号 —
    // 此前对 controlTypeToWin32Class()==nullptr 的类型直接 continue, 导致
    // Form2 上的 vb6_hwnd_ucPieChart1 等未声明 (C2065).
    emitControlHandleDecls(frmDesc);

    // Fix <vbeclipse>: 控件跨模块访问器 — 其他模块 `frmViewViews.tvwViews` 引用
    // 控件 HWND (控件句柄变量是本 .c 的 static, 外部直接发裸名必 C2065)。
    // 只给非数组、非 Menu/ImageList 的控件发 (数组/ImageList 的槽形态不同, 有需要再扩)。
    {
        std::function<void(const FrmControl&)> accRec = [&](const FrmControl& ctrl) {
            if (!ctrl.controlName.empty() && ctrl.index < 0
                && ctrl.controlType != FrmControlType::Menu
                && ctrl.controlType != FrmControlType::ImageList) {
                std::string acc = "vb6_form_ctrl_" + cIdent(formName) + "_" + cIdent(ctrl.controlName);
                h_.emitLine("void* " + acc + "(void);");
                c_.emitLine("void* " + acc + "(void) { return vb6_hwnd_" + cIdent(ctrl.controlName) + "; }");
            }
            for (const auto& child : ctrl.children) accRec(child);
        };
        for (const auto& ctrl : frmDesc.formControl.children) accRec(ctrl);
        h_.emitBlank();
    }

    c_.emitLine("static int vb6_formLoading_" + cIdent(formName) + " = 0;");
    // Fix 112: Form_Load 是否已执行 (首次 Form_Resize 必须晚于 Form_Load)
    c_.emitLine("static int vb6_formLoaded_" + cIdent(formName) + " = 0;");

    c_.emitBlank();

    // WndProc前向声明
    h_.emitLine("LRESULT CALLBACK " + wndProc + "(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);");
    // 创建函数前向声明
    h_.emitLine("void " + createFn + "(void* hwnd, void* hInstance);");
    // Show函数前向声明
    h_.emitLine("void " + showFn + "(void* hMDIClient, int modal);");
    h_.emitBlank();
