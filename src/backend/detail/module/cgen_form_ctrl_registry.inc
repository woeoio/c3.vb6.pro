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
                knownFormControlTypeNames_[ctrlNameLower] = ctrl.controlTypeName;
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
                        moduleClassVars_[ctrlNameLower] = ucSpec->form.formName;
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

        // 设计器 OCX 控件的事件连接点。VB6 里 `<Control>_<Event>` 是控件**事件源接口**
        // 的接收器 (控件创建后 Advise 连接点), 而不是窗体自己的事件: 从前只有
        // `Dim WithEvents x As ...` 会登记 knownWithEventsVars_, 设计期放的 OCX 控件
        // 没人登记 —— 于是 `CommandBars_Execute` 这类处理器成了「有函数没人调」,
        // FindControl(...).Execute 的 COM 调用虽然成功, 事件却永远到不了 VB 代码
        // (Outlook2013: 主题切换没执行 ⇒ ribbon/状态栏一片黑)。
        //
        // 登记条件: 控件类型解析到 ComClass + 该 ComClass 有事件源接口 + 本模块
        // 真写了至少一个 `<控件名>_<事件名>` 处理器。登记后 evt_decl/evt_impl 照常
        // 发包装器, 控件创建点 (cgen_form_create_controls.inc) 照常 Advise。
        {
            std::function<void(const FrmControl&)> regOcxEvtRec = [&](const FrmControl& ctrl) {
                if (!ctrl.controlName.empty()) {
                    Symbol* cs = lookupDotted(ctrl.controlTypeName);
                    if (!cs || cs->kind != SymbolKind::ComClass) {
                        size_t dot = ctrl.controlTypeName.find('.');
                        if (dot != std::string::npos)
                            cs = symTab_.lookup(ctrl.controlTypeName.substr(dot + 1));
                    }
                    if (cs && cs->kind == SymbolKind::ComClass && cs->comHasSourceIface
                        && !cs->comSourceIfaceIid.empty() && !cs->eventNames.empty()) {
                        // 处理器必须**本模块自有**: 裸 lookup 会命中别家窗体同名的
                        // Public Sub (pageBackstageInfo.frm:148 也有 CommandBars 控件,
                        // 而 `Public Sub CommandBars_Execute` 只写在 frmMain:405) ——
                        // 那样会在本模块发一个指向 `vb6_pageBackstageInfo_CommandBars_Execute`
                        // 的包装器, LNK2019。判据同 cgen_form_show.inc 的 hasOwnResize29。
                        bool anyHandler = false;
                        for (const auto& evtName : cs->eventNames) {
                            Symbol* hSym = symTab_.lookupModule(ctrl.controlName + "_" + evtName);
                            if (hSym && !hSym->isExternal) {
                                anyHandler = true;
                                break;
                            }
                        }
                        if (anyHandler) {
                            knownWithEventsVars_[Symbol::toLower(ctrl.controlName)] = cs->name;
                        }
                    }
                }
                for (const auto& child : ctrl.children) regOcxEvtRec(child);
            };
            for (const auto& ctrl : frmDesc.formControl.children) regOcxEvtRec(ctrl);
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
    // 非数组、非 Menu 的控件都发跨模块访问器。ImageList 的槽是 vb6_com_<Name>
    // (IDispatch*, 见 emitControlHandleDecls:165), 不是 vb6_hwnd_<Name>, 下方分支处理。
    {
        // 控件数组只发一枚带下标的访问器 (递归会逐元素 visit, 按名去重) —
        // 跨窗体 `frmX.btnTab(0).Appearance` 引用 vb6_form_ctrl_<F>_<C>(idx)
        // (Outlook2013 frmMain:327-330, 数组访问器缺失 = LNK2019)。
        std::unordered_set<std::string> accEmitted;
        std::function<void(const FrmControl&)> accRec = [&](const FrmControl& ctrl) {
            if (!ctrl.controlName.empty() && ctrl.controlType != FrmControlType::Menu) {
                std::string acc = "vb6_form_ctrl_" + cIdent(formName) + "_" + cIdent(ctrl.controlName);
                // 兼容别名 `vb6_<Form>_<Ctrl>()`: 类成员调用机制 (Fix 015/088c 把
                // 窗体当类、控件当成员) 与若干回退会按这个形状发码 —— 直接提供
                // 定义, 不再追每一处发射点 (Outlook2013 frmMain.CommandBars 实测
                // LNK2019, 发射点在三套链式机制深处)。
                std::string alias = "vb6_" + cIdent(formName) + "_" + cIdent(ctrl.controlName);
                if (ctrl.index < 0) {
                    h_.emitLine("void* " + acc + "(void);");
                    h_.emitLine("void* " + alias + "(void);");
                    // 族4b: ImageList 槽为 vb6_com_<Name> (IDispatch*); 其余控件回 vb6_hwnd_<Name>。
                    // 旧逻辑把 ImageList 排除, 但 comwrite 侧又发了 vb6_form_ctrl_<Form>_ImageListN()
                    // 的调用 -> 定义缺失 (LNK2019)。这里补上 ImageList 分支。
                    if (ctrl.controlType == FrmControlType::ImageList) {
                        c_.emitLine("void* " + acc + "(void) { return vb6_com_" + cIdent(ctrl.controlName) + "; }");
                    } else {
                        c_.emitLine("void* " + acc + "(void) { return vb6_hwnd_" + cIdent(ctrl.controlName) + "; }");
                    }
                    c_.emitLine("void* " + alias + "(void) { return " + acc + "(); }");
                } else if (!accEmitted.count(ctrl.controlName)) {
                    accEmitted.insert(ctrl.controlName);
                    h_.emitLine("void* " + acc + "(int32_t idx);");
                    h_.emitLine("void* " + alias + "(int32_t idx);");
                    // 控件数组载体是 vb6_CtrlArr (hwnds[256] + count/lowerBound/upperBound),
                    // 不是 vb6_SafeArray1D*。用 VB6_SA_AT 会把裸 hwnds 数组当安全数组头解引用
                    // ⇒ 边界读成 lb=0 ub=0 ⇒ `pageBackstageSend.btnTab(1)` 直接
                    // "Subscript out of range" (Outlook2013 frmMain:328)。
                    c_.emitLine("void* " + acc + "(int32_t idx) { return vb6_CtrlArr_GetAt(&vb6_arr_"
                              + cIdent(ctrl.controlName) + ", idx); }");
                    c_.emitLine("void* " + alias + "(int32_t idx) { return " + acc + "(idx); }");
                }
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
