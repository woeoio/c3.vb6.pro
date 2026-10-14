// cgen_expr_member_form_builtin.inc - 函数体片段：Form 的 Me.member 与 IdentifierExpr 首段（App / 窗体默认实例 / 系统对象 / 控件属性）
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 132~362 行，逐行未改。

    // ============================================================
    // P20-44: OLE 拖放的 **DataObject 形参**成员 —— `Data.GetText` / `Data.Files`。
    // 处理器形参在 C 里是 `void** Data` (声明类型 DataObject), **不能掉 COM 派发**
    // (本机 hwnd 不是 IDispatch)。DataObject 指针 = *Data, 所以 RTL 调用首参是
    // `(void*)(*Data)`。
    // ============================================================
    if (node.object && node.object->kind == ASTNodeKind::IdentifierExpr) {
        auto& objIdDD = static_cast<IdentifierExpr&>(*node.object);
        std::string objLowerDD = Symbol::toLower(objIdDD.name);
        auto* symDD = symTab_.lookup(objLowerDD);
        if (symDD && !symDD->variableTypeName.empty()
            && Symbol::toLower(symDD->variableTypeName) == "dataobject") {
            std::string memDD = Symbol::toLower(node.memberName);
            std::string ddArg = "(void*)(*" + objIdDD.name + ")";
            if (memDD == "gettext")      { lastExpr_ = "vb6_oleDD_GetText(" + ddArg + ")"; return; }
            if (memDD == "getfilecount") { lastExpr_ = "vb6_oleDD_GetFileCount(" + ddArg + ")"; return; }
            // Files: 返回文件清单的"伪集合"。Count 走上面那条不行 (Files 本身不是
            // DataObject), 所以单独认 `Data.Files.Count` / `Data.Files(i)`。
            if (memDD == "files") {
                // 记一个 marker: 让外层 .Count / (i) 能接着处理
                lastExpr_ = "vb6_oleDD_GetFileCount(" + ddArg + ")  /* DataObject.Files.Count */";
                return;
            }
        }
    }

    // Fix 023e: Form-module Me.member access — handle Form-specific properties
    // (Height/Width/hwnd/Left/Top/Caption/Visible/Enabled/...) via vb6_GetControlXxx
    // helpers and route unknown members (ScaleHeight/ScaleWidth/Move/Refresh/Show/Cls/...)
    // through COM dispatch on the form HWND (treated as void* IDispatch*).
    // 背景: MeExpr 在 isFormModule_ 时被 visit(MeExpr) 解析为 vb6_hwnd_<FormName>,
    // 这是一个 void* 窗口句柄. Me.Height / Me.ScaleWidth / Me.Move(...) 等成员访问在
    // 下方 IdentifierExpr 分支中无法匹配 (因为 node.object->kind == MeExpr 而非
    // IdentifierExpr), 一直漏到 fallback 处直接 emit "vb6_hwnd_<FormName>.Member"
    // 形成 C2224 (void* 上的 .member 访问). 此处统一处理:
    //   1. 已知 Form 控件属性 (Height/Width/hwnd/Caption/Visible/Enabled/Font*/...)
    //      走 getControlPropReadFn 返回的 vb6_GetControlXxx(vb6_hwnd_<FormName>).
    //   2. 画布家族 (Print / Cls / PSet / Line / Circle / Point) **不走 COM** —— 下面打
    //      标记之后由 cgen_expr_call_callee_withm.inc 的两条码头按 controlCanvasMethod
    //      那张表翻成 RTL 调用 (账 #232②: 这两形以前就落进第 3 档，编得过、一笔不画).
    //   3. 其余未知的 Form 成员 (ScaleHeight/ScaleWidth/Move/Refresh/Show/...)
    //      走 COM dispatch (vb6_ComCall/ComGet*Prop)。
    // 注: form HWND 不是真正的 IDispatch*, 第 3 档是编译 stub (使 MSVC 接受代码); 真正的
    // 运行时正确性需要为 Form 实现专门的 IDispatch 或在 RTL 中添加专用 Form-属性
    // 辅助函数 (vb6_GetFormScaleWidth / vb6_MoveForm 等) — 留待后续 fix.
    if (node.object && node.object->kind == ASTNodeKind::MeExpr && isFormModule_ && !knownFormName_.empty()) {
        // 构造 form 的 HWND C 表达式: vb6_hwnd_<FormName> (保持原始大小写)
        std::string formHwnd;
        auto itOrig = knownFormControlOriginalNames_.find(knownFormName_);
        if (itOrig != knownFormControlOriginalNames_.end()) {
            formHwnd = "vb6_hwnd_" + cIdent(itOrig->second);
        } else {
            formHwnd = "vb6_hwnd_" + moduleName_;
        }
        // 优先用 Form 控件属性读函数解析已知属性
        std::string readFn = getControlPropReadFn(FrmControlType::Form, node.memberName);
        if (!readFn.empty()) {
            lastExpr_ = readFn + "(" + formHwnd + ")  /* Form." + node.memberName + " via Me */";
            return;
        }
        // Form 的绘图状态属性 (CurrentX / CurrentY / DrawWidth) **不在这里答** ——
        // 上面那句 getControlPropReadFn(Form, 名) 就是唯一入口, 读写两侧成对登记在
        // cgen_util_ctrl.cpp 的那两张表里 (账 #233)。这里曾经另摆过四条硬编码:
        // 它只覆盖读侧, 于是赋值语句找不到写侧出口, 发成
        // `vb6_Form_DrawGetWidth(vb6_hwnd_X) = 3;` —— C2106, 两架构都编不过。
        // 未知 Form 成员 → COM dispatch (作为编译 stub; 运行时不可靠)
        comObjExpr_ = formHwnd;
        comMemberName_ = node.memberName;
        isComMarker_ = true;
        isEarlyBoundCom_ = false;
        earlyBoundSym_ = nullptr;
        lastExpr_ = formHwnd;  // void* 表达式 (form HWND)
        return;
    }

    if (node.object && node.object->kind == ASTNodeKind::IdentifierExpr) {
        auto& objIdent = static_cast<IdentifierExpr&>(*node.object);
        std::string objLower = objIdent.name;
        std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
        std::string memLower = node.memberName;
        std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);

        if (objLower == "debug" && memLower == "print") {
            lastExpr_ = "vb6_DebugPrint";
            return;
        }
        if (objLower == "debug" && memLower == "assert") {
            lastExpr_ = "vb6_DebugAssert";
            return;
        }
        // P14.3.4: App全局对象属性
        if (objLower == "app") {
            if (memLower == "path") { lastExpr_ = "vb6_App_Path()"; return; }
            if (memLower == "exename") { lastExpr_ = "vb6_App_EXEName()"; return; }
            if (memLower == "helpfile") { lastExpr_ = "vb6_App_HelpFile()"; return; }  // Fix 090ac: 此前落到 (void*)0.HelpFile → 语法错
            if (memLower == "hinstance") { lastExpr_ = "vb6_App_hInstance()"; return; }
            if (memLower == "hinstancehnd") { lastExpr_ = "vb6_App_hInstance()"; return; }  // VB6别名
            if (memLower == "previnstance") { lastExpr_ = "vb6_App_PrevInstance()"; return; }  // Fix 158k: 此前落到 (void*)0.PrevInstance → C2059
            if (memLower == "title") { lastExpr_ = "vb6_App_EXEName()"; return; }  // 简化
            if (memLower == "major") { lastExpr_ = "0"; return; }
            if (memLower == "minor") { lastExpr_ = "0"; return; }
            if (memLower == "revision") { lastExpr_ = "0"; return; }
            // Fix 056b: App.ThreadID - 当前线程ID (windows.h 已包含)
            if (memLower == "threadid") { lastExpr_ = "(int32_t)GetCurrentThreadId()"; return; }
            // Fix 086: 补齐 App.LogMode / App.LogEvent (ToolsIDE/clsSubClass 引用)
            // Fix 164y: LogMode 必须为 0 —— VB6 语义 0=独立EXE, 1=IDE调试,
            // 2=IDE ActiveX. 生成的产物是独立 EXE, 此前硬编码 1 会让
            // `If App.LogMode <> 0` 为真, 执行 IDE-only 的
            // InitReleaseVisualStylesFixes (comctl32 v6 UAF workaround),
            // 在消息循环中触发 COM vtable 空指针崩溃 (VBFlexGridDemo).
            if (memLower == "logmode") { lastExpr_ = "(int32_t)0"; return; }
            if (memLower == "logevent") {
                // RTL 无 LogEvent — 转发到调试输出 (首参为消息文本, 类型参数忽略)
                lastExpr_ = "((void(*)(BSTR,int32_t))vb6_DebugWriteBSTR)";
                return;
            }
        }

        // Fix 086: 跨模块窗体默认实例属性读取 (FLogs.Visible →
        // vb6_GetControlVisible(vb6_form_hwnd_FLogs())). 窗体名在
        // knownFormModuleNames_ (driver预扫描) 中且非当前窗体时生效.
        if (knownFormModuleNames_.count(objLower)
            && (!isFormModule_ || objLower != knownFormName_)) {
            std::string extFormHwnd = "vb6_form_hwnd_" + cIdent(objIdent.name) + "()";
            // Fix 093a: 窗体默认实例方法 Show — VB6 的 `Frm.Show` 是窗体内置方法,
            // 窗体外的调用此前落到通用的类成员/全局符号解析, 误命中同名用户函数
            // (cLogs: `FLogs.Show` → vb6_FLogs_Show(<cTimeUse.Show 的默认实参>,0)
            // → LNK2019). 转发到窗体自带的 vb6_form_show_<Form>(NULL).
            if (memLower == "show") {
                // 转发到窗体自带的 vb6_form_show_<Form>(hMDIClient, modal).
                // 注意: 此处只交付 callee 对象段 (NULL), 窗口模式实参 (vbModal等)
                // 由上层 com_bind 拆分追加; 不可带尾注释, 否则 callee 不以 ')' 结尾,
                // com_bind 的拆包判定 (callee.back()==')') 会失效并生成
                // "vb6_form_show_x(NULL)(vbModal)" 造成 C2064/C2065.
                lastExpr_ = "vb6_form_show_" + cIdent(objIdent.name) + "(NULL)";
                return;
            }
            std::string extReadFn = getControlPropReadFn(FrmControlType::Form, node.memberName);
            if (!extReadFn.empty()) {
                lastExpr_ = extReadFn + "(" + extFormHwnd + ")  /* external form prop */";
                return;
            }
            // Fix <vbeclipse>: 外部窗体的**控件成员** (`frmViewViews.tvwViews`) →
            // 控件句柄访问器 `vb6_form_ctrl_<Form>_<Ctrl>()` (控件 HWND 是另某 .c
            // 里的 static, 直接发裸名必 C2065)。TreeView 控件再登记进
            // externalTreeViewAccs_, 让后续 `.Nodes` 链 (cgen_util_com) 认领。
            // 不打 COM 标记: 该访问器返回的是 HWND 不是 IDispatch; 后续成员链由
            // cgen_expr_member_generic_access 的 `vb6_form_ctrl_` 前缀识别接管。
            {
                const FrmControl* extCtrl = findExternalFormControl(objIdent.name, node.memberName);
                if (extCtrl) {
                    std::string acc = "vb6_form_ctrl_" + cIdent(objIdent.name)
                                    + "_" + cIdent(node.memberName) + "()";
                    if (extCtrl->controlType == FrmControlType::TreeView)
                        externalTreeViewAccs_[acc] = acc;
                    lastExpr_ = acc;
                    return;
                }
            }
        }

        // P18-C: Clipboard 对象
        if (objLower == "clipboard") {
            if (memLower == "settext") { lastExpr_ = "vb6_Clipboard_SetText"; return; }
            if (memLower == "gettext") { lastExpr_ = "vb6_Clipboard_GetText()"; return; }
            if (memLower == "clear") { lastExpr_ = "vb6_Clipboard_Clear"; return; }
            if (memLower == "getformat") { lastExpr_ = "vb6_Clipboard_GetFormat"; return; }
            // Fix 056: SetData method for Clipboard.SetData picture
            if (memLower == "setdata") { lastExpr_ = "vb6_Clipboard_SetData"; return; }
        }

        // P18-C: Screen 对象
        if (objLower == "screen") {
            if (memLower == "width") { lastExpr_ = "vb6_Screen_Width()"; return; }
            if (memLower == "height") { lastExpr_ = "vb6_Screen_Height()"; return; }
            if (memLower == "mousex") { lastExpr_ = "vb6_Screen_MouseX()"; return; }
            if (memLower == "mousey") { lastExpr_ = "vb6_Screen_MouseY()"; return; }
            if (memLower == "twipsperpixelx") { lastExpr_ = "vb6_Screen_TwipsPerPixelX()"; return; }
            if (memLower == "twipsperpixely") { lastExpr_ = "vb6_Screen_TwipsPerPixelY()"; return; }
            if (memLower == "activecontrol") { lastExpr_ = "(int32_t)(intptr_t)vb6_Screen_ActiveControl()"; return; }
            if (memLower == "activeform") { lastExpr_ = "(int32_t)(intptr_t)vb6_Screen_ActiveForm()"; return; }
            // Fix 161f: Screen.MousePointer 读 (CWaitCursor 等类里用;
            // 早先落进类成员 fallback → me->Screen → C2039)。
            if (memLower == "mousepointer") { lastExpr_ = "vb6_Screen_MousePointer()"; return; }
            // Fix <vbeclipse>: Screen.MouseIcon (读/写) — RTL 全局槽 (IPicture 指针)。
            // 写侧 cgen_util_comwrite 同名分支消费 "vb6_Screen_MouseIcon" 文本。
            if (memLower == "mouseicon") { lastExpr_ = "vb6_Screen_MouseIcon"; return; }
        }

        // P18-C: Printer 对象
        if (objLower == "printer") {
            if (memLower == "print") { lastExpr_ = "vb6_Printer_Print"; return; }
            if (memLower == "enddoc") {
                // 账 #232①-c: 名字问画布表（Printer.Cls 与 Printer.EndDoc 是同一条出口，表里两行都指向它）
                const std::string edFn = canvasDrawEntry("enddoc", CR_CANVAS_PRINTER);
                if (!edFn.empty()) { lastExpr_ = edFn; return; }
            }
            if (memLower == "newpage") { lastExpr_ = "vb6_Printer_NewPage"; return; }
            if (memLower == "width") { lastExpr_ = "vb6_Printer_Width()"; return; }
            if (memLower == "height") { lastExpr_ = "vb6_Printer_Height()"; return; }
            if (memLower == "currentx") { lastExpr_ = "vb6_Printer_CurrentX()"; return; }
            if (memLower == "currenty") { lastExpr_ = "vb6_Printer_CurrentY()"; return; }
            // Task #39: Printer.hDC — ChooseFont/PrintDlg 的 hDC 字段来源
            // (cDlg.cls ShowFont: tChooseFont.hDC = Printer.hDC)
            if (memLower == "hdc") { lastExpr_ = "(void*)vb6_Printer_hDC()"; return; }
            // 2026-10-06: Printer 绘图方法 (PSet/Line/Circle/Point/Cls) + 绘图状态。
            //   此前这些落进下方通用 COM 绑定 ⇒ Printer 不是真 IDispatch, 运行期
            //   静默 no-op, `Printer.Line ...` 在报表里画不出线。
            //   方法**不在这里**翻译 —— 走 comObjExpr_ 标记, 由
            //   cgen_expr_call_callee_withm.inc 的绘图段统一发射 (与 Form 同款,
            //   实参个数/可选形参的判定逻辑只写一份)。
            if (memLower == "pset" || memLower == "line" || memLower == "circle"
                || memLower == "point" || memLower == "cls") {
                comObjExpr_ = "Printer";
                comMemberName_ = node.memberName;
                isComMarker_ = true;
                isEarlyBoundCom_ = false;
                earlyBoundSym_ = nullptr;
                lastExpr_ = "Printer";
                return;
            }
            // 绘图状态读 (CurrentX/CurrentY 上面已给, 这里补 ForeColor/DrawWidth)
            if (memLower == "forecolor") { lastExpr_ = "vb6_Printer_DrawGetForeColor()"; return; }
            if (memLower == "drawwidth") { lastExpr_ = "vb6_Printer_DrawGetWidth()"; return; }
        }

        // P18-C: Forms 集合
        if (objLower == "forms") {
            if (memLower == "count") { lastExpr_ = "vb6_Forms_Count()"; return; }
            if (memLower == "item") { lastExpr_ = "(int32_t)(intptr_t)vb6_Forms_Item"; return; }
        }

        // Fix 161: Console 对象 (对齐 twinBASIC) —— 控制台 I/O 的现代口径。
        //   Console.WriteLine([text]) / Console.Write(text)  输出 (走 vb6_ConWriteOutW,
        //    控制台 WriteConsoleW / 管道按代码页, 中文不乱)
        //   Console.ReadLine() As String   读一行
        //   Console.ReadKey()  As String   读一个键 (取行首字符)
        //   Console.Read()     As String   同 ReadKey (twinBASIC 无换行等待的简化口径)
        // stdin 侧由 vb6_ConReadLine 按 GetConsoleMode 分流 (控制台 ReadConsoleW /
        // 管道 ReadFile+代码页), 修掉此前裸 fgetws 把管道字节当 UTF-16 的必然乱码。
        if (objLower == "console") {
            if (memLower == "writeline")  { lastExpr_ = "vb6_Console_WriteLine"; return; }
            if (memLower == "write")      { lastExpr_ = "vb6_Console_Write"; return; }
            if (memLower == "readline")   { lastExpr_ = "vb6_Console_ReadLine()"; return; }
            if (memLower == "readkey")    { lastExpr_ = "vb6_Console_ReadKey()"; return; }
            if (memLower == "read")       { lastExpr_ = "vb6_Console_ReadKey()"; return; }
        }

        // P7.5+P7.6: 窗体控件属性读取
        // 情况1: ctrl.Property (非数组) → vb6_GetControlXxx(vb6_hwnd_ctrl)
        // 情况2: ctrlArr(idx).Property (数组) → vb6_GetControlXxx(vb6_CtrlArr_GetAt(&vb6_arr_ctrl, idx))
        {
            // P16: WithEvents控件属性读取 (优先于标准控件, 因为变量名可能同名)
            auto itWECtrl = knownWithEventsCtrlVars_.find(objLower);
            if (itWECtrl != knownWithEventsCtrlVars_.end()) {
                std::string readFn = getControlPropReadFn(itWECtrl->second, node.memberName);
                if (!readFn.empty()) {
                    auto itOrig = knownWithEventsCtrlOrigNames_.find(objLower);
                    std::string weVarName = (itOrig != knownWithEventsCtrlOrigNames_.end()) ? itOrig->second : objLower;
                    lastExpr_ = readFn + "(" + weVarName + ")  /* WithEvents ctrl */";
                    return;
                }
            }
            // Fix 161f-extlist: `lv.hWnd` / `lv.View` —— lv 是**形参/局部持有的
            // ListView** (`Sub RefillList(lv As ListView)`)。C 侧槽变量本身就是 HWND,
            // 与设计期控件名 (knownFormControls_ + vb6_hwnd_<名>) 拼法不同:
            // 这里直接用裸槽名当 hwnd 实参, 不发 vb6_hwnd_ 前缀。
            // 不认这条就落 COM 晚绑定 `vb6_ComGetProp((*lv), L"hWnd")` —— 拿
            // vb6_ComIface_IListView* 当 IDispatch 用 (extlist CStopWatch/LVSortK 实测:
            // hWnd 读到 Empty ⇒ 后续 SendMessage 全打 0)。
            if (listViewSlotVars_.count(objLower)) {
                std::string readFnLv = getControlPropReadFn(FrmControlType::ListView, node.memberName);
                if (!readFnLv.empty()) {
                    lastExpr_ = readFnLv + "((void*)" + objLower + ")  /* ListView slot param */";
                    return;
                }
            }
            // P7.5: 非数组控件属性读取
            auto itCtrl = knownFormControls_.find(objLower);
            if (itCtrl != knownFormControls_.end()) {
                // Fix 112: 工程内 UserControl 子控件 → 类成员直接调用.
                // `ucChartBar1.AddSerie(Value)` 此前落 COM 晚绑定
                // `vb6_ComCall(vb6_hwnd_ucChartBar1, L"AddSerie", ...)`, 而 HWND 不是
                // IDispatch. 这里按 .ctl 模块解析成员, this 取宿主窗口的实例.
                auto itUC = knownUserControlCtrlVars_.find(objLower);
                if (itUC != knownUserControlCtrlVars_.end()) {
                    // 账 #191: 这段整体抽到 CCodeGen::emitUcInstanceMemberExpr —— 控件数组的元素
                    // (`ucChartBar1(i).AddSerie ...`) 要走同一个出口，不能各抄一份。
                    if (emitUcInstanceMemberExpr(itUC->second, node.memberName,
                                                makeCtrlHwndArg(objLower, itCtrl->second))) {
                        return;
                    }
                }
                // 账 #189: 数组整体成员先判 (Count 以前不在这里 ⇒ 掉进 COM 兜底拿
                // vb6_hwnd_<数组名>，C2065)。这三条**永远不是**控件属性, 所以放在 readFn 之前。
                if (knownControlArrays_.count(objLower)) {
                    std::string metaExpr = ctrlArrayMetaMemberExpr(objIdent.name, memLower);
                    if (!metaExpr.empty()) {
                        lastExpr_ = metaExpr;
                        return;
                    }
                }
                std::string readFn = getControlPropReadFn(itCtrl->second, node.memberName);
                if (!readFn.empty()) {
                    lastExpr_ = readFn + "(" + makeCtrlHwndArg(objLower, itCtrl->second) + ")  /* ctrl prop read */";
                    return;
                }
                                // P7.9: WebBrowser method access (Navigate/GoBack/GoForward/Refresh)
                if (itCtrl->second == FrmControlType::WebBrowser) {
                    std::string memLower = node.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                    if (memLower == "navigate" || memLower == "goback" || memLower == "goforward" || memLower == "refresh") {
                        // Set marker for IndexOrCallExpr to handle
                        comObjExpr_ = objLower;  // Store lowercase control name
                        comMemberName_ = memLower;
                        isComMarker_ = true;
                        isEarlyBoundCom_ = false;
                        earlyBoundSym_ = nullptr;
                        lastExpr_ = objLower;  // Placeholder expression
                        return;
                    }
                }
                // P13.3: ListBox/ComboBox method access (AddItem/RemoveItem/Clear/List)
                // C29-1b: 文件系统三控件本质就是这两类窗口 (Drive=组合框、Dir/File=列表框),
                // 可索引属性 `Dir1.List(i)` 与 `File1.Clear` 之类走完全相同的一条路。
                if (itCtrl->second == FrmControlType::ListBox || itCtrl->second == FrmControlType::ComboBox
                    || itCtrl->second == FrmControlType::DirListBox
                    || itCtrl->second == FrmControlType::FileListBox
                    || itCtrl->second == FrmControlType::DriveListBox) {
                    std::string memLower = node.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                    if (memLower == "additem" || memLower == "removeitem" || memLower == "clear" || memLower == "list") {
                        comObjExpr_ = objLower;
                        comMemberName_ = memLower;
                        isComMarker_ = true;
                        isEarlyBoundCom_ = false;
                        earlyBoundSym_ = nullptr;
                        lastExpr_ = objLower;
                        return;
                    }
                }
                // D6 / C29-9: CommonDialog 的六个 Show* —— 打标给
                // cgen_expr_call_callee_withm.inc 消费（形状同 PictureBox.Print/Cls 那一刀：
                // comObjExpr_ 存**小写控件名**，消费侧查 knownFormControls_ 拿类型后发
                // vb6_CdShow*(hwnd)）。不这么做就会落到下面 Fix 023e/089d 那条通用兜底，
                // 编成运行期 no-op 的 vb6_ComCall(hwnd, L"ShowOpen") —— 正是本批要拆的东西。
                if (itCtrl->second == FrmControlType::CommonDialog) {
                    std::string memLower = node.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                    if (memLower == "showopen" || memLower == "showsave" || memLower == "showcolor"
                        || memLower == "showfont" || memLower == "showprinter" || memLower == "showabout") {
                        comObjExpr_ = objLower;
                        comMemberName_ = memLower;
                        isComMarker_ = true;
                        isEarlyBoundCom_ = false;
                        earlyBoundSym_ = nullptr;
                        lastExpr_ = objLower;
                        return;
                    }
                }
                // C29-RT-c: RichTextBox 的 LoadFile / SaveFile / Find —— 打标给
                // cgen_expr_call_callee_withm.inc 消费（形状同 CommonDialog.Show* 那一刀：
                // comObjExpr_ 存**小写控件名**，消费侧查 knownFormControls_ 拿类型后发 vb6_RTB_*）。
                // 少了这一刀就落到下面的通用兜底 ⇒ vb6_ComCall(vb6_hwnd_rt1, L"Find", ...) 一声不响。
                // C29-WS-a/b: Winsock 的方法（Bind / SendData / GetData / PeekData / Close
                // + TCP 那三条 Listen / Accept / Connect）同一族打标。WS-a 那格刻意不发 TCP
                // 三条 —— 那时它们是空函数，发出去就是本线反复踩的"编得过、运行期一声不响"。
                if (itCtrl->second == FrmControlType::Winsock) {
                    std::string wsMemLower = node.memberName;
                    std::transform(wsMemLower.begin(), wsMemLower.end(), wsMemLower.begin(), ::tolower);
                    if (wsMemLower == "bind" || wsMemLower == "senddata"
                        || wsMemLower == "getdata" || wsMemLower == "peekdata"
                        || wsMemLower == "close" || wsMemLower == "listen"
                        || wsMemLower == "connect" || wsMemLower == "accept") {
                        comObjExpr_ = objLower;
                        comMemberName_ = wsMemLower;
                        isComMarker_ = true;
                        isEarlyBoundCom_ = false;
                        earlyBoundSym_ = nullptr;
                        lastExpr_ = objLower;
                        return;
                    }
                }
                if (itCtrl->second == FrmControlType::RichTextBox) {
                    std::string memLower = node.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                    if (memLower == "loadfile" || memLower == "savefile" || memLower == "find"
                        || memLower == "simnotify") {
                        comObjExpr_ = objLower;
                        comMemberName_ = memLower;
                        isComMarker_ = true;
                        isEarlyBoundCom_ = false;
                        earlyBoundSym_ = nullptr;
                        lastExpr_ = objLower;
                        return;
                    }
                }
                // Fix 185 / 账 #232②: 画布家族（Print / Cls / Line）是**绘制**方法, 不是 COM 成员。
                // 必须落在下面 Fix 023e/089d 兜底**之前**: 兜底把 comObjExpr_ 设成
                // HWND 表达式(:296), 调用侧的裸名码头就拿不到控件类型了 —— 今天
                // "Picture2.Print \"Drag/drop me\"" 编成
                // vb6_ComCall(vb6_hwnd_Picture2, L"Print", …), 而控件 HWND 不是真的
                // IDispatch*, 运行期静默 no-op, 那个框一直是空白。
                // 账 #221 起名字出自 controlCanvasMethod 这一张表（cls / print / line 三档），
                // 这里不再各自列名字 —— 表加一档，两条码头与这一处打标记同步长。
                if (itCtrl->second == FrmControlType::PictureBox
                    && !controlCanvasMethod(itCtrl->second, memLower).empty()) {
                    comObjExpr_ = objLower;  // 小写控件名, 与 WebBrowser/ListBox 块同构
                    comMemberName_ = memLower;
                    isComMarker_ = true;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                    lastExpr_ = objLower;
                    return;
                }
                // ActiveX控件 (ImageList等): 属性访问走COM后期绑定
                // 合并(P20-40 + C29-9): **无窗口控件只剩 ImageList / Toolbar**。
                // StatusBar (msctls_status32) 与 CommonDialog (VB6_COMMONDIALOG 属性宿主)
                // 都是真窗口 ⇒ 走下面通用控件分支拿 vb6_hwnd_X; 归到 vb6_com_X 家族会生成
                // vb6_ComGetObjectProp(vb6_com_X, …) 而那个变量并不存在 (C2065)。
                // C29-5a: Toolbar 摘出 —— 它有真窗口了，属性读写走 cgen_util_ctrl 那张表，
                // 未登记的成员才落到下面那条通用 HWND 分支。
                if (itCtrl->second == FrmControlType::ImageList) {
                    std::string memLower = node.memberName;
                    std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
                    std::string ctrlOrigName = knownFormControlOriginalNames_.count(objLower) ?
                        knownFormControlOriginalNames_[objLower] : objIdent.name;
                    comObjExpr_ = "vb6_com_" + cIdent(ctrlOrigName);
                    comMemberName_ = node.memberName;
                    isComMarker_ = true;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                    lastExpr_ = comObjExpr_;  // IDispatch* expression
                    return;
                }
                // Fix 023e/089d: 控件 readFn 为空 (未知属性/方法如 Form.ScaleWidth,
                // Move/Refresh/Align等) 走 COM dispatch on control HWND, 避免落入
                // warn-and-fall-through 后由通用 fallback 先做默认属性展开再拼
                // ".Member" → C2224 (vb6_GetControlText(...) .Move / .Align).
                // Fix 089d 将 Form 分支扩展为所有非 Menu 控件 (FLogs List1.Move、
                // FLayer LContent.Move、FToastDrawer Picture1.Align 等).
                // 注: 与上面 MeExpr 分支一致, 这是编译 stub; 控件 HWND 不是真正的
                // IDispatch*, 运行期语义由真实 VB6 runtime 提供.
                if (itCtrl->second != FrmControlType::Menu) {
                    std::string ctrlHwnd = makeCtrlHwndArg(objLower, itCtrl->second);
                    comObjExpr_ = ctrlHwnd;
                    comMemberName_ = node.memberName;
                    isComMarker_ = true;
                    isEarlyBoundCom_ = false;
                    earlyBoundSym_ = nullptr;
                    lastExpr_ = ctrlHwnd;  // void* 表达式 (control HWND)
                    return;
                }
                diag_.warn(DiagnosticID::CodeGenUnsupportedFeature, SourceLocation{},
                    std::string("P7.5: Unknown control property '") + objIdent.name + "." + node.memberName +
                    "' for control type, generating struct field access (may not compile)");
            }
        }
