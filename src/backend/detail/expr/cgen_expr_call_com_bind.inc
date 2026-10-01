// cgen_expr_call_com_bind.inc - visit(IndexOrCallExpr&) 函数体片段：COM 绑定检测（P6.2/P6.3/P6.4）与 callee 拆分判定
// 由 src/backend/expr/cgen_expr_call.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 1168~1553 行，逐行未改。

    // --- COM后期绑定检测 (P6.2) + 前期绑定检测 (P6.3) + P6.4接口调用 ---
    // MemberAccessExpr为COM对象设置isComMarker_标志 + comObjExpr_/comMemberName_
    if (isComMarker_) {
        // 2026-09-30 (门 #233 c29sbevt): **控件的判据专用方法先落地**。
        // 位置: 必须在下面**所有**分支 (Sim 汇总那一块 / axSlotObj / Data / OLE) **之前** ——
        // 这些分支的顺序改过好几轮, 而 `StatusBar1.SimClick 1, 0` 一旦落进 axSlotObj 那条
        // "裸控件槽 + 取属性 + Item" 的路, 就编成
        //   vb6_ComCallObject(vb6_ComGetObjectProp(hwnd, L"SimClick"), L"Item", …)
        // 编得过、跑起来一条事件都不发 (CI 上 `c29sbevt` / `_x86` 的 `Got:` 只有 SB1-COUNT=3)。
        // 判据只用**登记表** (`statusBarNameOfExpr`, 它认宿主槽本身与两种集合形), 不靠前缀猜。
        {
            std::string sbHost = statusBarNameOfExpr(comObjExpr_);
            if (!sbHost.empty() && Symbol::toLower(comMemberName_) == "simclick") {
                isComMarker_ = false;
                std::vector<std::string> sa;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    sa.push_back(lastExpr_);
                }
                lastExpr_ = "vb6_StatusBar_SimClick((void*)vb6_hwnd_" + sbHost + ", "
                          + (sa.size() > 0 ? sa[0] : std::string("0")) + ", "
                          + (sa.size() > 1 ? sa[1] : std::string("0")) + ")";
                comObjExpr_.clear();
                comMemberName_.clear();
                return;
            }
        }
        // ActiveX控件COM属性: 先获取属性对象, 再用Item(idx)索引
        // comObjExpr_以"vb6_com_"开头 = ActiveX控件变量
        // C29-7: ListView 是**真窗口** (`vb6_hwnd_X`), 而 ImageList 是 `vb6_com_X` ——
        // 这条"控件属性对象 + Item(idx)"的路两种槽都要认, 否则 `ListView1.ListItems(1)`
        // 会被拼成 `vb6_ComCallObject(vb6_hwnd_ListView1, L"ListItems", …)` —— 对 HWND
        // 当 IDispatch 用, 链接过、运行期读数全空。
        // C29-8c: TreeView1.SimNodeClick n / .SimExpand n, expanded —— **判据专用**方法
        // (与 C29-4 的 StatusBar.SimClick 同一先例: 无头环境点不了鼠标，RTL 程序化发一条
        // 真 WM_NOTIFY，整条"派发分支 + 通知换算 + NodeAt 造对象"才算被验到)。不对应任何
        // VB6 语义，写进 ai/029 边界。
        // ⚠ 必须抢在下面那条 axSlotObj 分支**之前**: 那条现在是"裸控件槽 + 取属性 + Item"
        // 的形状 (C29-8b 把 TreeView/Toolbar 的裸槽放了进去)，而 `tv1.SimNodeClick 2` 长得
        // 一模一样 —— 不抢先就被吞成 ComCallObject(getprop "SimNodeClick", Item(2)):
        // 编得过、跑起来什么都不发 (实测踩过)。这里也不能用后面那两个 objExpr/memberName，
        // 此刻还在 marker 形态上，读 comObjExpr_/comMemberName_。
        if (comObjExpr_.find("vb6_hwnd_") == 0) {
            std::string tvBare = comObjExpr_.substr(9);
            std::string tvMem = Symbol::toLower(comMemberName_);
            bool tvSim = treeViewVars_.count(tvBare) > 0
                && (tvMem == "simnodeclick" || tvMem == "simexpand");
            // C29-5c: 工具栏那两条判据方法同形 (SimButtonClick → WM_COMMAND,
            // SimButtonMenuClick → WM_NOTIFY/TBN_DROPDOWN，两条都发给父窗)。
            bool tbSim = !toolbarNameOfExpr(comObjExpr_).empty()
                && (tvMem == "simbuttonclick" || tvMem == "simbuttonmenuclick");
            // C29-DT-c: DTPicker 的三条同形 (SimChange → DTN_DATETIMECHANGE、
            // SimDropDown → DTN_DROPDOWN、SimCloseUp → DTN_CLOSEUP，三条都发给父窗)。
            // 全部**无参** —— VB6 那三个处理器没有参数，所以不需要下面那两个 a0/a1。
            bool dtSim = dtpickerVars_.count(tvBare) > 0
                && (tvMem == "simchange" || tvMem == "simdropdown" || tvMem == "simcloseup");
            // C29-MV-c: MonthView 的判据方法同形 (SimDateClick d → MCN_SELCHANGE 发给父窗)。
            // 它**带一个 Date 参数**（DT-c 那三条无参，所以那里 a0 用不上）。
            bool mvSim = monthviewVars_.count(tvBare) > 0 && tvMem == "simdateclick";
            // C29-SL-c: Slider 的判据方法同形 (SimNotify code, pos → 原生那一档的
            // WM_HSCROLL/WM_VSCROLL 发给父窗)，**带两个整型参数**。
            // C29-SL-d: SimStdEvent kind, wParam 同形（把一条常规事件的消息放进控件自己的
            // 队列），形参表与 SimNotify 一样是两个整型，所以共用这一支。
            bool slSim = !sliderNameOfExpr(comObjExpr_).empty()
                && (tvMem == "simnotify" || tvMem == "simstdevent");
            bool slStd = slSim && tvMem == "simstdevent";
            // C29-SL-i: ClearSel 是 **VB6 真方法**（类型库 ISlider dispid 0x000e，无参），
            // 不是判据助手。走同一支改道是因为它同样会被 axSlotObj 那支先吃掉 —— 落回兜底就编成
            // 「取 clearsel 属性 + 丢掉结果」，跑起来一声不响（账 #143 记的就是这一族的语句形）。
            // C29-SL-l: 这一族扩成一张表（`controlZeroArgMethod`）—— **带括号**在这一形收尾、
            // **不带括号**在 cgen_call.cpp 的语句路收尾，两条码头共用同一张表。只接一头就是
            // 账 #143 那条"编得过、跑了、什么都没发生"（本线踩过三次的那同一条）。
            std::string zaHwnd;
            FrmControlType zaType = FrmControlType::Unknown;
            std::string zaFn = formCtrlSlot(comObjExpr_, zaType, zaHwnd)
                ? controlZeroArgMethod(zaType, tvMem) : std::string();
            bool slClear = !zaFn.empty();
            if (tvSim || tbSim || dtSim || mvSim || slSim || slClear) {
                isComMarker_ = false;
                std::vector<std::string> sa;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    sa.push_back(lastExpr_);
                }
                std::string a0 = sa.size() > 0 ? sa[0] : std::string("0");
                std::string a1 = sa.size() > 1 ? sa[1] : std::string("0");
                lastExpr_ = slClear
                    ? (zaFn + "((void*)" + zaHwnd + ")")
                    : slStd
                    ? ("vb6_Slider_SimStdEvent((void*)" + comObjExpr_ + ", " + a0 + ", " + a1 + ")")
                    : slSim
                    ? ("vb6_Slider_SimNotify((void*)" + comObjExpr_ + ", " + a0 + ", " + a1 + ")")
                    : mvSim
                    ? ("vb6_MV_SimDateClick((void*)" + comObjExpr_ + ", " + a0 + ")")
                    : dtSim
                    ? (std::string(tvMem == "simchange"   ? "vb6_DTP_SimChange"
                           : tvMem == "simdropdown"       ? "vb6_DTP_SimDropDown"
                                                          : "vb6_DTP_SimCloseUp")
                       + "((void*)" + comObjExpr_ + ")")
                    : (tvMem == "simnodeclick")
                    ? ("vb6_TreeView_SimNodeClick((void*)" + comObjExpr_ + ", " + a0 + ")")
                    : (tvMem == "simexpand")
                    ? ("vb6_TreeView_SimExpand((void*)" + comObjExpr_ + ", " + a0 + ", " + a1 + ")")
                    : (tvMem == "simbuttonclick")
                    ? ("vb6_Toolbar_SimButtonClick((void*)" + comObjExpr_ + ", " + a0 + ")")
                    : ("vb6_Toolbar_SimButtonMenuClick((void*)" + comObjExpr_ + ", " + a0 + ")");
                comObjExpr_.clear();
                comMemberName_.clear();
                return;
            }
        }

        // C29-8b/C29-5b: TreeView / Toolbar 的**裸控件**槽也进这条"取集合属性 + Item"
        // 的路; 集合对象本身 (vb6_TreeView_Nodes(…) 开头) 不能进, 否则
        // `Nodes.Add(…)` 会被当成"Add 属性 + Item 下标", 编成删不改的错形状。
        // Fix 161f: `lv.ListItems(1)` —— 形参/局部槽 (lv As ListView) 本身就是
        // HWND, 不是 vb6_hwnd_ 前缀形态, 必须单独认, 否则落
        // vb6_ComGetObjectProp(lv, L"ListItems") = 拿 HWND 当 IDispatch 用
        // (extlist RefillList 实测: 窗口起来但 0 行)。
        // <vbeclipse>: 上面那句"集合对象本身不能进"过去由 `vb6_hwnd_` 那个**前缀锚**兜住
        // (dev 的写法是 `find("vb6_hwnd_")==0 && (listView||treeView||toolbar 命中)`)。
        // 这一支持改成 isListViewHostExpr / *NameOfExpr 这些**子串**判据后，锚就失效了：
        // vb6_TreeView_Nodes((void*)vb6_hwnd_tv1) 里嵌着 vb6_hwnd_tv1，子串一命中就被当成
        // 宿主槽，`Nodes.Add(…)` 于是编成"取 Add 属性 + Item 下标"——节点根本没加进去，
        // Index 全读空 (实测 12 条控件回读红: ctrltreeview / ctrltoolbar / c29listview /
        // c29sbevt 各 x64+x86)。现在由 isControlCollectionExpr 显式挡 (一处有名字的话)。
        bool axSlotObj = !isControlCollectionExpr(comObjExpr_)
                      && (comObjExpr_.find("vb6_com_") == 0
                          || comObjExpr_.find("vb6_hwnd_") == 0
                          || isListViewHostExpr(comObjExpr_)
                          || !treeViewNameOfExpr(comObjExpr_).empty()
                          || !toolbarNameOfExpr(comObjExpr_).empty()
                          // C29-Data: Data1.Recordset 的透传形态 (Fields/Item 走 Step1+Item 路)。
                          || comObjExpr_.find("vb6_Data_Self(") == 0);
        // Fix <VBFlexGridDemo>: 窗体 (Me) 的 ScaleX/ScaleY(x, fromScale, toScale) 是
        // **方法**, 不是"集合属性 + Item 下标"。axSlotObj 只按 vb6_hwnd_ 前缀认宿主槽,
        // 窗体 hwnd 同样命中 ⇒ Me.ScaleX(8, 3, Me.ScaleMode) 被编成
        //   vb6_ComCallObject(vb6_ComGetObjectProp(vb6_hwnd_MainForm, L"ScaleX"), L"Item", …)
        // 返回 void* ⇒ 与 float 相减 ⇒ C2113×2 (MainForm.c 640/641)。
        // 让它们落到下面的通用 COM 方法调用 (BASE 8c3e981 正是
        // vb6_ComCall(vb6_hwnd_MainForm, L"ScaleX", …))。ScaleX/ScaleY 不可能是集合。
        const std::string axMemLower = Symbol::toLower(comMemberName_);
        const bool axIsScaleMethod = (axMemLower == "scalex" || axMemLower == "scaley");
        if (axSlotObj && !node.positional.empty() && !axIsScaleMethod) {
            std::string axObjExpr = std::move(comObjExpr_);
            std::string axMember = std::move(comMemberName_);
            isComMarker_ = false;
            // Step 1: 获取属性对象 (如 ListImages 集合)
            // C29-3: ImageList 的 ListImages 是**真集合对象** (原生复刻, 见
            // vb6forms_memberobj.c), 不能对它发 vb6_ComGetObjectProp —— vb6_com_X 槽里
            // 放的是 HIMAGELIST 实例指针不是 IDispatch, 那样做等于拿垃圾当 vtable。
            std::string collectionExpr;
            {
                if (Symbol::toLower(axMember) == "listimages") {
                    std::string bareIL = imageListNameOfExpr(axObjExpr);
                    if (!bareIL.empty())
                        collectionExpr = "vb6_ImageList_ListImages((void*)vb6_com_" + bareIL + ")";
                } else if (Symbol::toLower(axMember) == "buttons") {
                    // C29-5b: Toolbar1.Buttons —— 宿主槽是 vb6_hwnd_X (真窗口)。
                    std::string bareTB = toolbarNameOfExpr(axObjExpr);
                    if (!bareTB.empty())
                        collectionExpr = "vb6_Toolbar_Buttons((void*)vb6_hwnd_" + bareTB + ")";
                } else if (Symbol::toLower(axMember) == "nodes") {
                    // C29-8b: TreeView1.Nodes —— 宿主槽是 vb6_hwnd_X (真窗口)。
                    std::string bareTV = treeViewNameOfExpr(axObjExpr);
                    if (!bareTV.empty())
                        collectionExpr = "vb6_TreeView_Nodes((void*)vb6_hwnd_" + bareTV + ")";
                    // Fix <vbeclipse>: 跨窗体 TreeView (`frmViewViews.tvwViews.Nodes`) —
                    // 宿主槽是控件句柄访问器文本。
                    if (collectionExpr.empty()) {
                        auto itExtTv = externalTreeViewAccs_.find(axObjExpr);
                        if (itExtTv != externalTreeViewAccs_.end())
                            collectionExpr = "vb6_TreeView_Nodes((void*)" + itExtTv->second + ")";
                    }
                } else {
                    // C29-7: ListView 的 ListItems / ColumnHeaders 同一口径。
                    // ⚠ 宿主槽是 **vb6_hwnd_X** (ListView 是真窗口), 不是 vb6_com_X。
                    std::string memLV = Symbol::toLower(axMember);
                    if (memLV == "listitems" || memLV == "columnheaders") {
                        std::string lvHwnd = listViewHwndExprOf(axObjExpr);
                        if (!lvHwnd.empty())
                            collectionExpr = (memLV == "listitems")
                                ? ("vb6_ListView_ListItems((void*)" + lvHwnd + ")")
                                : ("vb6_ListView_ColumnHeaders((void*)" + lvHwnd + ")");
                    }
                }
                if (collectionExpr.empty())
                    collectionExpr = "vb6_ComGetObjectProp(" + axObjExpr + ", L\"" + axMember + "\")";
            }
            // C29-Data: `Recordset.Fields("name")` —— 取值即 FieldValueStr (不是集合 Item)。
            // memberName 在 axSlotObj 路径上是 axMember (如 "Fields"), 实参 0 是字段名/下标。
            if (Symbol::toLower(axMember) == "fields") {
                std::string dhD = dataSelfHwndExpr(axObjExpr);
                if (!dhD.empty()) {
                    emitExpr(*node.positional[0]);
                    std::string a0 = lastExpr_;
                    bool allNum = !a0.empty() && a0.find_first_not_of("0123456789") == std::string::npos;
                    if (allNum) a0 = "L\"" + a0 + "\"";
                    lastExpr_ = "vb6_Data_FieldValueStr((void*)" + dhD + ", " + a0 + ")";
                    return;
                }
            }
            // Step 2: 调用 Item(idx) 获取集合中的元素
            std::vector<std::string> packedArgs;
            for (size_t i = 0; i < node.positional.size(); i++) {
                std::string packFn = comPackExpr(*node.positional[i]);
                emitExpr(*node.positional[i]);
                { std::string resolved = resolveComMarkerForPack(packFn); if (!resolved.empty()) lastExpr_ = resolved; }
                packedArgs.push_back(packFn + "(" + lastExpr_ + ")");
            }
            int32_t argc = (int32_t)packedArgs.size();
            std::string argsArray = "(void*[]){";
            for (int i = 0; i < argc; i++) {
                if (i > 0) argsArray += ", ";
                argsArray += packedArgs[i];
            }
            argsArray += "}";
            // 返回对象类型 (用于后续 .Picture 等链式访问)
            lastExpr_ = "vb6_ComCallObject(" + collectionExpr + ", L\"Item\", " + argsArray + ", " + std::to_string(argc) + ")";
            return;
        }
        isComMarker_ = false;  // 消费标记
        std::string objExpr = std::move(comObjExpr_);
        std::string memberName = std::move(comMemberName_);

        // C29-Data: Data1.Recordset 的调用面 —— Fields("name")(→FieldValueStr) 与
        // MoveFirst/Last/Next/Previous/Refresh。objExpr 是 vb6_Data_Self( 透传形态。
        if (objExpr.find("vb6_Data_Self(") == 0) {
            if (getenv("C3_CG_TRACE")) {
                FILE* tf = fopen("C:/Users/Administrator/Documents/c3.vb6.pro/.temp/cg3.log", "a");
                if (tf) { fprintf(tf, "[COMBIND-DATA] member=[%s]\n", memberName.c_str()); fclose(tf); }
            }
            std::string dtM = Symbol::toLower(memberName);
            std::vector<std::string> da;
            for (size_t i = 0; i < node.positional.size(); i++) {
                emitExpr(*node.positional[i]);
                da.push_back(lastExpr_);
            }
            if (dtM == "fields") {
                // VB6: Fields("name") / Fields(1) —— 取值即 FieldValueStr
                std::string a0 = da.empty() ? std::string("L\"\"") : da[0];
                // 数字下标转宽字面量 (FieldValueStr 的参数是 nameOrIndex 统一收宽串)
                if (!a0.empty() && a0.find_first_not_of("0123456789") == std::string::npos)
                    a0 = "L\"" + a0 + "\"";
                lastExpr_ = "vb6_Data_FieldValueStr(" + objExpr + ", " + a0 + ")";
                return;
            }
            if (dtM == "movenext")     { lastExpr_ = "vb6_Data_MoveNext(" + objExpr + ")"; return; }
            if (dtM == "moveprevious") { lastExpr_ = "vb6_Data_MovePrevious(" + objExpr + ")"; return; }
            if (dtM == "movefirst")    { lastExpr_ = "vb6_Data_MoveFirst(" + objExpr + ")"; return; }
            if (dtM == "movelast")     { lastExpr_ = "vb6_Data_MoveLast(" + objExpr + ")"; return; }
            if (dtM == "refresh")      { lastExpr_ = "vb6_Data_Refresh(" + objExpr + ")"; return; }
        }
        // C29-4: StatusBar1.SimClick n, dbl —— **判据专用**方法 (无头环境点不了鼠标,
        // RTL 的 vb6_StatusBar_SimClick 程序化发一条真实 WM_NOTIFY, 走完整派发链)。
        // 不对应任何 VB6 语义, 写进 ai/029 边界。
        if (objExpr.find("vb6_hwnd_") == 0) {
            std::string sbBare = objExpr.substr(9);
            if (statusBarVars_.count(sbBare) &&
                (Symbol::toLower(memberName) == "simclick")) {
                std::vector<std::string> sa;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    sa.push_back(lastExpr_);
                }
                lastExpr_ = "vb6_StatusBar_SimClick((void*)" + objExpr + ", "
                          + (sa.size() > 0 ? sa[0] : std::string("0")) + ", "
                          + (sa.size() > 1 ? sa[1] : std::string("0")) + ")";
                return;
            }
        }

        // C29-OLE: OLE 容器的方法全走 vb6_OleCon_* (真 OLE 容器, 非 COM 占位)。
        // objExpr = vb6_hwnd_X (OLE 容器是真窗口控件)。不拦的话会拼成
        // `vb6_ComCall(vb6_hwnd_X, L"CreateEmbed", …)` —— 对 HWND 当 IDispatch 用。
        if (!oleConNameOfExpr(objExpr).empty()) {
            std::string oleSlot = "vb6_hwnd_" + oleConNameOfExpr(objExpr);
            std::string oleMember = memberName;
            std::transform(oleMember.begin(), oleMember.end(), oleMember.begin(), ::tolower);
            std::vector<std::string> oa;
            for (size_t i = 0; i < node.positional.size(); i++) {
                emitExpr(*node.positional[i]);
                oa.push_back(lastExpr_);
            }
            // VB6 的 CreateEmbed(SourceDoc, SourceItem) 两参都可省
            auto oleArg = [&](size_t i, const char* dflt) -> std::string {
                return (oa.size() > i && !oa[i].empty()) ? oa[i] : dflt;
            };
            if (oleMember == "createembed") {
                lastExpr_ = "vb6_OleCon_CreateEmbed((void*)" + oleSlot + ", "
                          + oleArg(0, "NULL") + ")";
                return;
            }
            if (oleMember == "createlink") {
                lastExpr_ = "vb6_OleCon_CreateLink((void*)" + oleSlot + ", "
                          + oleArg(0, "NULL") + ", " + oleArg(1, "NULL") + ")";
                return;
            }
            if (oleMember == "doverb") {
                lastExpr_ = "vb6_OleCon_DoVerb((void*)" + oleSlot + ", "
                          + (oa.empty() ? std::string("0") : oa[0]) + ")";
                return;
            }
            if (oleMember == "close") {
                lastExpr_ = "vb6_OleCon_Close((void*)" + oleSlot + ")";
                return;
            }
            if (oleMember == "insertobjdlg") {
                lastExpr_ = "vb6_OleCon_InsertObjDlg((void*)" + oleSlot + ")";
                return;
            }
            if (oleMember == "savetofile") {
                lastExpr_ = "vb6_OleCon_SaveToFile((void*)" + oleSlot + ", "
                          + oleArg(0, "L\"\"") + ")";
                return;
            }
            if (oleMember == "readfromfile") {
                lastExpr_ = "vb6_OleCon_ReadFromFile((void*)" + oleSlot + ", "
                          + oleArg(0, "L\"\"") + ")";
                return;
            }
            if (oleMember == "copy") {
                lastExpr_ = "vb6_OleCon_Copy((void*)" + oleSlot + ")";
                return;
            }
            if (oleMember == "paste" || oleMember == "pastespecialdlg") {
                lastExpr_ = "vb6_OleCon_Paste((void*)" + oleSlot + ")";
                return;
            }
            // Delete / FetchVerbs 等暂未实现的先落默认 COM 路径 (会静默答空, 与 029 §六-2 同)。
        }

        // P20-39: ImageList 原生复刻 —— ListImages.Add/Remove/Clear/Item 直接打在 RTL 上,
        // 不要去养一个假 IDispatch (mscomctl.ocx 未注册, 且 32 位 inproc OCX 进不了 x64 进程)。
        // 这里必须在通用 COM 分支**之前**拦下, 否则 next 步就会拼出
        // `vb6_ComCall(vb6_ComGetObjectProp(vb6_com_X, L"ListImages"), L"Add", ...)`。
        {
            std::string ilMember = memberName;
            std::transform(ilMember.begin(), ilMember.end(), ilMember.begin(), ::tolower);
            // objExpr 是 `vb6_ComGetObjectProp(vb6_com_X, L"ListImages")`(都在 ② 形态上),
            // 所以按裸控件名认, 不能直接比 objExpr。
            if (!imageListNameOfExpr(objExpr).empty()) {
                if (ilMember == "add") {
                    std::vector<std::string> a;
                    for (size_t i = 0; i < node.positional.size(); i++) {
                        emitExpr(*node.positional[i]);
                        a.push_back(lastExpr_);
                    }
                    // VB6: ListImages.Add(index, key, picture) —— 位置对应, 缺的省掉
                    int32_t ilIdx = 0;
                    std::string ilKey = "NULL";
                    std::string ilPic = "NULL";
                    if (a.size() > 0) {
                        size_t pd = a[0].find('(');
                        std::string first = (pd != std::string::npos) ? a[0].substr(pd + 1) : a[0];
                        bool isNum = !first.empty();
                        for (char ch : first) if (!isdigit((unsigned char)ch)) { isNum = false; break; }
                        if (isNum) ilIdx = (int32_t)atol(first.c_str());
                    }
                    std::string ilSlot = "vb6_com_" + imageListNameOfExpr(objExpr);
                    if (a.size() > 1) ilKey = a[1];
                    if (a.size() > 2) ilPic = a[2];
                    // C29-3: Add **返回 ListImage 对象** (VB6 语义: `Set img = .Add(..)`),
                    // 不再是 int 下标 —— 旧实现返回下标时 `Set img` 接住的是个立即数,
                    // `img.Key` 会拿它当指针解引用 (0xC000041D)。
                    std::string ilAddCall = "vb6_ImageList_ListImages_Add((void*)" + ilSlot + ", "
                              + std::to_string(ilIdx) + ", (const wchar_t*)(" + ilKey + "), "
                              + ilPic + ")";
                    // 老写法 `n = .Add(..)` (Let, 非 Set): VB6 取的是 ListImage 的
                    // **默认属性 Index** —— 不转的话就是 `void*` 赋给整型, C 直接编不过。
                    lastExpr_ = memObjLetScalar_
                              ? ("vb6_ComGetIntProp(" + ilAddCall + ", L\"Index\")")
                              : ilAddCall;
                    return;
                }
                if (ilMember == "remove") {
                    std::string ilSlot = "vb6_com_" + imageListNameOfExpr(objExpr);
                    std::string rmArg = "NULL";
                    for (size_t i = 0; i < node.positional.size(); i++) {
                        emitExpr(*node.positional[i]);
                        rmArg = lastExpr_;
                        break;
                    }
                    // VB6 的 Remove 两种都吃: 数字按下标、字符串按 Key。
                    // 数字必须走下标版 —— 转成 wchar_t* 会解引用地址 1, 直接崩。
                    bool rmIsNum = false;
                    {
                        size_t lp = rmArg.find('(');
                        std::string t = (lp != std::string::npos) ? rmArg.substr(lp + 1) : rmArg;
                        size_t end = t.find_first_of("),", 0);
                        if (end != std::string::npos) t = t.substr(0, end);
                        rmIsNum = !t.empty();
                        for (char ch : t) if (!isdigit((unsigned char)ch)) { rmIsNum = false; break; }
                    }
                    if (rmIsNum) {
                        lastExpr_ = "vb6_ImageList_RemoveAtIndex((void*)" + ilSlot + ", " + rmArg + ")";
                    } else {
                        lastExpr_ = "vb6_ImageList_RemoveImage((void*)" + ilSlot
                                  + ", (const wchar_t*)(" + rmArg + "), 0)";
                    }
                    return;
                }
                if (ilMember == "clear") {
                    std::string ilSlot = "vb6_com_" + imageListNameOfExpr(objExpr);
                    lastExpr_ = "vb6_ImageList_ClearImages((void*)" + ilSlot + ")";
                    return;
                }
            }
        }

        // P20-40: StatusBar 原生复刻 —— Panels.Add/Remove/Clear 直接打在 RTL 上。
        // 同样必须在通用 COM 分支**之前**拦下, 否则会拼出
        // `vb6_ComCall(vb6_ComGetObjectProp(vb6_com_X, L"Panels"), L"Add", ...)`。
        {
            std::string sbMember = memberName;
            std::transform(sbMember.begin(), sbMember.end(), sbMember.begin(), ::tolower);
            std::string sbHwnd = statusBarHwndVarOfExpr(objExpr);
            if (!sbHwnd.empty()) {
                if (sbMember == "add") {
                    std::vector<std::string> a;
                    for (size_t i = 0; i < node.positional.size() && i < 3; i++) {
                        emitExpr(*node.positional[i]);
                        a.push_back(lastExpr_);
                    }
                    // Panels.Add([index], [key], [text]) —— VB6 的位置对应, 缺的省掉。
                    int32_t sbIdx = 0;
                    std::string sbKey = "NULL";
                    std::string sbTxt = "L\"\"";
                    if (a.size() > 0) {
                        size_t pd = a[0].find('(');
                        std::string first = (pd != std::string::npos) ? a[0].substr(pd + 1) : a[0];
                        bool isNum = !first.empty();
                        for (char ch : first) if (!isdigit((unsigned char)ch)) { isNum = false; break; }
                        if (isNum) sbIdx = (int32_t)atol(first.c_str());
                    }
                    if (a.size() > 1 && !a[1].empty() && a[1] != "vb6_ComPackMissing(...)") sbKey = a[1];
                    if (a.size() > 2 && !a[2].empty() && a[2] != "vb6_ComPackMissing(...)") sbTxt = a[2];
                    lastExpr_ = "vb6_StatusBar_AddPanel((void*)" + sbHwnd + ", "
                              + std::to_string(sbIdx) + ", (const wchar_t*)(" + sbKey + "), "
                              + sbTxt + ")";
                    return;
                }
                if (sbMember == "remove") {
                    std::string rmArg = "0";
                    for (size_t i = 0; i < node.positional.size(); i++) {
                        emitExpr(*node.positional[i]);
                        rmArg = lastExpr_;
                        break;
                    }
                    // 与 ImageList.Remove 同样的坑: 数字必须走下标版, 转成 wchar_t*
                    // 会把地址 1 当字符串解引用, 直接崩。
                    bool rmIsNum = false;
                    {
                        size_t lp = rmArg.find('(');
                        std::string t = (lp != std::string::npos) ? rmArg.substr(lp + 1) : rmArg;
                        size_t end = t.find_first_of("),", 0);
                        if (end != std::string::npos) t = t.substr(0, end);
                        rmIsNum = !t.empty();
                        for (char ch : t) if (!isdigit((unsigned char)ch)) { rmIsNum = false; break; }
                    }
                    if (rmIsNum) {
                        lastExpr_ = "vb6_StatusBar_RemovePanel((void*)" + sbHwnd
                                  + ", NULL, " + rmArg + ")";
                    } else {
                        lastExpr_ = "vb6_StatusBar_RemovePanel((void*)" + sbHwnd
                                  + ", (const wchar_t*)(" + rmArg + "), 0)";
                    }
                    return;
                }
                if (sbMember == "clear") {
                    lastExpr_ = "vb6_StatusBar_ClearPanels((void*)" + sbHwnd + ")";
                    return;
                }
            }
        }

        // Fix 057: Me.Controls.Add(ProgID, Name) → vb6_Form_ControlsAdd(hwnd, L"ProgID", L"Name")
        // 检测链式COM: objExpr = "vb6_ComGetObjectProp(vb6_hwnd_Form1, L"Controls")", memberName = "Add"
        {
            std::string memLower = memberName;
            std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
            std::string objExprLower = objExpr;
            std::transform(objExprLower.begin(), objExprLower.end(), objExprLower.begin(), ::tolower);
            if (memLower == "add" &&
                objExpr.find("vb6_ComGetObjectProp(vb6_hwnd_") == 0 &&
                objExprLower.find("l\"controls\")") != std::string::npos) {
                // 提取 form HWND 变量: vb6_ComGetObjectProp(vb6_hwnd_Form1, L"Controls") → vb6_hwnd_Form1
                size_t parenStart = strlen("vb6_ComGetObjectProp(");
                size_t commaPos = objExpr.find(", ", parenStart);
                std::string hwndExpr = objExpr.substr(parenStart, commaPos - parenStart);
                // 发射参数: Controls.Add(ProgID, Name)
                std::vector<std::string> argExprs;
                for (size_t i = 0; i < node.positional.size() && i < 2; i++) {
                    emitExpr(*node.positional[i]);
                    argExprs.push_back(lastExpr_);
                }
                if (argExprs.size() >= 2) {
                    lastExpr_ = "vb6_Form_ControlsAdd(" + hwndExpr + ", " + argExprs[0] + ", " + argExprs[1] + ")";
                    return;
                }
                // 参数不足, 降级为普通COM调用
            }
        }

        // tB Interface 契约 (B04): x.Method(args) → x->vt-><slot>(x, args)
        if (!ivrefIfaceName_.empty()) {
            const std::string ivIface = ivrefIfaceName_;
            ivrefIfaceName_.clear();
            std::string slotKey;
            if (const IfaceView* v = ivLookupIface(ivIface)) {
                slotKey = ivSlotKeyForMember(*v, memberName);
            }
            if (!slotKey.empty()) {
                std::vector<std::string> ivArgs;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    ivArgs.push_back(lastExpr_);
                }
                for (auto& named : node.named) {
                    emitExpr(*named.value);
                    ivArgs.push_back(lastExpr_);
                }
                std::string call = objExpr + "->vt->" + cIdent(slotKey) + "(" + objExpr;
                for (size_t i = 0; i < ivArgs.size(); i++) call += ", " + ivArgs[i];
                call += ")";
                lastExpr_ = call;
                return;
            }
            // 接口里没有这个成员: 契约检查已报过错, 落回后续路径不再级联
        }

        // P6.4: 接口引用方法调用 (Dim x As IFoo → x.Method → x.vtbl->Method(x.obj, args))
        {
            std::string objLower = objExpr;
            std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
            auto itIfaceVar = knownIfaceVars_.find(objLower);
            if (itIfaceVar != knownIfaceVars_.end() && !isEarlyBoundCom_) {
                std::string ifaceName = itIfaceVar->second;
                std::string ifaceId = cIdent(ifaceName);
                // 生成参数列表
                std::vector<std::string> callArgs;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    callArgs.push_back(lastExpr_);
                }
                for (auto& named : node.named) {
                    emitExpr(*named.value);
                    callArgs.push_back(lastExpr_);
                }
                std::string argsStr;
                for (size_t i = 0; i < callArgs.size(); i++) {
                    if (i > 0) argsStr += ", ";
                    argsStr += callArgs[i];
                }
                // x.vtbl->Method(x.obj, args...)
                std::string call = objExpr + ".vtbl->" + cIdent(memberName) + "(" + objExpr + ".obj";
                if (!argsStr.empty()) call += ", " + argsStr;
                call += ")";
                lastExpr_ = call;
                return;
            }
        }

        // P6.3: 前期绑定 — 利用类型签名确定返回类型，统一走后期绑定(IDispatch)
        // 原因: vtable直接调用的函数签名可能不是VARIANT* (如get_Count用long*)，
        // vb6_ComVtableGet*辅助函数统一用VARIANT*签名会导致调用错误
        if (isEarlyBoundCom_ && earlyBoundSym_) {
            isEarlyBoundCom_ = false;
            const Symbol* comSym = earlyBoundSym_;
            earlyBoundSym_ = nullptr;

            // 查找方法签名以确定返回类型
            std::string memLower = memberName;
            std::transform(memLower.begin(), memLower.end(), memLower.begin(), ::tolower);
            auto it = comSym->comMethods.find(memLower);
            bool hasArgs = (!node.positional.empty() || !node.named.empty());


            // 无参属性Get: 用后期绑定属性读取，根据返回类型选函数
            if (it != comSym->comMethods.end() && it->second.isPropertyGet && !hasArgs) {
                const auto& sig = it->second;
                std::string returnType = mapType(sig.returnType);
                if (returnType == "BSTR") {
                    lastExpr_ = "vb6_ComGetStringProp(" + objExpr + ", L\"" + memberName + "\")";
                } else if (returnType == "int32_t" || returnType == "int16_t") {
                    lastExpr_ = "vb6_ComGetIntProp(" + objExpr + ", L\"" + memberName + "\")";
                } else if (returnType == "double" || returnType == "float") {
                    lastExpr_ = "vb6_ComGetDoubleProp(" + objExpr + ", L\"" + memberName + "\")";
                } else if (returnType == "void*") {
                    lastExpr_ = "vb6_ComGetObjectProp(" + objExpr + ", L\"" + memberName + "\")";
                } else {
                    lastExpr_ = "vb6_ComGetStringProp(" + objExpr + ", L\"" + memberName + "\")";
                }
                return;
            }
            // P24-07: 有参数的前期绑定方法 → 利用签名选择类型化COM调用函数
            if (it != comSym->comMethods.end() && hasArgs) {
                const auto& sig = it->second;
                std::string returnType = mapType(sig.returnType);
                std::vector<std::string> packedArgs;
                // Fix 160-com-byref: ByRef 对象出参 (Chilkat ObjectOf2/ArrayOf2 的 obj/arr 形参)。
                // 这类形参 type=Object 且 !isByVal; 调用方必须以 void** 形式把变量地址传出,
                // 才能接住方法写入的新对象指针 (VT_DISPATCH|VT_BYREF)。仅当实参是简单对象变量
                // (已知 ComClass/Object 变量) 时才取址 —— 表达式出参需临时量, 超出本修复范围,
                // 退化为按值打包以维持既有行为, 不引入编译错误。
                auto isKnownObjectVarName = [&](const std::string& e) -> bool {
                    std::string l = e;
                    std::transform(l.begin(), l.end(), l.begin(), ::tolower);
                    return knownObjectVars_.count(l) || knownTypedComVars_.count(l);
                };
                auto byRefObjectParamAt = [&](size_t pos) -> bool {
                    if (pos < sig.params.size()) {
                        const auto& p = sig.params[pos];
                        // Fix 160-com-byref: COM 对象指针形参 (VT_PTR(VT_DISPATCH)) 在 C3 类型映射里
                        // 常落到 Variant, 故不苛求 type==Object, 只要求方向非 In (Out/InOut)。
                        // 真正的对象判定交给实参侧 isKnownObjectVarName —— 非对象实参不会走此分支。
                        return !p.isByVal;
                    }
                    return false;
                };
                auto byRefObjectParamNamed = [&](const std::string& nm) -> bool {
                    std::string nl = Symbol::toLower(nm);
                    for (const auto& p : sig.params)
                        if (Symbol::toLower(p.name) == nl && !p.isByVal && p.type == Vb6Type::Object)
                            return true;
                    return false;
                };
                for (size_t i = 0; i < node.positional.size(); i++) {
                    std::string packFn = comPackExpr(*node.positional[i]);
                    emitExpr(*node.positional[i]);
                    { std::string resolved = resolveComMarkerForPack(packFn); if (!resolved.empty()) lastExpr_ = resolved; }
                    if (byRefObjectParamAt(i)
                        && node.positional[i]->kind == ASTNodeKind::IdentifierExpr
                        && isKnownObjectVarName(lastExpr_)) {
                        packedArgs.push_back("vb6_ComPackObjectRef((void**)&" + lastExpr_ + ")");
                    } else {
                        packedArgs.push_back(packFn + "(" + lastExpr_ + ")");
                    }
                }
                for (auto& named : node.named) {
                    std::string packFn = comPackExpr(*named.value);
                    emitExpr(*named.value);
                    { std::string resolved = resolveComMarkerForPack(packFn); if (!resolved.empty()) lastExpr_ = resolved; }
                    if (byRefObjectParamNamed(named.name)
                        && named.value->kind == ASTNodeKind::IdentifierExpr
                        && isKnownObjectVarName(lastExpr_)) {
                        packedArgs.push_back("vb6_ComPackObjectRef((void**)&" + lastExpr_ + ")");
                    } else {
                        packedArgs.push_back(packFn + "(" + lastExpr_ + ")");
                    }
                }
                int32_t argc = (int32_t)packedArgs.size();
                std::string argsArray = "(void*[]){";
                for (int i = 0; i < argc; i++) {
                    if (i > 0) argsArray += ", ";
                    argsArray += packedArgs[i];
                }
                argsArray += "}";
                std::string callArgs = objExpr + ", L\"" + memberName + "\", " + argsArray + ", " + std::to_string(argc);
                if (returnType == "BSTR") {
                    lastExpr_ = "vb6_ComCallBSTR(" + callArgs + ")";
                } else if (returnType == "int32_t" || returnType == "int16_t") {
                    lastExpr_ = "vb6_ComCallInt(" + callArgs + ")";
                } else if (returnType == "double" || returnType == "float") {
                    lastExpr_ = "vb6_ComCallDouble(" + callArgs + ")";
                } else if (returnType == "void*") {
                    lastExpr_ = "vb6_ComCallObject(" + callArgs + ")";
                } else {
                    lastExpr_ = "vb6_ComCall(" + callArgs + ")";  // 未知返回类型: 返回void*
                }
                isComMarker_ = false;
                return;
            }
            // 有参数的方法/属性Put/签名未找到 → 降级为后期绑定 (fall through)
            isEarlyBoundCom_ = false;
        }
        isEarlyBoundCom_ = false;

        if (!node.positional.empty() || !node.named.empty() || !node.omittedArgs.empty()) {
            // 有参数: obj.Method(args) → vb6_ComCall(obj, L"Method", variantArgs, argc)
            // Fix 104: 省略实参必须在该形参位置传 VT_ERROR/DISP_E_PARAMNOTFOUND, 否则后续
            // 实参前移 (c.Add "x", , , 1 的 1 会被当成第 2 个形参 Key 而非第 4 个 After).
            // 解析器 (Fix 142 口径) 为每个形参位置保留槽位, 并把省略位下标记入 omittedArgs
            // (带括号 obj.M(a, , c) 与无括号 obj.M a, , c 两条路径一致),
            // 故这里只需按 positional 下标对齐形参表, 省略位补 vb6_ComPackMissing().
            std::vector<std::string> packedArgs;
            auto emitComArg = [&](const ExprPtr& argExpr) {
                std::string packFn = comPackExpr(*argExpr);
                emitExpr(*argExpr);
                { std::string resolved = resolveComMarkerForPack(packFn); if (!resolved.empty()) lastExpr_ = resolved; }
                packedArgs.push_back(packFn + "(" + lastExpr_ + ")");
            };
            for (size_t pos = 0; pos < node.positional.size(); pos++) {
                if (node.omittedArgs.count(pos)) {
                    packedArgs.push_back("vb6_ComPackMissing()");
                    continue;
                }
                emitComArg(node.positional[pos]);
            }
            for (auto& named : node.named) {
                emitComArg(named.value);
            }

            int32_t argc = (int32_t)packedArgs.size();
            std::string argsArray;
            argsArray = "(void*[]){";
            for (int i = 0; i < argc; i++) {
                if (i > 0) argsArray += ", ";
                argsArray += packedArgs[i];
            }
            argsArray += "}";

            lastExpr_ = "vb6_ComCall(" + objExpr + ", L\"" + memberName + "\", " +
                        argsArray + ", " + std::to_string(argc) + ")";
            isComMarker_ = false;  // P24-04: 参数emission可能设置脏isComMarker_
            return;
        } else {
            // 无参数: obj.Method() → vb6_ComCall(obj, L"Method", NULL, 0)
            lastExpr_ = "vb6_ComCall(" + objExpr + ", L\"" + memberName + "\", NULL, 0)";
            isComMarker_ = false;  // P24-04: 清除残留标记
            return;
        }
    }

    // Fix 044b: Check if callee is a WithMemberExpr class method that might need
    // Optional param padding. If so, skip the early-return and split the callee
    // to allow Optional padding in the normal flow below.
    bool needsSplitForOptionalPad = false;
    if (node.callee && node.callee->kind == ASTNodeKind::WithMemberExpr) {
        if (!withObjectInfoStack_.empty()) {
            const auto& info = withObjectInfoStack_.back();
            if (info.kind == WithObjKind::ClassInstance && !info.className.empty()) {
                needsSplitForOptionalPad = true;
            }
        }
    }
    // Fix 090c: 类成员方法无实参调用 (obj.M() / Trim(obj.M())) — 若 callee 已被 MAE 生成
    // 为完整调用 vb6_cls_X_M(obj) 且该方法声明有参数, 需拆分走 Optional padding 流程
    // (见下方 return 判定与 4040 拆分条件).
    bool needsSplitForMissingArgs = false;
    if (!needsSplitForOptionalPad && node.callee
        && node.callee->kind == ASTNodeKind::MemberAccessExpr
        && node.positional.empty() && node.named.empty()) {
        auto& maeFix090c = static_cast<MemberAccessExpr&>(*node.callee);
        std::string clsFix090c =
            maeFix090c.object ? inferClassTypeOfExpr(*maeFix090c.object) : "";
        if (!clsFix090c.empty()) {
            std::vector<ParameterInfo> ppFix090c;
            bool ibFix090c = false;
            if (findClassMemberCallParams(clsFix090c, maeFix090c.memberName,
                                          ppFix090c, ibFix090c)
                && !ppFix090c.empty() && !ibFix090c) {
                needsSplitForMissingArgs = true;
            }
        }
    }

    // 如果callee已经是func(args)形式(如类方法调用 vb6_Counter_GetCount(c)),
    // 且IndexOrCallExpr没有额外参数, 直接使用callee避免双重括号
    if (callee.size() >= 2 && callee.back() == ')' && node.positional.empty() && node.named.empty()) {
        // 检查是否是完整的函数调用（以右括号结尾且匹配左括号）
        int depth = 0;
        bool isCompleteCall = false;
        for (int i = (int)callee.size() - 2; i >= 0; i--) {
            if (callee[i] == ')') depth++;
            else if (callee[i] == '(') {
                if (depth == 0) { isCompleteCall = true; break; }
                depth--;
            }
        }
        if (isCompleteCall && !needsSplitForOptionalPad && !needsSplitForMissingArgs) {
            // Fix 146 回归 (vbman cLogs C2198 / T2 ExtShow): external form show 转发
            // (cgen_expr_member_form_builtin.inc) 只交付 vb6_form_show_<Form>(NULL),
            // modal 实参原本只由本文件的"有参拆包"路径追加; 无参调用 (Frm.Show) 走
            // 这里直接复用 callee, 会缺 modal 形参 -> C2198. 无参即补默认 modal=0
            // (VB6 语义: 无参 Show = modeless).
            if (callee.compare(0, 14, "vb6_form_show_") == 0) {
                lastExpr_ = callee.substr(0, callee.size() - 1) + ", 0)";
                return;
            }
            lastExpr_ = callee;
            return;
        }
        // needsSplitForMissingArgs / needsSplitForOptionalPad: 需继续执行拆分+补齐流程
    }

    // 如果callee已经是func()形式(无参内置函数调用如vb6_Now())，
    // 需要拆开重组为func(args)，因为IndexOrCallExpr表示带参数调用
    bool calleeIsZeroArgCall = false;
    if (callee.size() >= 2 && callee.substr(callee.size() - 2) == "()") {
        calleeIsZeroArgCall = true;
        callee = callee.substr(0, callee.size() - 2);
    }

    // Fix 083e: Variant 数组嵌套索引 — vGateway(lIdx)(0) 不是 f(a,b) 而是嵌套元素访问.
    // callee 若是 vb6_VariantArrayGet(&arr, idx) 完整调用, 外层 (0) 应对内层结果再取元素.
    // (否则 P6.5 拆开会把索引合并成 vb6_VariantArrayGet(&arr, idx, 0) → C2197 参数太多)
    // Fix 084f: 旧实现用 (vb6_VARIANT){<内层调用>} 复合字面量以值初始化结构体,
    // 首成员是 vb6_vartype → 触发 C2440 "vb6_VARIANT → vb6_vartype".
    // 改用按值辅助函数 vb6_VariantArrayGetVal(内层调用, 外层索引).
    if (callee.compare(0, 20, "vb6_VariantArrayGet(") == 0 && !node.positional.empty()) {
        std::string innerCall = callee;
        std::string outerIdx;
        emitExpr(*node.positional[0]);
        // Fix 084o: 外层索引为 Variant 时转 Long
        outerIdx = toLongIfVariant(std::move(lastExpr_), node.positional[0].get());
        lastExpr_ = "vb6_VariantArrayGetVal(" + innerCall + ", " + outerIdx + ")";
        return;
    }

    // P6.5修复: 如果callee已经是func(obj)形式(如类方法调用 vb6_Button_SetCaption(btn)),
    // 且IndexOrCallExpr有额外参数, 需要拆开重组为func(obj, userArgs...),
    // 避免生成 func(obj)(userArgs) 双重括号
    std::string classMethodObjArg;  // 如果非空, 表示callee已被拆开, 需要前置此参数

    // Fix 015: 若 MemberAccessExpr.Fix015 路径已通过 pendingChainObj_ 交付对象参数,
    // 移交给 classMethodObjArg (随后会被前置到参数列表).
    // 此时 callee 是裸函数名 (如 "vb6_cDataBase_Exec"), 上面的 split 路径因
    // callee.back() != ')' 不会触发, 故不会被双重设置.
    if (!pendingChainObj_.empty()) {
        classMethodObjArg = std::move(pendingChainObj_);
        pendingChainObj_.clear();
    }

    if (callee.size() >= 2 && callee.back() == ')'
        && ((!node.positional.empty() || !node.named.empty())
            || needsSplitForOptionalPad || needsSplitForMissingArgs)) {
        // 检查是否是完整的函数调用（以右括号结尾且匹配左括号）
        int depth = 0;
        int openPos = -1;
        for (int i = (int)callee.size() - 2; i >= 0; i--) {
            if (callee[i] == ')') depth++;
            else if (callee[i] == '(') {
                if (depth == 0) { openPos = i; break; }
                depth--;
            }
        }
        if (openPos > 0) {
            // callee = "funcName(existingArgs)" → 拆开
            std::string funcPart = callee.substr(0, openPos);
            classMethodObjArg = callee.substr(openPos + 1, callee.size() - openPos - 2);
            callee = funcPart;
            // Fix 086: 链式默认属性调用 — 内层是无参 prop_get_(如 .Root("data")) 或
            // COM 调用 (Dic(N)(RouteName)) 时, 内层返回 COM 对象, 外层索引是对返回
            // 对象的 Item 调用. 不能并入内层参数表 (C2197 参数太多).
            // 有声明参数的 prop_get (Prop(k)) 仍走合并 (内层只发了this).
            bool innerIsChainedObj = false;
            if (!node.positional.empty() && node.named.empty()) {
                // funcPart 是拆开后的裸函数名 (无括号)
                if (funcPart == "vb6_ComCall" || funcPart == "vb6_ComCallObject"
                    || funcPart == "vb6_ComGetObjectProp"
                    || funcPart == "vb6_VariantFromComResult") {
                    innerIsChainedObj = true;
                } else {
                    size_t gp = funcPart.find("_prop_get_");
                    if (gp != std::string::npos) {
                        std::string propName = funcPart.substr(gp + 10);
                        Symbol* propSym86 = symTab_.lookupModuleByKind(propName, SymbolKind::PropertyGet);
                        if (propSym86 && propSym86->params.empty()) {
                            innerIsChainedObj = true;
                        } else {
                            // Fix 092w: 内层 prop_get 调用已自带实参 (classMethodObjArg
                            // 顶层段数 >= 2, 即 this + 至少 1 个实参) 说明内层调用已自足,
                            // 外层索引是对其返回对象 (Variant/COM 对象) 的默认成员访问,
                            // 不能并入内层参数表 (否则实参个数超出签名).
                            // Demo_Database 506: TestDB.Rows(1)("score") — Rows 是 cCollection
                            // 字段, (1) 取元素返回 Variant(Dictionary), ("score") 应在其结果上
                            // 走默认成员 Item 的后期绑定调用.
                            int topArgs092w = classMethodObjArg.empty() ? 0 : 1;
                            int depth092w = 0;
                            for (char ch092w : classMethodObjArg) {
                                if (ch092w == '(' || ch092w == '[' || ch092w == '{') depth092w++;
                                else if (ch092w == ')' || ch092w == ']' || ch092w == '}') depth092w--;
                                else if (ch092w == ',' && depth092w == 0) topArgs092w++;
                            }
                            if (topArgs092w >= 2) innerIsChainedObj = true;
                        }
                    }
                }
            }
            if (innerIsChainedObj) {
                std::string innerObj = funcPart + "(" + classMethodObjArg + ")";
                // Fix 092w: 内层函数返回 vb6_VARIANT (如 cCollection.prop_get_Item) 时,
                // vb6_ComCall 的 obj 形参是 void* — 需先经 vb6_VariantToObjectVal 提取
                // 对象指针, 否则 vb6_VARIANT 直传 → C2440. (无参 prop_get 返回 void* 的
                // 既有路径保持直传不变.)
                if (variantReturnFuncs_ && variantReturnFuncs_->count(funcPart)) {
                    innerObj = "vb6_VariantToObjectVal(" + innerObj + ")";
                }
                std::vector<std::string> packedArgs86;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    std::string packFn86 = comPackExpr(*node.positional[i]);
                    emitExpr(*node.positional[i]);
                    { std::string resolved86 = resolveComMarkerForPack(packFn86); if (!resolved86.empty()) lastExpr_ = resolved86; }
                    packedArgs86.push_back(packFn86 + "(" + lastExpr_ + ")");
                }
                int32_t argc86 = (int32_t)packedArgs86.size();
                std::string argsArray86 = "(void*[]){";
                for (int i = 0; i < argc86; i++) {
                    if (i > 0) argsArray86 += ", ";
                    argsArray86 += packedArgs86[i];
                }
                argsArray86 += "}";
                lastExpr_ = "vb6_VariantFromComResult(vb6_ComCall(" + innerObj + ", L\"Item\", "
                          + argsArray86 + ", " + std::to_string(argc86) + "))";
                return;
            }
        }
    }

    // 位置参数
    // 需要检查被调用函数的参数签名: ByRef参数在调用点需要传指针(&arg)