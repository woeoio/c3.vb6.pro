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
            bool zaSlot = formCtrlSlot(comObjExpr_, zaType, zaHwnd);
            std::string zaFn = zaSlot ? controlZeroArgMethod(zaType, tvMem) : std::string();
            // 账 #196: 同一个接收者问法下的**一个实参方法**（TextHeight/TextWidth）。
            // 零实参那张表命中不了它，两条各查一次、各拼各的实参表（零实参只交句柄）。
            std::string oaFn = zaSlot ? controlOneArgMethod(zaType, tvMem) : std::string();
            bool slClear = !zaFn.empty();
            if (tvSim || tbSim || dtSim || mvSim || slSim || slClear || !oaFn.empty()) {
                isComMarker_ = false;
                std::vector<std::string> sa;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    sa.push_back(lastExpr_);
                }
                std::string a0 = sa.size() > 0 ? sa[0] : std::string("0");
                std::string a1 = sa.size() > 1 ? sa[1] : std::string("0");
                lastExpr_ = !oaFn.empty()
                    ? (oaFn + "((void*)" + zaHwnd + ", " + a0 + ")")
                    : slClear
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
                          || !toolbarNameOfExpr(comObjExpr_).empty());
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
            // 账 #299 第二十一刀: 这张表以前在这里也抄了一份 (listimages/buttons/nodes/
            // listitems+columnheaders 四族，**没有 panels**) ⇒ `Set po = SB.Panels(1)` 与
            // `With SB.Panels(1)` 的头一直发成拿 HWND 当 IDispatch。现在问唯一那一处，
            // 认不出来才兜底。
            std::string collectionExpr = memberCollectionObjectExpr(axObjExpr, axMember);
            if (collectionExpr.empty())
                collectionExpr = "vb6_ComGetObjectProp(" + axObjExpr + ", L\"" + axMember + "\")";
            // Step 2: 调用 Item(idx) 获取集合中的元素
            std::vector<std::string> packedArgs;
            for (size_t i = 0; i < node.positional.size(); i++) {
                std::string packFn = comPackExpr(*node.positional[i]);
                emitExpr(*node.positional[i]);
                { std::string resolved = resolveComMarkerForPack(packFn); if (!resolved.empty()) lastExpr_ = resolved; }
                packedArgs.push_back(packArgValue(packFn, lastExpr_));
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

        // 账 #278 §B124: 这一格 (`vb6_Data_Self(` 透传形态的直译) 与 §B124 列的那几格同代、
        // 同样恒假，撤掉；顺带带走里面那台把日志路径写死在本机用户目录上的 getenv 探针。
        // 账 #196 第三条: 窗体型接收者的 `ScaleX`/`ScaleY` (`picA.ScaleX(…)`、`Me.ScaleX(…)`)。
        // 不拦就落到下面那条通用 COM 方法调用 —— `vb6_ComCallDouble(vb6_hwnd_X, L"ScaleX", …)`：
        // 对一枚假 IDispatch 发 Invoke ⇒ 编得过、链接过、跑起来回 0（#143 那一族在方法面上的第三味；
        // 上面 152 行那条 `axIsScaleMethod` 只是让它别走更错的"集合 + Item"那一形）。
        // 换算只吃那两个显式的 from/to（接收者自己的 ScaleMode 由发码侧算成参数交进来），
        // 所以四处接收者共用一处实现 `vb6_ScaleUnitX/Y`（见 vb6rtl_com.c 那段注释）。
        {
            std::string scMem = Symbol::toLower(memberName);
            FrmControlType scType = FrmControlType::Unknown;
            std::string scHwndUnused;
            std::string scFn = formCtrlSlot(objExpr, scType, scHwndUnused)
                             ? controlScaleMethod(scType, scMem) : std::string();
            if (!scFn.empty()) {
                std::vector<std::string> sc;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    sc.push_back(lastExpr_);
                }
                std::string sa0 = sc.size() > 0 ? sc[0] : std::string("0");
                std::string sa1 = sc.size() > 1 ? sc[1] : std::string("1");
                std::string sa2 = sc.size() > 2 ? sc[2] : std::string("3");
                // 与这一族其他"自己拼完整调用"的出口同一动作：标记位要清掉，
                // 否则语句路会以为还有一个 COM 成员调用没消费（#143 那一族的另一半）。
                isComMarker_ = false;
                comObjExpr_.clear();
                comMemberName_.clear();
                lastExpr_ = scFn + "(" + sa0 + ", " + sa1 + ", " + sa2 + ")";
                return;
            }
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
                // 账 #278 §B72 第十四刀: 与语句码头问同一张表
                std::string fnIlB = controlZeroArgMethod(FrmControlType::ImageList, ilMember);
                if (!fnIlB.empty()) {
                    std::string ilSlot = "vb6_com_" + imageListNameOfExpr(objExpr);
                    lastExpr_ = fnIlB + "((void*)" + ilSlot + ")";
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
                // 账 #278 §B72 第十四刀: 同上（表里 (StatusBar, clear) 那一行）
                std::string fnSbB = controlZeroArgMethod(FrmControlType::StatusBar, sbMember);
                if (!fnSbB.empty()) {
                    lastExpr_ = fnSbB + "((void*)" + sbHwnd + ")";
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
            const IfaceView* ivV221 = ivLookupIface(ivIface);
            const std::string ivSlot221 =
                ivV221 ? ivSlotKeyForMember(*ivV221, memberName) : std::string();
            if (!ivSlot221.empty()) {
                std::vector<std::string> ivArgs;
                for (size_t i = 0; i < node.positional.size(); i++) {
                    emitExpr(*node.positional[i]);
                    ivArgs.push_back(lastExpr_);
                }
                for (auto& named : node.named) {
                    emitExpr(*named.value);
                    ivArgs.push_back(lastExpr_);
                }
                const std::string ivCall = ivrefVtableCall(ivIface, memberName, objExpr, ivArgs);
                if (!ivCall.empty()) {
                    lastExpr_ = ivCall;
                    return;
                }
            }
            // 接口里没有这个成员: 契约检查已报过错, 落回后续路径不再级联
        }

        // P6.4: 接口引用方法调用 (Dim x As IFoo → x.Method → x.vtbl->Method(x.obj, args))
        // 账 #305 新账 #471 (第一百七十五刀): 那一发的形状交给唯一主人 `legacyViewCallExpr`（语句路
        // 也问它）。**换主人不换形状**：拼法逐字照旧 ⇒ 存量产物应当一字不动（清单零移动是本刀的护栏）。
        // 这里先问一次「是不是视图」是**必须的**：实参的 emitExpr 有副作用，不能白发了再退。
        // `isEarlyBoundCom_` 那道闸留在本站 —— 它说的是「这条表达式路上接收者已被前期绑定接管」，
        // 是本路的上下文，不是名册那一问的一部分。
        {
            std::string objLower = objExpr;
            std::transform(objLower.begin(), objLower.end(), objLower.begin(), ::tolower);
            if (slotIsLegacyIfaceView(objLower) && !isEarlyBoundCom_) {
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
                const std::string p64Call = legacyViewCallExpr(objExpr, memberName, callArgs);
                if (!p64Call.empty()) {
                    lastExpr_ = p64Call;
                    return;
                }
            }
        }

        // P6.3: 前期绑定 — 利用类型签名确定返回类型，统一走后期绑定(IDispatch)
        // 原因: vtable直接调用的函数签名可能不是VARIANT* (如get_Count用long*)，
        // vb6_ComVtableGet*辅助函数统一用VARIANT*签名会导致调用错误
        if (isEarlyBoundCom_ && earlyBoundSym_) {
            isEarlyBoundCom_ = false;
            const Symbol* comSym = earlyBoundSym_;
            earlyBoundSym_ = nullptr;

            // 账 #245/§B94: 这张表只从 comSigViewOf 读一次，下面两处共用同一份视图。
            ComSigView sv = comSigViewOf(comSym, memberName);
            bool hasArgs = (!node.positional.empty() || !node.named.empty());


            // 无参属性Get: 用后期绑定属性读取，根据返回类型选函数
            if (sv.found && sv.isGet && !hasArgs) {
                lastExpr_ = comGetterExpr(sv, objExpr, memberName);
                return;
            }
            // P24-07: 有参数的前期绑定方法 → 利用签名选择类型化COM调用函数
            if (sv.found && hasArgs) {
                const Symbol::ComMethodSig& sig = *sv.sig;
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
                        packedArgs.push_back(packArgValue(packFn, lastExpr_));
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
                        packedArgs.push_back(packArgValue(packFn, lastExpr_));
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
                lastExpr_ = comTypedCallExpr(sv, callArgs);
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
                // Fix 266: 打包器的**型**与 marker 解析的**型**不同源 ——
                // comPackExpr 按 AST 推的给出 typed 打包器 (ComPackDouble/Int/Bool/BSTR),
                // 而 marker 解析可能交回**无型** VARIANT 取值
                // `vb6_VariantFromComResult(vb6_ComGetProp(obj, L"m"))` ⇒
                // `vb6_ComPackDouble(vb6_VARIANT)` = C2440 "无法从 vb6_VARIANT 转换为 double"。
                // 这一格收口到 packArgValue 一处（见下方 packedArgs.push_back 那一行；
                // 与 callee_ident / callee_member / prelude 三个片段共用同一个函数，
                // 不在每个站点各抄一份 —— 实测只堵一处等于没堵）。
                // Fix <vbeclipse>: 实参表达式本身已返回**堆 VARIANT*** (ComGetProp /
                // ComCall / ComGetObjectProp … 返回 void* 承载一个 COM VARIANT) 时, 它
                // 已经是 ComCall 形参数组要的东西, 不能再包一层 vb6_ComPackObject —— 那会
                // 把这个 VARIANT* 当 IDispatch*, 读其头 4 字节 (VT_EMPTY 时 vt+reserved=0)
                // 为 lpVtbl=NULL, 再 NULL+4 取 AddRef 槽 → av read target=0x4。
                // (ucPerspective 停靠: ucTab.Add 的 Icon 实参 = View.Icon 晚绑定读数。)
                bool argIsVariantPtr =
                       lastExpr_.compare(0, 15, "vb6_ComGetProp(") == 0
                    || lastExpr_.compare(0, 17, "vb6_ComGetPropArg(") == 0
                    || lastExpr_.compare(0, 12, "vb6_ComCall(") == 0
                    || lastExpr_.compare(0, 20, "vb6_ComCallByDispid(") == 0
                    || lastExpr_.compare(0, 18, "vb6_ComCallObject(") == 0
                    || lastExpr_.compare(0, 18, "vb6_ComGetObjectProp(") == 0
                    || lastExpr_.compare(0, 23, "vb6_ComGetPropertyObject(") == 0
                    || lastExpr_.compare(0, 23, "vb6_ComGetPropertyVariant(") == 0;
                if (argIsVariantPtr && packFn == "vb6_ComPackObject") {
                    packedArgs.push_back(lastExpr_);   // 已是 VARIANT*, 直接作实参 (ComCall 会 VariantClear+free)
                } else {
                    packedArgs.push_back(packArgValue(packFn, lastExpr_));
                }
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
    if (isVariantSlotElemRead(callee) && !node.positional.empty()) {
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
    //
    // Fix <Command2>: 但"裸函数名 + this"这一形态**必须**先拼回完整调用, 否则下面
    // P6.5 的拆分与 Fix 242b 的「项目类集合默认成员索引」识别整段被跳过 ——
    // 它们都以 `callee` 形如 `func(obj)` 为前提。UserControl 实例的成员
    // (emitUcInstanceMemberExpr 的 asCallCallee_ 分支) 交的正是裸名:
    //   `MyUC1.Things("abc").DoIt`
    //     → 裸名 vb6_MyUC_prop_get_Things + pendingChainObj_ = (vb6_cls_MyUC*)…
    // 拼不回去 ⇒ `("abc")` 落进通用位置实参路, 又被 Fix 041c 按「Things 是 0 参
    // PropertyGet」截掉 (args.size()=1 > calleeParams.size()=0 ⇒ maxArgs=0)
    // ⇒ 实参静默消失, 接收者变成整个集合并引发运行期错误 9。
    // 拼回去之后 callee 形如 `func(obj)`, 下面的 split/242b 与普通类字段/属性
    // 走**同一条**路径 (实测 ComCtlsDemo ImageList.ListImages("abc").Draw 族)。
    // ⚠ 只在**下面的 P6.5 拆分确实会跑**时才拼: 拆分的门槛是
    //   `callee 形如 func(obj)` && (有位置/named 实参 || needsSplitForOptionalPad
    //   || needsSplitForMissingArgs)
    // 无实参又不需要补齐时 (如 `s(0).Side()` —— Side 无参) 拆分整段不进, 此时若已把
    // classMethodObjArg 清掉, 末尾 `callee + "(" + argList + ")"` 就会在**已完整的调用**
    // 后面再拼一层空括号 ⇒ `vb6_X_Side(recv)()` ⇒ MSVC C2064 (term does not evaluate
    // to a function taking N arguments)。tests/arr_cls 实测 (Fix 192 那一族)。
    if (!pendingChainObj_.empty()) {
        classMethodObjArg = std::move(pendingChainObj_);
        pendingChainObj_.clear();
        const bool splitWillRun = !classMethodObjArg.empty() && !callee.empty()
            && callee.find('(') == std::string::npos
            && ((!node.positional.empty() || !node.named.empty())
                || needsSplitForOptionalPad || needsSplitForMissingArgs);
        if (splitWillRun) {
            callee += "(" + classMethodObjArg + ")";
            classMethodObjArg.clear();
        }
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
            // Fix 242b: 内层无参 `_prop_get_` 的返回类型是**项目类集合**时, 外层 (idx) 是该
            // 集合默认成员 Item 的索引 —— 记下项目类 Item 的发射函数, 下方改走项目类路径,
            // 而不是把它当 IDispatch (项目类实例不是 COM 对象)。见使用点注释。
            std::string collItemFn242b;
            std::string collCls242b;
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
                        // Fix <Command2>: `_prop_get_` 必须按 **funcPart 里的类前缀**解析
                        // (`vb6_<Cls>_prop_get_<Prop>`)。原实现只用 lookupModuleByKind
                        // 在当前模块作用域里按**名字**找 PropertyGet —— 工程内多个类有同名
                        // 属性 (Value / Text / Caption / Picture …) 时会命中**别的类**那枚。
                        // 实测 ComCtlsDemo MainForm.frm:598
                        //   `HotKey1.Value(vbCtrlMask + vbAltMask) = vbKeyA`
                        // 撞上一枚 0 参的 Value ⇒ 被误判成"集合默认成员索引"，发成
                        //   vb6_VariantFromComResult(vb6_ComCall(<HotKey.Value>, L"Item", …)) = …
                        // ⇒ C2198 (实参太多) + C2106 (不是左值)，工程直接编不过。
                        // 判据与 resolveClassMemberCall / getClassMethodReturnType 同源:
                        // 按 sourceModule (或本模块名) 匹配类，名字只作次要条件。
                        std::string clsPrefix86;
                        if (gp > 4) clsPrefix86 = funcPart.substr(4, gp - 4);
                        Symbol* propSym86 = nullptr;
                        if (!clsPrefix86.empty() && symTab_.moduleScope()) {
                            const std::string pLower86 = Symbol::toLower(propName);
                            const std::string cLower86 = Symbol::toLower(clsPrefix86);
                            for (const auto& [k86, s86] : symTab_.moduleScope()->symbols()) {
                                if (s86->lowerName != pLower86) continue;
                                if (s86->kind != SymbolKind::PropertyGet) continue;
                                bool m86 = false;
                                if (s86->isExternal) {
                                    if (Symbol::toLower(s86->sourceModule) == cLower86) m86 = true;
                                } else if (isClassModule_
                                           && Symbol::toLower(moduleName_) == cLower86) {
                                    m86 = true;
                                }
                                if (m86) { propSym86 = s86.get(); break; }
                            }
                        }
                        if (!propSym86 && clsPrefix86.empty()) {
                            propSym86 = symTab_.lookupModuleByKind(propName,
                                                                   SymbolKind::PropertyGet);
                        }
                        if (propSym86 && propSym86->params.empty()) {
                            innerIsChainedObj = true;
                            // Fix 242b: 该无参属性返回**项目类集合**时记下项目类 Item 发射
                            // 函数 (下方走项目类路径)。判据用 getClassMethodReturnType(返回名,
                            // "Item") 非空 = 返回名前确实是一个项目 Class 且它有 Item 成员。
                            if (!propSym86->variableTypeName.empty()) {
                                const std::string retName242b = propSym86->variableTypeName;
                                if (!getClassMethodReturnType(retName242b, "Item").empty()) {
                                    const std::string itemFn242b =
                                        resolveClassMemberCall(retName242b, "Item");
                                    if (!itemFn242b.empty()) {
                                        collItemFn242b = itemFn242b;
                                        collCls242b = retName242b;
                                    }
                                }
                            }
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
                // Fix 242b: 项目类集合索引 —— 内层 `_prop_get_` 返回项目类 (C 侧
                // vb6_cls_<Coll>*) 时, 外层 (idx) 是该项目类集合默认成员 Item 的索引,
                // 必须发 vb6_<Coll>_prop_get_Item(<coll>, <idx>) 直接拿元素类指针。
                //
                // 此前一律发 vb6_VariantFromComResult(vb6_ComCall(<coll>, L"Item", …))
                // —— 把**项目类实例**当 IDispatch: C 里 vb6_cls_X* 隐式转 void* 编得过,
                // 但结果是 VARIANT, 下游按元素类做成员访问时 `(void*)VARIANT` ⇒ C2440。
                // 实测 ComCtlsDemo (Fix 242 同时把元素类推断接上之后暴露):
                //   CoolBar.ctl:874 `Me.Bands(i).Caption`
                //     Bands As CbrBands (项目类) → CbrBands.Item(i) As CbrBand
                //     ⇒ C2440 vb6_VARIANT → void* ×157
                //   ToolBar.ctl  `Buttons(i).ButtonMenus(ii).Key` 族
                //     ⇒ C2440 vb6_VARIANT → vb6_cls_TbrButton* ×25
                // 与 Pattern L (cgen_expr_call_callee_ident/member.inc 的"项目类字段
                // 集合索引"发 vb6_<Type>_prop_get_Item) 同口径, 只是这里集合来自**属性**。
                if (!collItemFn242b.empty()) {
                    std::vector<ParameterInfo> collParams242b;
                    findClassMemberGetParams(collCls242b, "Item", collParams242b);
                    // 形参个数必须与外层索引个数一致, 否则是别的形态 → 退回 COM。
                    if (collParams242b.size() == node.positional.size()) {
                        std::string args242b;
                        for (size_t i = 0; i < node.positional.size(); i++) {
                            emitExpr(*node.positional[i]);
                            std::string one242b = std::move(lastExpr_);
                            if (collParams242b[i].type == Vb6Type::Variant) {
                                one242b = boxToVariant(node.positional[i].get(), one242b);
                            }
                            if (i) args242b += ", ";
                            args242b += one242b;
                        }
                        lastExpr_ = collItemFn242b + "(" + innerObj + ", " + args242b + ")";
                        return;
                    }
                }
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
                    packedArgs86.push_back(packArgValue(packFn86, lastExpr_));
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