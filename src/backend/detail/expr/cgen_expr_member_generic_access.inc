// cgen_expr_member_generic_access.inc - 函数体片段：通用成员访问 + 链式 COM 检测 / Variant 数组元素 / UDT 字段链
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 944~1063 行，逐行未改。

    // 通用成员访问 (结构体字段 / 链式COM访问)
    {
        // 链中对象求值旗标 (见 cgen_state.inc inMemberObjectEmit_): 自类型过程
        // 作接收者时, 子分支据它把 comMemberName_ 留空。
        const bool savedInObjEmit = inMemberObjectEmit_;
        inMemberObjectEmit_ = true;
        emitExpr(*node.object);
        inMemberObjectEmit_ = savedInObjEmit;
    }

    // 链式COM调用检测: 如果object求值产生COM标记, 先resolve为对象值
    // (fso.GetFolder("x") → vb6_ComCallObject → 返回IDispatch*)
    // 然后在新对象上访问成员 → 设置新的COM标记
    if (isComMarker_) {
        if (std::getenv("C3_DBG_ST")) {
            std::cerr << "[ST-GEN] marker obj=" << comObjExpr_
                      << " member=" << comMemberName_ << std::endl;
        }
        // object是COM属性访问, 解析为对象值
        resolveComValue("Object");
        // 检查resolveComValue后的结果是否是对象表达式 (ComCallObject/ComGetObjectProp)
        // 如果是, 说明这是一个链式COM对象访问, 设置新的COM标记
        std::string objExpr = lastExpr_;
        // 检测是否是COM对象表达式 (由ComCallObject/ComGetObjectProp返回的void*)
        // 这些都是void*类型, 可以作为COM对象继续访问成员
        if (objExpr.find("vb6_ComCallObject(") == 0 ||
            objExpr.find("vb6_ComGetObjectProp(") == 0 ||
            objExpr.find("vb6_CreateObject(") == 0 ||
            // Fix <vbeclipse>: 外部窗体控件句柄访问器 vb6_form_ctrl_<Form>_<Ctrl>() —
            // 返回 HWND, 其上的集合成员 (.Nodes 等) 由 resolveComValue 的
            // externalTreeViewAccs_ 分支认领; 其余成员走晚绑定 (运行期受限)。
            objExpr.find("vb6_form_ctrl_") == 0 ||
            // C29-3: 原生复刻的控件成员集合 (ImageList.ListImages 那一族) 返回的也是
            // 真 IDispatch —— 必须当 COM 对象继续访问, 否则拼出 `X.Count` (X 是 void*
            // 表达式) → 直接是 C 语法错。
            objExpr.find("vb6_ImageList_ListImages(") == 0 ||
            // C29-7: ListView 的两个集合同理 —— 都是真 IDispatch, 成员与方法要走晚绑定。
            objExpr.find("vb6_ListView_ListItems(") == 0 ||
            objExpr.find("vb6_ListView_ColumnHeaders(") == 0 ||
            // C29-OLE: `OLE1.Object` 返回嵌入对象的 IDispatch* —— 链 `.Something` 走晚绑定。
            objExpr.find("vb6_OleCon_GetObject(") == 0 ||
            // C29-Data: `Data1.Recordset` 的真 IDispatch (memberobj 的 RS 族)。
            objExpr.find("vb6_Data_RecordsetObj(") == 0 ||
            // C29-8b: `TreeView1.Nodes` 的真 IDispatch 集合同理。
            objExpr.find("vb6_TreeView_Nodes(") == 0 ||
            // C29-5b: `Toolbar1.Buttons` 的真 IDispatch 集合同理。
            objExpr.find("vb6_Toolbar_Buttons(") == 0 ||
            // 账 #299: `StatusBar1.Panels` 的真 IDispatch 集合同理 —— 这一族的名单在三处
        // (这里两处 + cgen_state.inc 的 isControlCollectionExpr)，漏一处就拼出
        // `vb6_StatusBar_Panels(…).Count` → cl C2224。哨兵 check_statusbar_panels_authority 钉三处同值。
            objExpr.find("vb6_StatusBar_Panels(") == 0) {
            // 链式COM: 设置COM标记, objExpr是中间对象表达式
            comObjExpr_ = objExpr;
            comMemberName_ = node.memberName;
            isComMarker_ = true;
            lastExpr_ = objExpr;  // 保持对象表达式
            if (std::getenv("C3_DBG_ST")) std::cerr << "[ST-GEN2] chain marker " << node.memberName << std::endl;
            return;
        }
        if (std::getenv("C3_DBG_ST")) std::cerr << "[ST-GEN-TAIL] struct text " << node.memberName << std::endl;
        // 账 (Outlook2013 frmShortcutBar:202): 自类型过程调用作链中接收者
        // (`vb6_frmMain_RibbonBar().FindTab(...)`) — resolveComValue 的空成员守卫
        // 交回的调用文本以 "()" 收尾。vb6_ 前缀 + "()" 收尾的调用当真 IDispatch
        // 接收者续链 (晚绑定), 否则落 `obj + "." + member` 结构体文本 (C2224)。
        {
            const bool endsParens79 = objExpr.size() > 5
                && objExpr.compare(0, 4, "vb6_") == 0
                && objExpr.compare(objExpr.size() - 2, 2, "()") == 0;
            if (std::getenv("C3_DBG_ST")) {
                std::string last279 = objExpr.size() >= 2 ? objExpr.substr(objExpr.size() - 2) : objExpr;
                std::cerr << "[ST-RULE] ends=" << endsParens79 << " sz=" << objExpr.size()
                          << " last2='" << last279 << "' objExpr='" << objExpr << "'" << std::endl;
            }
            if (endsParens79) {
                comObjExpr_ = objExpr;
                comMemberName_ = node.memberName;
                isComMarker_ = true;
                lastExpr_ = objExpr;
                return;
            }
        }
        // 非COM对象值, 按结构体字段处理
        lastExpr_ = objExpr + "." + cIdent(node.memberName);
        return;
    }

    // 无标记态的链式接收者: 对象表达式整枚是 **COM 调用结果** (vb6_ComCallObject /
    // ComGetObjectProp / ComCall — 旗标在 IndexOrCallExpr 消费时已复位)。
    // `...FindTab(12000).Visible = 0` 的 `.Visible` 一站走这里 —— 此前落
    // `obj + "." + member` 结构体文本 (C2224)。按真 IDispatch 接收者立标记续链。
    {
        const std::string objRes = lastExpr_;
        if (objRes.compare(0, 19, "vb6_ComCallObject(") == 0
            || objRes.compare(0, 21, "vb6_ComGetObjectProp(") == 0
            || objRes.compare(0, 12, "vb6_ComCall(") == 0) {
            // 接收者必须是**真 IDispatch***: `vb6_ComCall` 交回的是堆上 VARIANT*,
            // 直接喂给下一站 (`vb6_ComGetStringProp(<VARIANT*>, L"Name")`) 就是
            // 拿 VARIANT 盒当对象 —— 运行期 C3_LATEBIND_MISS kind=notdispatchable,
            // 属性读交空 (test_com2 `fso.GetFolder("C:\Users").Name` 实测)。
            // comObjectRefFromCallExpr 正是「这一发要当对象用 ⇒ 换到交回对象的那一枚
            // 出口」的唯一主人 (ComCall → ComCallObject, ComGetProp → ComUnpackObject)。
            comObjExpr_ = comObjectRefFromCallExpr(objRes);
            comMemberName_ = node.memberName;
            isComMarker_ = true;
            isEarlyBoundCom_ = false;
            earlyBoundSym_ = nullptr;
            lastExpr_ = comObjExpr_;
            return;
        }
    }

    std::string obj = std::move(lastExpr_);

    // 账 #257 第十四层: 对象表达式若是以「裸函数名 + pendingChainObj_」协议交付的,
    // 这里必须**把 this 补回去**。
    //
    // 协议本身没问题 (emitUcInstanceMemberExpr / class_fallback / voidptr_com 的
    // asCallCallee_ 分支都按它交付: lastExpr_ = 裸函数名, pendingChainObj_ = this 实参),
    // 问题在**谁算消费者**: 上面 `emitExpr(*node.object)` 发射的是**对象**, 是值上下文;
    // 而 `asCallCallee_` 是从外层 CallStmt/IndexOrCallExpr 一路带下来的, 于是对象也被
    // 按 callee 交付了裸名。此后 obj 便是"一个函数名"而不是值 ——
    // 下游 class_fallback 的 `thisArg088c = "(void*)" + obj` / generic_access 自己的
    // `"(void*)" + objExpr` 都把它当对象交出去 ⇒ **接收者丢失**, 编译期完全看不见。
    //
    // 实测 ComCtlsDemo MainForm.frm:474 `ListView1.ColumnHeaders.Add , , "Col1"`:
    //   内层 `ListView1.ColumnHeaders` 交 裸名 vb6_ListView_prop_get_ColumnHeaders
    //   + pendingChainObj_ = (vb6_cls_ListView*)vb6_UC_InstanceOf(vb6_hwnd_ListView1);
    //   外层（本节点, member="Add"）经 Fix 088c 认出接收者类 LvwColumnHeaders,
    //   发 `vb6_LvwColumnHeaders_Add((void*)vb6_ListView_prop_get_ColumnHeaders, ...)`
    //   ⇒ `me` 是**代码地址** 0x4BB2A0, `me->ShadowObjectPointer`([me+4]) 读到的是
    //   指令字节 0x45C724EC ⇒ objset 的 IsRealComObject 当场 AV
    //   (av read target=0x45c724ec, 与反汇编 [0x4BB2A4..0x4BB2A7] 逐字节吻合)。
    //   同形共 76 处 (ListView.ListItems/ColumnHeaders/Groups/SelectedItem,
    //   TreeView.Nodes, ImageCombo.ComboItems, ToolBar.ButtonMenus,
    //   ImageList.ListImages, LvwListItem.ListSubItems …)。
    //
    // 与 Fix 261 (实参位同样要清 asCallCallee_/消费协议) 是同一条纪律的两个位置。
    if (!pendingChainObj_.empty()) {
        obj += "(" + pendingChainObj_ + ")";
        pendingChainObj_.clear();
    }

    // Fix 090af: Variant 数组元素 (VB6_SA_AT(vb6_VARIANT, arr, i)) 的对象成员访问.
    // VB6: Dim d() As Variant ... If TypeOf d(0) Is cJson Then Set d(0) = d(0).Root
    // d(0) 持有对象时 .member 是后期绑定调用 (同 knownVariantVars_ 标识符路径
    // P24-04 1753). 此前落入通用 fallback 生成 VB6_SA_AT(...).member 结构体字段
    // 访问 → C2039 (vb6_VARIANT 无该成员) + 级联 C2198/C2440.
    if (obj.find("VB6_SA_AT(vb6_VARIANT,") == 0) {
        std::string varRef = "&(" + obj + ")";
        comObjExpr_ = "vb6_VariantToObject(" + varRef + ")";
        comMemberName_ = node.memberName;
        isComMarker_ = true;
        lastExpr_ = "vb6_VariantFromComResult(vb6_ComGetProp(" + comObjExpr_
                  + ", L\"" + node.memberName + "\"))";
        return;
    }

    // P20-44: DataObject.Files 的外层 —— `Data.Files.Count`。内层 precheck 给的
    // lastExpr_ 带尾部标记 "/* oleDD-files */", 这里消费它。
    {
        static const char* kFilesMark = " /* oleDD-files */";
        size_t fm = obj.find(kFilesMark);
        if (fm != std::string::npos) {
            std::string ddArg = obj.substr(0, fm);
            std::string memFiles = Symbol::toLower(node.memberName);
            if (memFiles == "count") { lastExpr_ = "vb6_oleDD_GetFileCount(" + ddArg + ")"; return; }
            if (memFiles == "item")  { lastExpr_ = ddArg; return; }  // Files(i) 走 IndexOrCall 拦截
        }
    }

    // 链式COM检测2: object求值结果是COM对象表达式
    // (从IndexOrCallExpr产生的COM调用结果, 是void*类型的IDispatch*)
    // Fix 153: vb6_PropertyPage_SelectedControls(0) 返回 void* (被编辑控件引用),
    // 其上的 .Rows/.Cols/.ClipSeparators 等成员须走后期绑定 COM, 否则生成
    // `vb6_PropertyPage_SelectedControls(0).Rows` → C2224 (void* 无该成员)。
    // Fix 158e: vb6_Common_PtrToObj(expr) (用户 Object 属性转换助手, 返回 void*)
    // 与 vb6_CreateObject(progid) (返回 void* IDispatch*) 同属 COM 对象表达式 —
    // 其上的方法 (.ListImages / .SendKeys) 也须走后期绑定, 否则拼 `.Member` → C2224
    //   VBFlexGrid.c: ptrToObj(...).ListImages(Image).Index   (ObjPointer 恢复 COM 对象)
    //   Common.c:     CreateObject("WScript.Shell").SendKeys Text, Wait, 1
    if (obj.find("vb6_ComCallObject(") == 0 ||
        obj.find("vb6_ComGetObjectProp(") == 0 ||
        obj.find("vb6_ComCall(") == 0 ||
        obj.find("vb6_ComCallByDispid(") == 0 ||
        obj.find("vb6_ComGetProp(") == 0 ||
        // Fix <vbeclipse>: 同上 — 外部窗体控件句柄访问器前缀。
        obj.find("vb6_form_ctrl_") == 0 ||
        obj.find("vb6_Common_PtrToObj(") == 0 ||
        obj.find("vb6_CreateObject(") == 0 ||
        // C29-3: 同上 —— 原生复刻的成员集合 (ImageList.ListImages) 也是 COM 对象。
        obj.find("vb6_ImageList_ListImages(") == 0 ||
        // C29-7: ListView 的 ListItems / ColumnHeaders 同理。
        obj.find("vb6_ListView_ListItems(") == 0 ||
        obj.find("vb6_ListView_ColumnHeaders(") == 0 ||
        // C29-OLE: `OLE1.Object` 同理。
        obj.find("vb6_OleCon_GetObject(") == 0 ||
        // C29-Data: `Data1.Recordset` 的真 IDispatch 同理。
        obj.find("vb6_Data_RecordsetObj(") == 0 ||
        // C29-8b: `TreeView1.Nodes` 的真 IDispatch 集合同理。
        obj.find("vb6_TreeView_Nodes(") == 0 ||
        // C29-5b: `Toolbar1.Buttons` 的真 IDispatch 集合同理。
        obj.find("vb6_Toolbar_Buttons(") == 0 ||
        // 账 #299: `StatusBar1.Panels` 的真 IDispatch 集合同理 —— 这一族的名单在三处
        // (这里两处 + cgen_state.inc 的 isControlCollectionExpr)，漏一处就拼出
        // `vb6_StatusBar_Panels(…).Count` → cl C2224。哨兵 check_statusbar_panels_authority 钉三处同值。
        obj.find("vb6_StatusBar_Panels(") == 0 ||
        obj.find("vb6_PropertyPage_SelectedControls(") == 0 ||
        // 账 #305 第二百六十一刀「甲」: With UserControl 体内 `.Extender` 折出来的那枚
        // Extender **对象地址**。它不是 IDispatch, 但晚绑定出口的 `vb6_UC_IsExtender`
        // 认这个身份并把成员读写岔进 `vb6_UC_ExtenderField` 那条 per-instance 通道
        // (Fix 168 读 / Fix 294 写) ⇒ 这一枚是可以继续链式访问成员的对象表达式;
        // 不认它就拼出 `(&vb6_UserControl_Extender).Left` = C2039。
        obj.find("(&vb6_UserControl_Extender)") == 0) {
        // 对于vb6_ComCall/vb6_ComGetProp, 需要先解封为对象
        // 链式调用: ComCall返回VARIANT*, 需ComCallObject才能拿到IDispatch*
        std::string resolvedObj = obj;
        if (obj.find("vb6_ComCall(") == 0) {
            // vb6_ComCall → vb6_ComCallObject (同一参数, 返回void*而非VARIANT*)
            resolvedObj = "vb6_ComCallObject" + obj.substr(strlen("vb6_ComCall"));
        } else if (obj.find("vb6_ComCallByDispid(") == 0) {
            // vb6_ComCallByDispid → 返回 VARIANT* (void*, 承载 COM VARIANT).
            // 成员访问要先解出对象指针: VariantFromComResult 解引用 → ToObjectVal 取 IDispatch*.
            // (与 .member 链式调用的 P24-10 后期绑定默认成员同族)
            resolvedObj = "vb6_VariantToObjectVal(vb6_VariantFromComResult(" + obj + "))";
        } else if (obj.find("vb6_ComGetProp(") == 0) {
            // vb6_ComGetProp → vb6_ComGetObjectProp
            resolvedObj = "vb6_ComGetObjectProp" + obj.substr(strlen("vb6_ComGetProp"));
        }
        comObjExpr_ = resolvedObj;
        comMemberName_ = node.memberName;
        isComMarker_ = true;
        lastExpr_ = resolvedObj;
        return;
    }

    // Fix 158e: Forms(i).Name — 内置 Forms 集合元素是窗体 HWND void*
    // (Pattern M 发射: vb6_Forms_Item(i)). 窗体不是 IDispatch*, .Name 不能在
    // HWND 上 COM 派发 (运行期 crash); 走 RTL 窗体属性读 vb6_GetControlText
    // (窗体 Name ≈ 标题, Common.bas IsFormLoaded StrComp(Forms(i).Name, FormName))
    // 否则拼 `vb6_Forms_Item(...).Name` → C2224 (void* 无该成员).
    if (obj.find("vb6_Forms_Item(") == 0) {
        std::string nameLower158e = node.memberName;
        std::transform(nameLower158e.begin(), nameLower158e.end(), nameLower158e.begin(), ::tolower);
        if (nameLower158e == "name" || nameLower158e == "caption") {
            lastExpr_ = "vb6_GetControlText(" + obj + ")";
            return;
        }
    }

    // Fix 158d: 宿主伪对象 UserControl 的 COM 集合成员 (ParentControls) 在 RTL
    // 声明为全局 void* (Controls 集合: vb6rtl_userctl.h). 其上的链式成员访问
    // (.Count/.Item 等) 须走后期绑定 COM — 否则 M22 发射裸标识符后, 外层直接拼
    // `vb6_UserControl_ParentControls.Count` (void* 上取字段) → C2224.
    //   VBFlexGrid.ctl: If UserControl.ParentControls.Count = 0 Then (2370/2486)
    //   → vb6_ComGetIntProp((void*)vb6_UserControl_ParentControls, L"Count")
    //   与 Fix 153 (SelectedControls) / Fix 085 (void* UDT 字段) 走同一 marker 通道.
    if (obj == "vb6_UserControl_ParentControls" || obj == "vb6_UserControl_Controls") {
        // Fix <vbeclipse>: UserControl.Controls 同 Fix 158d 通道 (裸 `Controls.Item(k)`
        // / `Controls.Add(progid, name)` — VbEclipse ucPerspective 28 处)。
        // rev14: `vb6_UserControl_Controls` 是恒 NULL 的"编译先行"空桩, 改用 RTL 真实
        // 集合单例 vb6_UC_Controls() (按当前 UC 实例回落 formHwnd); 否则 Controls.Item/Add
        // 恒返回 NULL → ucFolder_AddView(NULL) 崩溃。
        // 账 #257: ParentControls 同样不该退回恒 NULL 的空桩 —— 换真实访问器
        // (父容器的集合), 与 Controls 同一口径。
        std::string ctlObjExpr = obj;
        if (obj == "vb6_UserControl_Controls")            ctlObjExpr = "vb6_UC_Controls()";
        else if (obj == "vb6_UserControl_ParentControls") ctlObjExpr = "vb6_UC_ParentControls()";
        comObjExpr_ = "(void*)" + ctlObjExpr;
        comMemberName_ = node.memberName;
        isComMarker_ = true;
        lastExpr_ = "(void*)" + ctlObjExpr;
        return;
    }

    // Fix 085: UDT 对象字段链的成员访问.
    // 内层 UDT 字段访问 (Fix 031 / Fix 084z-4 / WithMemberExpr / 通用 fallback) 在
    // 检测到字段是对象 (项目类 vb6_cls_X* / Collection·COM void*) 时, 于生成的访问
    // 文本后追加 "/* udt objfield <CType> */" 注释标记. 此处外层 MemberAccessExpr
    // 消费该标记:
    //   - vb6_cls_X* (项目类)  → 类方法/属性/数据字段: vb6_cZipArchive_frCopyCompressed
    //   - void* (Collection/COM)→ COM dispatch marker (同 Fix 023 机制)
    // 例: uFile.SourceArchive.frCopyCompressed(...)  (uFile.SourceArchive As cZipArchive)
    //     → vb6_cZipArchive_frCopyCompressed((void*)(uFile.SourceArchive), ...)
    //   withCtx.LocalCertificates.Item(lIdx)  (As Collection)
    //     → vb6_ComCall...(withCtx->LocalCertificates, L"Item", ...)
    if (obj.find("  /* udt objfield ") != std::string::npos) {
        size_t mkPos = obj.find("  /* udt objfield ");
        std::string objExpr = obj.substr(0, mkPos);
        std::string fldType = obj.substr(mkPos + 18);  // 跳过 "  /* udt objfield "
        size_t endMark = fldType.find(" */");
        if (endMark != std::string::npos) fldType = fldType.substr(0, endMark);
        if (fldType.rfind("vb6_cls_", 0) == 0) {
            // 项目类对象字段 (early bound): obj.Field.Method → vb6_Class_Method((void*)obj)
            // 注意: "vb6_cls_" 为 8 字符, 类名紧随其后 (vb6_cls_cZipArchive* → cZipArchive).
            std::string clsName = fldType.substr(8);
            if (!clsName.empty() && clsName.back() == '*') clsName.pop_back();
            std::string resolvedFn = resolveClassMemberCall(clsName, node.memberName);
            if (!resolvedFn.empty()) {
                std::string thisArg = "(void*)" + objExpr;
                // 类虚表 (ai/022 B08f-1/D35 站点④): 接收者是 `u.h` 这样的 UDT 对象字段，
                // 纯读、可以安全用两次 → 与 ⑤/⑨ 同法按实例派发。B08f-1 之前这一支对
                // 跨模块 UDT 字段根本走不到（类型名在语义层就丢了，发出来的是非法 C），
                // 所以这条派发是随"通路修好"一起才存在的；Let/Set 方向 3.4b 无槽，不改写。
                if (!suppressVirtDispatch_
                    && resolvedFn.find("_prop_let_") == std::string::npos
                    && resolvedFn.find("_prop_set_") == std::string::npos) {
                    const std::string dispatchFnUdt = virtDispatchCallee(
                        clsName, node.memberName, thisArg,
                        /*mustDispatch=*/true, node.loc);
                    if (!dispatchFnUdt.empty()) resolvedFn = dispatchFnUdt;
                }
                if (asCallCallee_) {
                    // 由外层 IndexOrCallExpr 把 thisArg 前置到参数首 (同 Fix 015)
                    pendingChainObj_ = thisArg;
                    lastExpr_ = resolvedFn;
                } else {
                    lastExpr_ = resolvedFn + "(" + thisArg + ")";
                }
                return;
            }
            // 类中无此方法/属性 → 类数据字段 (Public Field): objExpr->field
            // Fix 229: 接收者是 **Variant 载体** (VARIANT 值) 时没有结构体字段 ——
            // `objExpr->field` ⇒ C2039 "field 不是 vb6_VARIANT 的成员"。
            // 口径同 Fix 219/229 的其余三处: 改发后期绑定属性读。
            if (isVariantCarrierExpr(objExpr)) {
                lastExpr_ = lateBoundReadOnVariantCarrier(objExpr, node.memberName);
                return;
            }
            lastExpr_ = objExpr + "->" + cIdent(node.memberName);
            return;
        }
        if (fldType == "void*") {
            // Collection/COM 对象字段 → 下游 IndexOrCallExpr/BinaryExpr 走 COM dispatch
            comObjExpr_ = objExpr;
            comMemberName_ = node.memberName;
            isComMarker_ = true;
            lastExpr_ = objExpr;
            return;
        }
        // 兜底 (嵌套 UDT 等非对象标记不应出现) → 普通字段
        lastExpr_ = objExpr + "." + cIdent(node.memberName);
        return;
    }
