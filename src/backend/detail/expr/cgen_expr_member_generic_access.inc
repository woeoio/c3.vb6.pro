// cgen_expr_member_generic_access.inc - 函数体片段：通用成员访问 + 链式 COM 检测 / Variant 数组元素 / UDT 字段链
// 由 src/backend/expr/cgen_expr_member.cpp 拆出（2026-09-17），在函数体内被 #include，非独立编译单元。
// 内容 = 原文件第 944~1063 行，逐行未改。

    // 通用成员访问 (结构体字段 / 链式COM访问)
    emitExpr(*node.object);

    // 链式COM调用检测: 如果object求值产生COM标记, 先resolve为对象值
    // (fso.GetFolder("x") → vb6_ComCallObject → 返回IDispatch*)
    // 然后在新对象上访问成员 → 设置新的COM标记
    if (isComMarker_) {
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
            // C29-Data: Fields("x") 的求值结果 (FieldValueStr) 上再 .Value —— 透传。
            objExpr.find("vb6_Data_FieldValueStr(") == 0) {
            // 链式COM: 设置COM标记, objExpr是中间对象表达式
            comObjExpr_ = objExpr;
            comMemberName_ = node.memberName;
            isComMarker_ = true;
            lastExpr_ = objExpr;  // 保持对象表达式
            return;
        }
        // 非COM对象值, 按结构体字段处理
        lastExpr_ = objExpr + "." + cIdent(node.memberName);
        return;
    }

    std::string obj = std::move(lastExpr_);

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
        // C29-Data: FieldValueStr 上 .Value 同理。
        obj.find("vb6_Data_FieldValueStr(") == 0 ||
        obj.find("vb6_PropertyPage_SelectedControls(") == 0) {
        // 对于vb6_ComCall/vb6_ComGetProp, 需要先解封为对象
        // 链式调用: ComCall返回VARIANT*, 需ComCallObject才能拿到IDispatch*
        std::string resolvedObj = obj;
        if (obj.find("vb6_ComCall(") == 0) {
            // vb6_ComCall → vb6_ComCallObject (同一参数, 返回void*而非VARIANT*)
            resolvedObj = "vb6_ComCallObject" + obj.substr(strlen("vb6_ComCall"));
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
        std::string ctlObjExpr = (obj == "vb6_UserControl_Controls")
                                     ? "vb6_UC_Controls()" : obj;
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
