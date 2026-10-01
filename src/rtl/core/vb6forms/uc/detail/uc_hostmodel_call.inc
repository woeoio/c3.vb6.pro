// vb6forms_uc 拆分片段 uc_hostmodel_call (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 1424~1493 行，逐行未改
// 在函数体内 #include；路径一律用 basename

    obj = vb6_UC_UnwrapHost(obj);  // czUI fix: 包装器透明解包
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    (void)argc;
    if (!obj || !name) return 0;
    vb6_ho_setVariantEmpty(out);
    if (vb6_uc_isFont(obj)) return 1;
    if (vb6_uc_isColl(obj)) {   // Fix 112c: Collection 方法
        if (_wcsicmp(name, L"Add") == 0 && argc >= 1) {
            BSTR k = (argc >= 2) ? vb6_ho_variantToBstr((vb6_VARIANT*)argv[1]) : NULL;
            /* Fix 134: VB6 Add(Item[, Key, Before, After]) 的位置插入 */
            if (argc >= 3 && argv[2]) {
                vb6_VARIANT* b = (vb6_VARIANT*)argv[2];
                if (b->vt != vb6_vtEmpty) {
                    int32_t pos = vb6_ho_variantToLong(b);
                    if (k) { vb6_Collection_AddKeyedAt(obj, argv[0], k, pos); }
                    else   { vb6_Collection_AddAt(obj, argv[0], pos); }
                    return 1;
                }
            }
            if (argc >= 4 && argv[3]) {
                vb6_VARIANT* a = (vb6_VARIANT*)argv[3];
                if (a->vt != vb6_vtEmpty) {
                    int32_t pos = vb6_ho_variantToLong(a) + 1;
                    if (k) { vb6_Collection_AddKeyedAt(obj, argv[0], k, pos); }
                    else   { vb6_Collection_AddAt(obj, argv[0], pos); }
                    return 1;
                }
            }
            if (k) vb6_Collection_AddKeyed(obj, argv[0], k);
            else   vb6_Collection_Add(obj, argv[0]);
        } else if (_wcsicmp(name, L"Item") == 0 && argc >= 1) {
            vb6_VARIANT* a = (vb6_VARIANT*)argv[0];
            BSTR k = vb6_ho_variantToBstr(a);
            if (k) vb6_Collection_ItemByKey(obj, k, out);
            else   vb6_Collection_Item(obj, vb6_ho_variantToLong(a), out);
        } else if (_wcsicmp(name, L"Remove") == 0 && argc >= 1) {
            vb6_VARIANT* a = (vb6_VARIANT*)argv[0];
            BSTR k = vb6_ho_variantToBstr(a);
            if (k) vb6_Collection_RemoveByKey(obj, k);
            else   vb6_Collection_Remove(obj, vb6_ho_variantToLong(a));
        }
        return 1;
    }
    if (vb6_uc_isControls(obj)) {
        // Fix <vbeclipse> rev14: UserControl.Controls 的真实集合语义。
        // 后端的裸 `Controls.Item/Add/Remove` 会被发射成 vb6_ComCall*(vb6_UC_Controls(), ...),
        // 运行期落到这里。Item/Add 返回**实例指针** (项目类直调需要), 标准控件无实例时
        // 回落 HWND (宿主分派即可应答)。详见 uc_controls.c 的 vb6_UC_ControlsItemObj 注释。
        if (_wcsicmp(name, L"Item") == 0 && argc >= 1) {
            vb6_VARIANT* a = (vb6_VARIANT*)argv[0];
            BSTR key = vb6_ho_variantToBstr(a);
            void* it = key ? vb6_UC_ControlsItemObjByName(obj, key)
                           : vb6_UC_ControlsItemObj(obj, vb6_ho_variantToLong(a));
            vb6_ho_setVariantObject(out, it);
        } else if (_wcsicmp(name, L"Add") == 0 && argc >= 2) {
            // VB6 Controls.Add(progId, name[, left, top, width, height])
            BSTR progId = vb6_ho_variantToBstr((vb6_VARIANT*)argv[0]);
            BSTR cname  = vb6_ho_variantToBstr((vb6_VARIANT*)argv[1]);
            vb6_ho_setVariantObject(out, vb6_UC_ControlsAdd(obj, progId, cname));
        } else if (_wcsicmp(name, L"Remove") == 0 && argc >= 1) {
            vb6_VARIANT* a = (vb6_VARIANT*)argv[0];
            if (a && a->vt == vb6_vtDispatch && a->pdispVal) {
                vb6_UC_ControlsRemove(obj, a->pdispVal, NULL);   // 按对象
            } else {
                vb6_UC_ControlsRemove(obj, NULL, vb6_ho_variantToBstr(a));  // 按名
            }
        } else if (_wcsicmp(name, L"Count") == 0) {
            vb6_ho_setVariantLong(out, vb6_UC_ControlsCount(obj));
        }
        return 1;
    }
    // Fix <vbeclipse> rev18: UC **实例指针**归一成宿主 HWND。
    // 晚绑定调用点 (`Dim x As Variant: Set x = Controls.Item(id)`) 传进来的是
    // `vb6_cls_X*`, 而下面整条内建成员链 (与那道"三条都不匹配就 return 0"的闸门)
    // 只认 HWND ⇒ 以前对 UC 实例的每一个成员调用都被静默丢弃。
    // 实证 play78: ucPerspective.Refresh 的 `l_ucFolder.Move/Refresh` 全落空,
    // 7 个 ucFolder 宿主停在设计期 569x441 重叠。
    { void* h18 = vb6_UC_HwndOf(obj); if (h18) obj = h18; }

    if (!vb6_ho_find(obj) && !vb6_uc_findByHwnd(obj) && !IsWindow((HWND)obj)) return 0;

    if (_wcsicmp(name, L"Refresh") == 0) {
        InvalidateRect((HWND)obj, NULL, TRUE); UpdateWindow((HWND)obj);
        return 1;
    }
    if (_wcsicmp(name, L"Cls") == 0) {
        HDC dc = GetDC((HWND)obj);
        if (dc) { RECT rc; GetClientRect((HWND)obj, &rc);
                  HBRUSH b = CreateSolidBrush(RGB(255,255,255));
                  FillRect(dc, &rc, b); DeleteObject(b); ReleaseDC((HWND)obj, dc); }
        return 1;
    }
    if (_wcsicmp(name, L"SetFocus") == 0) { SetFocus((HWND)obj); return 1; }
    /* Fix 143: 原生列表控件经 COM 晚绑定调用 AddItem/RemoveItem/Clear.
     * VB6 里 ComboBox/ListBox 的 AddItem 是方法, cgen 生成
     * vb6_ComCall(vb6_hwnd_cboThemes, L"AddItem", args, 1), 而 HWND 不是 IDispatch.
     * 此前这些名字落进 line "未知方法一律空实现" → 调用被静默丢弃:
     * NewTab 示例 frmMain 的 `cboThemes.AddItem iTheme.Name` 一条都没进来,
     * 主题下拉框只剩创建时那 1 个空条目. */
    if (_wcsicmp(name, L"AddItem") == 0 && argc >= 1) {
        BSTR s = vb6_ho_variantToBstr((vb6_VARIANT*)argv[0]);
        if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
            fprintf(stderr, "[C3_HOST] AddItem hwnd=%p text='%ls'\n", obj, s ? s : L"(null)");
        vb6_AddItem(obj, s);
        return 1;
    }
    if (_wcsicmp(name, L"RemoveItem") == 0 && argc >= 1) {
        vb6_RemoveItem(obj, vb6_ho_variantToLong((vb6_VARIANT*)argv[0]));
        return 1;
    }
    if (_wcsicmp(name, L"Clear") == 0) { vb6_ClearList(obj); return 1; }
    /* Fix <vbeclipse> rev18: `Move` 以前和 ZOrder/Print/Line 一起被当空实现丢掉 ——
     * 而 VB6 里 UserControl/控件的 Move 是**真摆位**, 停靠布局完全靠它:
     *   ucPerspective.Refresh 的 `l_ucFolder.Move .LeftPos, .TopPos,
     *   .RightPos - .LeftPos, .BottomPos - .TopPos`。
     * 丢掉 ⇒ 7 个 ucFolder 宿主永远停在设计期尺寸重叠在左上角。
     * 单位是**缇** (调用点发的是 Folder/ucFolder 的位置计算值, 与
     * vb6_UserControl_ScaleWidth 同族), 换算复用 vb6_ControlMove (它内部走
     * vb6_TwipToX/Y, 且对 WS_CHILD 会先把屏幕坐标换回父客户区)。
     * VB6 允许 Move 只给 left/top (宽高不变) 或给 3 个 (高度不变)。 */
    if (_wcsicmp(name, L"Move") == 0) {
        double L18 = 0, T18 = 0, W18 = 0, H18 = 0;
        int mask18 = 0;
        if (argc >= 1 && argv[0]) { L18 = vb6_VariantToDouble(*(vb6_VARIANT*)argv[0]); mask18 |= 1; }
        if (argc >= 2 && argv[1]) { T18 = vb6_VariantToDouble(*(vb6_VARIANT*)argv[1]); mask18 |= 2; }
        if (argc >= 3 && argv[2]) { W18 = vb6_VariantToDouble(*(vb6_VARIANT*)argv[2]); mask18 |= 4; }
        if (argc >= 4 && argv[3]) { H18 = vb6_VariantToDouble(*(vb6_VARIANT*)argv[3]); mask18 |= 8; }
        if (mask18) vb6_ControlMove(obj, L18, T18, W18, H18, mask18);
        return 1;
    }
    if (_wcsicmp(name, L"Show") == 0) { ShowWindow((HWND)obj, SW_SHOW); return 1; }
    if (_wcsicmp(name, L"Hide") == 0) { ShowWindow((HWND)obj, SW_HIDE); return 1; }
    if (_wcsicmp(name, L"ZOrder") == 0 ||
        _wcsicmp(name, L"Print") == 0 || _wcsicmp(name, L"Line") == 0) {
        return 1;
    }
    return 1;   // 宿主对象: 未知方法一律空实现
