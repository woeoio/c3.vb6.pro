// vb6forms_uc 拆分片段 uc_hostmodel_setprop (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 1392~1419 行，逐行未改
// 在函数体内 #include；路径一律用 basename

    obj = vb6_UC_UnwrapHost(obj);  // czUI fix: 包装器透明解包
    const vb6_VARIANT* v = (const vb6_VARIANT*)inV;
    if (!obj || !name) return 0;
    if (vb6_uc_isFont(obj)) return vb6_ho_putFontMember(obj, name, v);
    if (vb6_uc_isControls(obj)) return 1;
    if (vb6_uc_isColl(obj)) return 1;   // Fix 112c: Collection 只读容器

    // Fix <vbeclipse> rev18: UC 实例指针 → 宿主 HWND (与 getprop / call 同口径)
    { void* h18 = vb6_UC_HwndOf(obj); if (h18) obj = h18; }

    if (!vb6_ho_find(obj) && !vb6_uc_findByHwnd(obj) && !IsWindow((HWND)obj)) return 0;

    if (_wcsicmp(name, L"Caption") == 0 || _wcsicmp(name, L"Text") == 0) {
        BSTR s = vb6_ho_variantToBstr(v);
        if (s) SetWindowTextW((HWND)obj, s);
        InvalidateRect((HWND)obj, NULL, TRUE);
        return 1;
    }
    if (_wcsicmp(name, L"Visible") == 0) { ShowWindow((HWND)obj, vb6_ho_variantToLong(v) ? SW_SHOW : SW_HIDE); return 1; }
    if (_wcsicmp(name, L"Enabled") == 0) { EnableWindow((HWND)obj, vb6_ho_variantToLong(v) ? TRUE : FALSE); return 1; }
    if (_wcsicmp(name, L"Left") == 0 || _wcsicmp(name, L"Top") == 0 ||
        _wcsicmp(name, L"Width") == 0 || _wcsicmp(name, L"Height") == 0) {
        int32_t l = 0, t = 0, w = 0, hh = 0;
        vb6_ho_ctrlRect(obj, &l, &t, &w, &hh);
        if (_wcsicmp(name, L"Left") == 0)   l = vb6_ho_variantToLong(v);
        if (_wcsicmp(name, L"Top") == 0)    t = vb6_ho_variantToLong(v);
        if (_wcsicmp(name, L"Width") == 0)  w = vb6_ho_variantToLong(v);
        if (_wcsicmp(name, L"Height") == 0) hh = vb6_ho_variantToLong(v);
        MoveWindow((HWND)obj, vb6_TwipToX(l), vb6_TwipToY(t), vb6_TwipToX(w), vb6_TwipToY(hh), TRUE);
        return 1;
    }
    // Fix <vbeclipse> rev18: 内建成员都没命中 → 写**本 UserControl 的自有 Property**
    // (ucFolder.LeftPos = … 这类)。注意必须放在内建档之后: `Visible`/`Caption` 等同名
    // 内建属性仍是窗口层语义。
    if (vb6_UC_OwnPropSet(obj, name, v)) return 1;

    return 1;   // 宿主对象: 未知属性写入一律忽略
