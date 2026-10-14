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
    // 账 #247: 与读侧同一个出口 —— 四档各写自己那一档 (单位 = 所在容器的 ScaleMode,
    // 并把写进去的数存进 VB 侧读数缓存)。以前这一支把四档一起读成缇、再整体 MoveWindow:
    // 写 .Left 会把 Top/Width/Height 也按缇重算一遍推回窗口, 而且谁都不存, 于是同一句
    // `对象.Width = 7222` 走这里读回 7215, 走发码读回 7222。
    if (_wcsicmp(name, L"Left") == 0)   { vb6_SetControlLeft(obj, vb6_ho_variantToLong(v)); return 1; }
    if (_wcsicmp(name, L"Top") == 0)    { vb6_SetControlTop(obj, vb6_ho_variantToLong(v)); return 1; }
    if (_wcsicmp(name, L"Width") == 0)  { vb6_SetControlWidth(obj, vb6_ho_variantToLong(v)); return 1; }
    if (_wcsicmp(name, L"Height") == 0) { vb6_SetControlHeight(obj, vb6_ho_variantToLong(v)); return 1; }
    // Fix <vbeclipse> rev18: 内建成员都没命中 → 写**本 UserControl 的自有 Property**
    // (ucFolder.LeftPos = … 这类)。注意必须放在内建档之后: `Visible`/`Caption` 等同名
    // 内建属性仍是窗口层语义。
    if (vb6_UC_OwnPropSet(obj, name, v)) return 1;

    return 1;   // 宿主对象: 未知属性写入一律忽略
