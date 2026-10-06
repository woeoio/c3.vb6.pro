// vb6forms_uc 拆分片段 uc_hostmodel_getprop (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 1304~1387 行，逐行未改
// 在函数体内 #include；路径一律用 basename

    obj = vb6_UC_UnwrapHost(obj);  // czUI fix: 包装器透明解包
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    if (!obj || !name) return 0;
    vb6_ho_setVariantEmpty(out);

    if (vb6_uc_isControls(obj)) {
        if (_wcsicmp(name, L"Count") == 0) { vb6_ho_setVariantLong(out, vb6_UC_ControlsCount(obj)); return 1; }
        return 1;  // 其它成员当作无值, 避免 C 崩溃
    }
    if (vb6_uc_isColl(obj)) {   // Fix 112c: RTL Collection
        if (_wcsicmp(name, L"Count") == 0) { vb6_ho_setVariantLong(out, vb6_Collection_Count(obj)); return 1; }
        return 1;
    }
    if (vb6_uc_isFont(obj)) return vb6_ho_getFontMember(obj, name, out);

    // Fix <vbeclipse> rev18: UC **实例指针**归一成宿主 HWND —— 晚绑定属性读
    // (`Dim x As Variant: Set x = Controls.Item(id)` 之后 `x.LeftPos`) 传进来的是
    // `vb6_cls_X*`, 而下面整条内建链只认 HWND, 以前一律落到末尾返回 Empty。
    { void* h18 = vb6_UC_HwndOf(obj); if (h18) obj = h18; }

    vb6_UCRec* r = vb6_uc_findByHwnd(obj);
    vb6_HostObjRec* h = vb6_ho_find(obj);
    if (!r && !h && !IsWindow((HWND)obj)) return 0;

    int32_t isForm = r ? 0 : vb6_ho_isForm(obj);
    int32_t scaleMode = r ? (r->desc ? r->desc->scaleMode : 1) : 1;

    if (_wcsicmp(name, L"hwnd") == 0 || _wcsicmp(name, L"hWnd") == 0) {
        vb6_ho_setVariantLong(out, (int32_t)(intptr_t)obj); return 1;
    }
    if (_wcsicmp(name, L"Name") == 0) {
        vb6_ho_setVariantBstr(out, vb6_BSTR_FromStr(h && h->name[0] ? h->name : (r && r->ctrlName[0] ? r->ctrlName : L"")));
        return 1;
    }
    if (_wcsicmp(name, L"Index") == 0) {
        vb6_ho_setVariantLong(out, h ? h->index : (r ? r->index : -1)); return 1;
    }
    if (_wcsicmp(name, L"ScaleMode") == 0) { vb6_ho_setVariantLong(out, scaleMode); return 1; }
    if (_wcsicmp(name, L"ScaleWidth") == 0 || _wcsicmp(name, L"ScaleHeight") == 0 ||
        _wcsicmp(name, L"Width") == 0   || _wcsicmp(name, L"Height") == 0 ||
        _wcsicmp(name, L"Left") == 0    || _wcsicmp(name, L"Top") == 0) {
        int32_t w = 0, hh = 0;
        if (_wcsicmp(name, L"ScaleWidth") == 0 || _wcsicmp(name, L"ScaleHeight") == 0) {
            vb6_ho_clientTwips(obj, &w, &hh);
        } else if (isForm) {
            RECT rc; GetWindowRect((HWND)obj, &rc);
            w = vb6_XToTwipX(rc.right - rc.left); hh = vb6_YToTwipY(rc.bottom - rc.top);
            if (_wcsicmp(name, L"Left") == 0 || _wcsicmp(name, L"Top") == 0) {
                w = vb6_XToTwipX(rc.left); hh = vb6_YToTwipY(rc.top);
            }
        } else {
            vb6_ho_ctrlRect(obj, NULL, NULL, &w, &hh);
            if (_wcsicmp(name, L"Left") == 0 || _wcsicmp(name, L"Top") == 0) {
                int32_t l = 0, t = 0; vb6_ho_ctrlRect(obj, &l, &t, NULL, NULL);
                w = l; hh = t;
            }
        }
        int32_t val = (_wcsicmp(name, L"ScaleWidth") == 0 || _wcsicmp(name, L"Width") == 0 ||
                       _wcsicmp(name, L"Left") == 0) ? w : hh;
        // ScaleWidth/ScaleHeight 需按对象自身 ScaleMode 换算
        if (_wcsicmp(name, L"ScaleWidth") == 0 || _wcsicmp(name, L"ScaleHeight") == 0) {
            if (scaleMode == 3)   /* Fix 184: twips -> pixels via the shared DPI source */
                val = (_wcsicmp(name, L"ScaleHeight") == 0) ? vb6_TwipToY(val) : vb6_TwipToX(val);
        }
        vb6_ho_setVariantLong(out, val); return 1;
    }
    // Fix <vbeclipse>: `Visible` 必须读**自身 WS_VISIBLE 位**, 不能用 IsWindowVisible。
    // IsWindowVisible 是 Win32 的**链式**语义 —— 自身可见但任一祖先不可见就返回 0。
    // 而 VB6 的 `对象.Visible` 只反映该对象自身的 Visible 属性, 与祖先无关。
    // 混用会让停靠面板全空: ucFolder.ViewArea_Resize 里
    // `If View.View.Visible Then View.View.Move ...` 是布局的唯一驱动, 而它
    // 全部发生在 frmMain Form_Load 期间 —— 那时 frmMain 自己还没 ShowWindow,
    // 链上层层 IsWindowVisible=0 ⇒ Move 永不执行 ⇒ 视图窗体停在 0x0。
    // (实测 play78: [VR] enter x32, 每轮每项 vis=0; 同轮 [ST] 显示
    //  ShowView 已把视图 style 由 460B0040 写成 560B0040, WS_VISIBLE=1。)
    // 顺带这也修了拖动/最小化期间的误判: 祖先临时隐藏时对象自身 Visible 不变。
    if (_wcsicmp(name, L"Visible") == 0)  {
        LONG st = GetWindowLongW((HWND)obj, GWL_STYLE);
        vb6_ho_setVariantLong(out, (st & WS_VISIBLE) ? -1 : 0); return 1;
    }
    if (_wcsicmp(name, L"Enabled") == 0)  { vb6_ho_setVariantLong(out, IsWindowEnabled((HWND)obj) ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"hDC") == 0)      { vb6_ho_setVariantLong(out, (int32_t)(intptr_t)GetDC((HWND)obj)); return 1; }
    if (_wcsicmp(name, L"Count") == 0)    { vb6_ho_setVariantLong(out, 0); return 1; }
    if (_wcsicmp(name, L"Controls") == 0) {
        vb6_UCControls* c = vb6_uc_newControls(obj);
        vb6_ho_setVariantDispatch(out, c); return 1;
    }
    if (_wcsicmp(name, L"Caption") == 0 || _wcsicmp(name, L"Text") == 0) {
        // Fix 190: 窗口层是 Unicode —— 必须 GetWindowTextW。旧代码用 GetWindowTextA
        // 再逐字节提升为 wchar_t (= 按 Latin-1 解码 ACP 字节), 中文读回来就是乱码,
        // UserControl 里 `Parent.Caption` 一类的读取全受影响。
        wchar_t wbuf[512] = {0};
        GetWindowTextW((HWND)obj, wbuf, 511);
        vb6_ho_setVariantBstr(out, vb6_BSTR_FromStr(wbuf)); return 1;
    }
    // 字体属性 (Font / TitleFont / *Font) → 返回 Font 代理
    {
        size_t ln = wcslen(name);
        if (_wcsicmp(name, L"Font") == 0 ||
            (ln > 4 && _wcsicmp(name + ln - 4, L"Font") == 0)) {
            vb6_UCFontRec* fr = (vb6_UCFontRec*)vb6_UC_NewFont();
            vb6_ho_setVariantDispatch(out, fr); return 1;
        }
    }
    // Fix <vbeclipse> rev18: 内建成员都没命中 → 试**本 UserControl 的自有 Property**
    // (ucFolder.LeftPos/TopPos/RightPos/BottomPos 这类)。cgen 为每个 .ctl 发的
    // vb6_UcPropDesc 表就是为这一档准备的 —— 没有它, ucPerspective.Refresh 的
    // `l_ucFolder.Move .LeftPos, .TopPos, …` 会读到 0, 停靠区就摆不出来。
    if (vb6_UC_OwnPropGet(obj, name, out)) return 1;

    return 1;   // 宿主对象: 未知属性一律返回 Empty (不崩)
