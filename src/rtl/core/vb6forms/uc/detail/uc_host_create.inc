// vb6forms_uc 拆分片段 uc_host_create (函数体)
// 内容 = 拆分前 vb6forms_uc.c 第 634~738 行，逐行未改
// 在函数体内 #include；路径一律用 basename

    // Fix <vbeclipse> 2026-10-12 (第十二层): 一次性取走"推迟激活"令牌 —— 必须在**本函数
    // 最前面**取: 后面 desc->init 内部还会为设计子控件再走一次 HostCreate, 那些实例要按
    // 原行为立即 Show/Resize (它们没有 bag, 也就没有"之后再激活"的机会)。
    int deferActivate = vb6_uc_deferActivateTake();

    vb6_uc_trace("create.begin", typeName, NULL);
    const vb6_UserControlDesc* desc = vb6_uc_findDesc(typeName);
    if (!desc) return NULL;
    if (vb6_ucRecCount >= VB6_UC_MAX_INST) return NULL;

    // czUI fix: Ambient.BackColor 应反映容器窗体的 BackColor (VB6 语义)。
    // 控件 RedrawControl 用它做底色填充, 取不到就落到 BTNFACE 灰/系统色,
    // 深色窗体上圆角周围会出现灰圈。
    {
        // Fix 190: 窗口层已全面 Unicode —— 类名必须用 GetClassNameW 读, 再转回
        // RTL 内部的 UTF-8 键 (与 vb6_RegisterFormClassBg 登记时用的是同一套字节)。
        wchar_t parentClassW[128] = {0};
        if (hParent && GetClassNameW((HWND)hParent, parentClassW, 127) > 0) {
            char parentClass[256] = {0};
            vb6_wideToU8Buf(parentClassW, parentClass, (int)sizeof(parentClass));
            int bg = vb6_Forms_QueryClassBg(parentClass);
            if (bg >= 0) vb6_Ambient_BackColor = bg;
        }
    }

    vb6_uc_registerClass((HINSTANCE)hInstance);
    vb6_uc_gdiplusInit();

    vb6_UCRec* r = &vb6_ucRecs[vb6_ucRecCount];
    memset(r, 0, sizeof(*r));
    r->desc = desc;
    r->parent = (HWND)hParent;
    r->enabled = -1;
    r->index = index;
    // Fix 119: 装入 cgen 预先设定的实例字体 (.frm BeginProperty Font)
    if (vb6_ucPendingFont) {
        r->font = vb6_ucPendingFont;
        vb6_ucPendingFont = NULL;
    }
    r->scaleWidth = vb6_TwipToX(width);
    r->scaleHeight = vb6_TwipToY(height);
    r->extLeft = left;
    r->extTop = top;
    r->extWidth = width;    // Fix 162: 外框尺寸 (缇), 供 Extender/UserControl
    r->extHeight = height;  //          的 Width/Height 读取
    if (ctrlName) {
        // Fix 190: ctrlName 来自生成代码的字面量 = UTF-8。旧代码逐字节提升为
        // wchar_t 等于按 Latin-1 解码, 中文控件名会变成一串乱码宽字符。
        wchar_t* wctrl = vb6_u8ToWideDup(ctrlName);
        if (wctrl) {
            wcsncpy(r->ctrlName, wctrl, VB6_UC_NAME_LEN - 1);
            r->ctrlName[VB6_UC_NAME_LEN - 1] = 0;
            free(wctrl);
        } else {
            r->ctrlName[0] = 0;
        }
    }

    // 账 #305 第二百五十八刀: 无窗宿主 (InvisibleAtRuntime) —— 记录已填好, 窗口一个都不建。
    // 设计期子控件靠 vb6_UserControl_hWnd == NULL 自己早退 (vb6_UC_CreateDesignImage 那条
    // `if (!parent) return NULL;` 就是现成的闸), 不用在这里另写一套。
    const int windowless = vb6_uc_windowless;
    HWND hwnd = NULL;
    if (!windowless) {
    hwnd = CreateWindowExW(0, L"VB6_UserControlHost", L"",
                                // Fix 124: WS_CLIPSIBLINGS 必不可少 —— 没有它, 一个窗口
                                // 重绘时会画到同层兄弟控件上, 于是 z 序在下的不透明面板
                                // (Charts 2020 的 LabelPlus1 白底) 每次重绘都会把上层控件
                                // (三个 ucProgressCircular 圆环) 擦掉。离屏 dump 是各控件
                                // 单独画进自己的 DIB 再合成, 体现不出这个裁剪问题 ——
                                // 这就是"dump 里有、软件里看不到"的根因。
                                WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
                                vb6_TwipToX(left), vb6_TwipToY(top),
                                vb6_TwipToX(width), vb6_TwipToY(height),
                                (HWND)hParent, NULL, (HINSTANCE)hInstance, r);
    vb6_uc_trace("window.created", typeName, hwnd);
    if (!hwnd) return NULL;
    /* 账 #305 第二百七十八刀: VB6 的 UC 宿主窗口自己答字体 (参考那 29 枚每枚都带), C3 这一户
     * 以前一个都没发过。设计期 BeginProperty Font 走 r->font, 没有就交 NULL = VB6 默认档。
     * 必须在 desc->init 之前发: 初始化里就画第一帧的 UC 要看见自己那枚字体。 */
    vb6_UC_ApplyFontToWindow(hwnd, r->font);
    }
    r->hwnd = hwnd;
    /* 账 #305 第299刀: UC 实例也要被戳 (它可能是 `.frm` 嵌进 OCX 块那一族)。无窗宿主
       (windowless, 第二百五十八刀) 时 hwnd 为 NULL ⇒ vb6_GeomOwnerStamp 自己早退。 */
    vb6_GeomOwnerStamp(hwnd, vb6_GeomAuthoredOwner(hParent));
    vb6_ucRecCount++;

    // Fix 122: 修正 VB6 z 序 (先声明者在上)
    if (!windowless) vb6_uc_fixZOrder(hwnd);

    // 登记: 供 Controls 枚举 / TypeName / 宿主对象分派
    if (vb6_ucHoCount < VB6_UC_MAX_OBJ) {
        vb6_HostObjRec* h = &vb6_ucHo[vb6_ucHoCount++];
        memset(h, 0, sizeof(*h));
        h->hwnd = (void*)hwnd;
        h->isForm = 0;
        h->index = index;
        h->instance = r->me;
        for (size_t i = 0; r->ctrlName[i] && i < VB6_UC_NAME_LEN - 1; i++) h->name[i] = r->ctrlName[i];
        {
            const char* p = desc->typeName;
            const char* dot = strrchr(desc->typeName, '.');
            if (dot) p = dot + 1;
            size_t n = strlen(p);
            if (n >= VB6_UC_NAME_LEN) n = VB6_UC_NAME_LEN - 1;
            for (size_t i = 0; i < n; i++) h->typeName[i] = (wchar_t)(unsigned char)p[i];
            h->typeName[n] = 0;
        }
    }

    // 实例化 + 初始化 (换入宿主状态, 让 UserControl_Initialize 读到正确的 ScaleMode)
    {
        vb6_UCSaved saved;
        vb6_uc_push(r, &saved);
if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) { fprintf(stderr, "[C3_UC] HostCreate stage=create type=%s\n", typeName); fflush(stderr); }
    r->me = desc->create ? desc->create() : NULL;
    vb6_heap_probe("after-create", typeName);
    if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) { fprintf(stderr, "[C3_UC] HostCreate stage=init type=%s me=%p\n", typeName, r->me); fflush(stderr); }
    if (r->me && desc->init) desc->init(r->me);
    vb6_heap_probe("after-init", typeName);
    if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) { fprintf(stderr, "[C3_UC] HostCreate stage=init-done type=%s\n", typeName); fflush(stderr); }
        r->ready = 1;   // Fix 112: 之后窗口消息 (Resize/Show/Paint) 才允许进入实例
        vb6_uc_trace("instance.ready", desc->typeName, r->me);
        vb6_uc_pop(&saved);
        if (vb6_ucHoCount) vb6_ucHo[vb6_ucHoCount - 1].instance = r->me;
    }

    if (!windowless) ShowWindow(hwnd, SW_SHOW);
    // Fix 112: 窗口创建期间被 ready=0 拦下的 Resize/Show 在此补发
    // Fix <vbeclipse> 2026-10-12 (第十二层): 加载路径(带 bag 的顶层 UC)推迟到 ReadProperties
    // 之后 —— VB6 的顺序是 Initialize → ReadProperties → Show/Resize, 而 C3 的 ReadProps 由
    // 窗体创建代码在 HostCreate 之后调用。不推迟的话 Resize 会在设计期属性恢复**之前**跑,
    // 读到全是默认值/零值的字段 (实测 czUI: m_ControlType 未恢复 = 0 = czTitleBar ⇒
    // `If m_ControlType = czTitleBar Then ApplyFormRoundCorners` 走错分支 ⇒ AV)。
    // 由 vb6_UC_ActivateHost(hwnd) 在 ReadProps 之后补齐这一段。
    if (!windowless && deferActivate) {
        r->pendingActivate = 1;
        vb6_uc_trace("activate.defer", desc->typeName, r->me);
        return (void*)hwnd;
    }
    if (!windowless) vb6_uc_activate(r);
    // 账 #305 第二百五十八刀: 无窗宿主 (InvisibleAtRuntime) 没有 Show/Resize 可发 (VB6 那台
    // 同样不发), 而调用方那枚槽还要做成员访问 ⇒ 交回**实例**而不是句柄;
    // vb6_UC_InstanceOf 那头跟着认实例 (见 uc_host.c)。
    if (windowless) return (void*)r->me;
    return (void*)hwnd;
