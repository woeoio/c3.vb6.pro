// uc_host.c - vb6forms_uc 拆分片：UserControl 宿主实例状态 + 公开 API
//
// 内容 = 拆分前 vb6forms_uc.c 第 41~298 / 613~629 / 631~739 / 741~753 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// UserControl 宿主描述 (生成代码为每个 .ctl 提供)
// ============================================================

/* vb6_UserControlDesc 定义见 vb6forms.h (生成代码与 RTL 共用同一份) */

const vb6_UserControlDesc* g_uc_descs[VB6_UC_MAX_DESC];
int32_t g_uc_descCount = 0;

vb6_UCRec g_uc_recs[VB6_UC_MAX_INST];
int32_t g_uc_recCount = 0;
// Fix 112: 可选生命周期日志；每行立即关闭文件，异常退出也保留最后阶段。
void vb6_uc_trace(const char* phase, const char* type, void* me) {
    const char* path = getenv("C3_UC_TRACE");
    if (!path || !*path) return;
    FILE* f = fopen(path, "a");
    if (!f) return;
    fprintf(f, "%lu %s %s me=%p\n", (unsigned long)GetTickCount(), phase, type ? type : "?", me);
    fclose(f);
}

vb6_UCRec* g_uc_current = NULL;   // 最近一次进入的实例 (供 Refresh/PropertyChange)

// Fix 119: 待应用的实例字体。VB6 里每个控件实例有独立的 Font 对象(.frm 的
// BeginProperty Font 块), 且控件会在 InitProperties 里以 UserControl.Font 为
// 默认字体基准 (如 m_TitleFont.Size = UserControl.Font.Size + 8)。因此字体必须在
// 实例初始化**之前**就位 —— cgen 在 vb6_UC_HostCreate 之前调用
// vb6_UC_SetPendingFont(), HostCreate 把它装进 r->font, push 时既成为
// vb6_UserControl_Font 也成为 Ambient.Font。
vb6_ComIface_Font* g_uc_pendingFont = NULL;

void vb6_UC_SetPendingFont(void* f) {
    g_uc_pendingFont = (vb6_ComIface_Font*)f;
}

// Fix 122: VB6 的 z 序规则 —— .frm 中**先声明**的控件在**最上层**。
// cgen 按 .frm 顺序创建子窗口, 而 Win32 是"后创建者在上" → 顺序恰好相反:
// Form2 里 LabelPlus1 (最后声明) 于是盖住了先声明的三个 ucProgressCircular 圆环,
// 用户看到的就是"圆环不见了"。这里把每个新宿主插到"上一个宿主"**之下**, 使先声明者
// 保持在上 (同父窗口内才处理, 避免跨容器错插)。
static HWND g_uc_lastHost = NULL;
static HWND g_uc_lastHostParent = NULL;

static void vb6_uc_fixZOrder(HWND hwnd) {
    HWND parent = GetParent(hwnd);
    if (g_uc_lastHost && g_uc_lastHostParent == parent && IsWindow(g_uc_lastHost))
        SetWindowPos(hwnd, g_uc_lastHost, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_NOOWNERZORDER);
    g_uc_lastHost = hwnd;
    g_uc_lastHostParent = parent;
}

// ============================================================
// 宿主对象登记表: hwnd → VB6 名 / 类型名
// ============================================================

vb6_HostObjRec g_ho[VB6_UC_MAX_OBJ];
int32_t g_hoCount = 0;

// 合成对象: Controls 集合

// 合成对象: Font (stdole.StdFont 的最小形态)

vb6_UCFontRec* g_uc_fonts;
extern vb6_ComIface_Font g_vb6_UserControl_FontObj;

// Fix 112c: 前置声明 (定义在宿主分派节, Collection 实现会用到)

// 安全解引用守卫: 控件/窗体对象以 HWND 形式传入, 而 HWND 是内核句柄而非用户指针;
// 直接按 tag 结构解引用会触发 0xC0000005. 解引用 tag 前先校验目标地址可读.
int32_t vb6_uc_ptrReadable(const void* p, size_t n) {
    if (!p) return 0;
#ifdef _WIN32
    MEMORY_BASIC_INFORMATION mbi;
    if (VirtualQuery(p, &mbi, sizeof(mbi)) == 0) return 0;
    if (mbi.State != MEM_COMMIT) return 0;
    if (mbi.Protect & (PAGE_NOACCESS | PAGE_GUARD)) return 0;
    return ((const char*)mbi.BaseAddress + mbi.RegionSize) >= ((const char*)p + n);
#else
    (void)n;
    return 1;
#endif
}

int32_t vb6_uc_isControls(const void* p) {
    return vb6_uc_ptrReadable(p, sizeof(int32_t)) &&
           ((const vb6_UCControls*)p)->tag == VB6_UC_CONTROLS_TAG;
}
int32_t vb6_uc_isFont(const void* p) {
    // Fix 112e: 先按身份识别裸字体；不可将 Name 字段当 IDispatch vtable。
    if (!p) return 0;
    if (p == &g_vb6_UserControl_FontObj) return 1;
    for (vb6_UCFontRec* f = g_uc_fonts; f; f = f->next)
        if (p == f->font || p == f) return 1;
    return 0;
}

vb6_HostObjRec* vb6_ho_find(const void* hwnd) {
    for (int32_t i = 0; i < g_hoCount; i++) {
        if (g_ho[i].hwnd == hwnd) return &g_ho[i];
    }
    return NULL;
}

vb6_UCRec* vb6_uc_findByHwnd(const void* hwnd) {
    for (int32_t i = 0; i < g_uc_recCount; i++) {
        if ((void*)g_uc_recs[i].hwnd == hwnd) return &g_uc_recs[i];
    }
    return NULL;
}

vb6_UCRec* vb6_uc_findByInstance(const void* inst) {
    for (int32_t i = 0; i < g_uc_recCount; i++) {
        if (g_uc_recs[i].me == inst) return &g_uc_recs[i];
    }
    return NULL;
}

// ============================================================
// 宿主状态换入/换出 (进程级 vb6_UserControl_* 全局)
// ============================================================

static void vb6_uc_defaultFont(void) {
    // vb6_UserControl_Font / vb6_Ambient_Font 为空时补一个默认字体对象
    if (!vb6_UserControl_Font) {
        vb6_UserControl_Font = (vb6_ComIface_Font*)vb6_UC_NewFont();
        vb6_Ambient_Font = vb6_UserControl_Font;
    }
    if (vb6_UserControl_Font && !vb6_UserControl_Font->Name)
        vb6_UserControl_Font->Name = vb6_BSTR_FromStr(L"MS Sans Serif");
    if (!vb6_Ambient_Font) vb6_Ambient_Font = vb6_UserControl_Font;
}

void vb6_uc_push(vb6_UCRec* r, vb6_UCSaved* saved) {
    saved->scaleWidth = vb6_UserControl_ScaleWidth;
    saved->scaleHeight = vb6_UserControl_ScaleHeight;
    saved->scaleMode = vb6_UserControl_ScaleMode;
    saved->hDC = vb6_UserControl_hDC;
    saved->containerHwnd = vb6_UserControl_ContainerHwnd;
    saved->enabled = vb6_UserControl_Enabled;
    saved->font = vb6_UserControl_Font;
    saved->ambientFont = vb6_Ambient_Font;
    saved->extLeft = vb6_Extender_Left;
    saved->extTop = vb6_Extender_Top;
    saved->extWidth = vb6_Extender_Width;      // Fix 162
    saved->extHeight = vb6_Extender_Height;    // Fix 162
    saved->hWnd = vb6_UserControl_hWnd;
    saved->autoRedraw = vb6_UserControl_AutoRedraw;
    saved->ext = vb6_UserControl_Extender;
    saved->current = g_uc_current;
    saved->displayName = (void*)vb6_Ambient_DisplayName;

    vb6_uc_defaultFont();
    // Fix <vbeclipse> rev18: UserControl.ScaleWidth/ScaleHeight 按**缇**交出去。
    // r->scaleWidth/height 一路存的是**像素** (uc_host_create.inc 的
    // `vb6_TwipToX(width)` / `rc.right`、uc_host_window.c WM_SIZE 的 LOWORD(lParam)),
    // 而 desc->scaleMode 声明的是 1 (缇), 而且全链路的坐标都按缇:
    //   · `vb6_ControlMove` (所有 `子控件.Move 20,20,ScaleWidth-30,…` 都经它) 收缇;
    //   · `vb6_UC_CreateDesignEdit` 的 .ctl 设计期 Left/Top/Width/Height 是缇;
    //   · ucPerspective 的 `LeftPos = UserControl.ScaleWidth` 之后要
    //     `LeftPos / Screen.TwipsPerPixelX` 才是像素 (见 CalculateFolderPositionByRef)。
    // 只有这两处是像素 ⇒ `ViewCaption.Move 20, 20, ScaleWidth - 30, …` 会得到 35px 宽的
    // 子控件, 停靠区 7 个文件夹宿主也只算出 27x27/42x8 那一批 (实测 play78 r24)。
    // (对照: 宿主对象路径的 `obj.ScaleWidth` 走 vb6_ho_clientTwips, 本来就是缇。)
    vb6_UserControl_ScaleWidth = vb6_XToTwipX(r->scaleWidth);
    vb6_UserControl_ScaleHeight = vb6_YToTwipY(r->scaleHeight);
    vb6_UserControl_ScaleMode = r->desc ? r->desc->scaleMode : 1;
    vb6_UserControl_hDC = r->hdc;
    vb6_UserControl_ContainerHwnd = (int32_t)(intptr_t)r->parent;
    vb6_UserControl_Enabled = r->enabled;
    if (r->font) {
        vb6_UserControl_Font = (vb6_ComIface_Font*)r->font;
        vb6_Ambient_Font = (vb6_ComIface_Font*)r->font;
    }
    vb6_Extender_Left = r->extLeft;
    vb6_Extender_Top = r->extTop;
    // Fix 162: 外框尺寸同步。UserControl.<Member> 与 Extender.<Member> 由 cgen
    // 发成两个不同标识符, VB6 里同值, 故四处一起填。
    vb6_Extender_Width = r->extWidth;
    vb6_Extender_Height = r->extHeight;
    vb6_UserControl_Width = r->extWidth;
    vb6_UserControl_Height = r->extHeight;
    vb6_UserControl_hWnd = r->hwnd;                 // Fix 133u
    vb6_UserControl_AutoRedraw = 1;                 // Fix 133u: 事件驱动重绘
    vb6_UserControl_Extender.Visible = -1;          // Fix 133u: 默认可见
    vb6_UserControl_Extender.Height = r->scaleHeight; // Fix 133u
    g_uc_current = r;

    // Fix 116: Ambient.DisplayName = 控件实例名 (VB6 语义)。
    // 控件内常见用法: m_Title = .ReadProperty("Title", Ambient.DisplayName)
    // → 此前恒为 NULL, 于是 ucChartArea1 / ucPieChart1 / ucTreeMaps1 等标题全空。
    // 名字在实例创建时写入 r->ctrlName, 这里惰性缓存成 BSTR 复用, 避免每次
    // push 都分配 (push 在每次绘制/事件都会发生)。
    if (!r->displayNameBstr && r->ctrlName[0])
        r->displayNameBstr = SysAllocString(r->ctrlName);
    vb6_Ambient_DisplayName = r->displayNameBstr;
}

void vb6_uc_pop(const vb6_UCSaved* saved) {
    vb6_UserControl_ScaleWidth = saved->scaleWidth;
    vb6_UserControl_ScaleHeight = saved->scaleHeight;
    vb6_UserControl_ScaleMode = saved->scaleMode;
    vb6_UserControl_hDC = saved->hDC;
    vb6_UserControl_ContainerHwnd = saved->containerHwnd;
    vb6_UserControl_Enabled = saved->enabled;
    vb6_UserControl_Font = (vb6_ComIface_Font*)saved->font;
    vb6_Ambient_Font = (vb6_ComIface_Font*)saved->ambientFont;
    vb6_Extender_Left = saved->extLeft;
    vb6_Extender_Top = saved->extTop;
    vb6_Extender_Width = saved->extWidth;            // Fix 162
    vb6_Extender_Height = saved->extHeight;          // Fix 162
    vb6_UserControl_Width = saved->extWidth;         // Fix 162
    vb6_UserControl_Height = saved->extHeight;       // Fix 162
    vb6_UserControl_hWnd = saved->hWnd;              // Fix 133u
    vb6_UserControl_AutoRedraw = saved->autoRedraw;  // Fix 133u
    vb6_UserControl_Extender = saved->ext;           // Fix 133u
    g_uc_current = saved->current;
    vb6_Ambient_DisplayName = (BSTR)saved->displayName;   // Fix 116
}

// 仅换入不换出 — 供「窗体代码直接调用控件公开方法」路径使用
// (vb6_ucChartBar_AddSerie(...) 内部会读 ScaleWidth / 调 UserControl.Refresh)。
void vb6_UC_Enter(void* hwnd) {
    vb6_UCRec* r = vb6_uc_findByHwnd(hwnd);
    if (!r) return;
    vb6_UCSaved saved;
    vb6_uc_push(r, &saved);   // 有意不弹栈: "当前实例"语义 = 最近进入者
}

void vb6_UC_RefreshCurrent(void) {
    // czUI fix: 不能 invalidate+update — UserControl_Paint/RedrawControl 末尾的
    // `If AutoRedraw Then UserControl.Refresh` 会在 WM_PAINT 内强制同步重绘,
    // 形成 PAINT→Refresh→PAINT 无限循环 (czFormDemo 顶部 ~110fps 闪烁)。
    // RTL 绘制是立即模式 (直接画到窗口 DC, 无 AutoRedraw 离屏位图),
    // WM_PAINT 本身已完成全量绘制, Refresh 在此架构下无事可做。
    (void)0;
}

// Fix 133u: UserControl.Cls — 清空控件的客户区背景, 下一轮 WM_PAINT 重绘.
// .ctl 里控件重绘是事件驱动 (InvalidateRect → UserControl_Paint → GDI+ 绘制),
// 此处置位背景即可, 具体清空效果由绘制路径自然覆盖.
void vb6_UserControl_Cls(void) {
    // czUI fix: 不能用 InvalidateRect 实现 — 控件在 RedrawControl 开头调用 Cls,
    // 而 RedrawControl 又由 WM_PAINT 驱动, invalidate 会造成
    // "PAINT→Cls→invalidate→PAINT" 无限重绘循环 (czFormDemo ~90fps 闪烁)。
    // RTL 的 WM_PAINT 每次 BeginPaint 后都是全新表面, 且控件随后会完整重绘,
    // 因此 Cls 在本架构下等价于空操作。
}

// Fix 133u: UserControl.Parent (容器窗体对象).
//   czUI.ctl 用它做全屏/恢复: Parent.hWnd / Parent.Icon.Handle 读取窗体,
//   Parent.Move 移动窗体, With Parent 内 .Left/.Top/.Width/.Height 经
//   vb6_ComGetIntProp(hwnd, ...) 由"宿主对象 → 属性"解析 (vb6forms_uc.c).
//   这里以控件的容器窗口为起点, GetAncestor(GA_ROOT) 找到顶层窗体窗口.
static HWND vb6_uc_ParentHwndRaw(void) {
    if (!g_uc_current) return NULL;
    HWND ctrl = g_uc_current->hwnd ? g_uc_current->hwnd
               : (HWND)(intptr_t)g_uc_current->parent;
    if (!ctrl) return NULL;
    HWND root = GetAncestor(ctrl, GA_ROOT);
    HWND parent = g_uc_current->parent;
    if (parent && root && parent != root)
        return root;              // 顶层窗体窗口
    return parent ? parent : root;
}

void* vb6_UC_ParentObject(void)     { return (void*)vb6_uc_ParentHwndRaw(); }
void* vb6_UC_ParentHwnd(void)       { return (void*)vb6_uc_ParentHwndRaw(); }

void* vb6_UC_ParentIconHandle(void) {
    HWND fw = vb6_uc_ParentHwndRaw();
    if (!fw) return NULL;
    HANDLE icon = (HANDLE)SendMessageW(fw, WM_GETICON, ICON_BIG, 0);
    if (!icon)
        icon = (HANDLE)(LONG_PTR)GetClassLongPtrW(fw, GCLP_HICON);
    return (void*)icon;
}

void vb6_UC_ParentMove(int32_t left, int32_t top, int32_t width, int32_t height) {
    HWND fw = vb6_uc_ParentHwndRaw();
    if (fw) MoveWindow(fw, left, top, width, height, TRUE);
}

void vb6_UC_Register(const vb6_UserControlDesc* desc) {
    if (!desc || !desc->typeName) return;
    for (int32_t i = 0; i < g_uc_descCount; i++) {
        if (_stricmp(g_uc_descs[i]->typeName, desc->typeName) == 0) return;  // 已注册
    }
    if (g_uc_descCount < VB6_UC_MAX_DESC) g_uc_descs[g_uc_descCount++] = desc;
}

static const vb6_UserControlDesc* vb6_uc_findDesc(const char* typeName) {
    if (!typeName) return NULL;
    const char* dot = strrchr(typeName, '.');
    if (dot) typeName = dot + 1;   // 接受 "Proyecto1.ucChartBar"
    for (int32_t i = 0; i < g_uc_descCount; i++) {
        if (_stricmp(g_uc_descs[i]->typeName, typeName) == 0) return g_uc_descs[i];
    }
    return NULL;
}

void* vb6_UC_HostCreate(const char* typeName, int32_t left, int32_t top,
                        int32_t width, int32_t height, void* hParent, void* hInstance,
                        const char* ctrlName, int32_t index) {
#include "uc_host_create.inc"

}

void* vb6_UC_InstanceOf(void* hwnd) {
    vb6_UCRec* r = vb6_uc_findByHwnd(hwnd);
    return r ? r->me : NULL;
}

void* vb6_UC_HwndOf(void* instance) {
    vb6_UCRec* r = vb6_uc_findByInstance(instance);
    return r ? (void*)r->hwnd : NULL;
}

int32_t vb6_UC_IsHostHwnd(void* hwnd) {
    return vb6_uc_findByHwnd(hwnd) != NULL;
}

// Fix <vbeclipse> rev18: UserControl **自有 Property 的按名桥** (晚绑定访问用)。
// 见 vb6forms_controls.h 里 vb6_UcPropDesc 的说明。obj 可以是 UC 实例指针, 也可以是
// 宿主 HWND (两种形态的调用点都有); thunk 一律拿**实例**调用 (desc->props[i].get/set 的
// 第一个形参是 `vb6_cls_X*`)。
static const vb6_UcPropDesc* vb6_uc_ownPropLookup(void* obj, const wchar_t* name, void** outInst) {
    vb6_UCRec* r;
    if (!obj || !name) return NULL;
    r = vb6_uc_findByInstance(obj);
    if (!r) r = vb6_uc_findByHwnd(obj);
    if (!r || !r->desc || !r->desc->props || r->desc->propCount <= 0) return NULL;
    for (int32_t i = 0; i < r->desc->propCount; i++) {
        const vb6_UcPropDesc* pd = &r->desc->props[i];
        if (pd->name && _wcsicmp(pd->name, name) == 0) {
            if (outInst) *outInst = r->me;
            return pd;
        }
    }
    return NULL;
}

int32_t vb6_UC_OwnPropGet(void* obj, const wchar_t* name, void* outV) {
    void* inst = NULL;
    const vb6_UcPropDesc* pd = vb6_uc_ownPropLookup(obj, name, &inst);
    if (!pd || !pd->get || !inst) return 0;
    pd->get(inst, outV);
    return 1;
}

int32_t vb6_UC_OwnPropSet(void* obj, const wchar_t* name, const void* inV) {
    void* inst = NULL;
    const vb6_UcPropDesc* pd = vb6_uc_ownPropLookup(obj, name, &inst);
    if (!pd || !pd->set || !inst) return 0;
    return pd->set(inst, inV) ? 1 : 0;
}

// ---- 设计器子控件: .ctl 设计面上的 TextBox → 每实例一个真实 EDIT 子窗口 ----
void* vb6_UC_CreateDesignEdit(int32_t left, int32_t top, int32_t width, int32_t height) {
    HWND parent = (HWND)vb6_UserControl_hWnd;
    if (!parent) return NULL;
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtrW(parent, GWLP_HINSTANCE);
    HWND edit = CreateWindowExW(0, L"EDIT", L"",
                                WS_CHILD | ES_AUTOHSCROLL,  /* 初始隐藏, 由控件代码控制 Visible */
                                vb6_TwipToX(left), vb6_TwipToY(top),
                                vb6_TwipToX(width), vb6_TwipToY(height),
                                parent, NULL, hInst, NULL);
    if (edit) {
        HFONT f = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        SendMessageW(edit, WM_SETFONT, (WPARAM)f, TRUE);
    }
    return edit;
}

// ---- czUI fix: 设计器 Timer (.ctl 设计面上的 VB.Timer) 逐实例实例化 ----
// 隐藏窗口 + 真实 SetTimer; WM_TIMER 时按 VB6_TimerEnabled/Interval 属性决定
// 是否回调 (vb6_SetTimerEnabled/Interval 已把这些值写进窗口属性)。
// 这让 czUI 的 toggle 滑动动画 / 全屏自动隐藏标题栏真正运转起来。
int vb6_uc_timersStarted = 0;  // czUI fix: 消息循环启动后设计器 Timer 才触发

static LRESULT CALLBACK vb6_uc_timerProc(HWND h, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_TIMER) {
        void (*cb)(void*) = (void (*)(void*))GetPropW(h, L"VB6_TimerCb");
        void* ctx = GetPropW(h, L"VB6_TimerCtx");
        if (!cb) return 0;
        if (!vb6_uc_timersStarted) return 0;
        if (getenv("C3_NO_TIMER_CB")) return 0;  /* bisect */
        if (!vb6_GetTimerEnabled(h)) return 0;
        // Interval 属性变化时重设 SetTimer 周期
        INT_PTR want = (INT_PTR)vb6_GetTimerInterval(h);
        INT_PTR cur = (INT_PTR)GetPropW(h, L"VB6_TimerCur");
        if (want > 0 && want != cur) {
            SetTimer(h, 1, (UINT)want, NULL);
            SetPropW(h, L"VB6_TimerCur", (HANDLE)want);
        }
        // czUI fix: 回调前换入归属实例的宿主上下文 (tmrTrack_Timer 读
        // ScaleWidth/hWnd 等按当前实例; 全局句柄已被最后创建的实例覆盖)
        vb6_UCRec* rec = NULL;
        for (int i = 0; i < g_uc_recCount; i++) {
            if (g_uc_recs[i].me == ctx) { rec = &g_uc_recs[i]; break; }
        }
        if (rec && !rec->ready) return 0;
        if (rec) {
            vb6_UCSaved saved;
            vb6_uc_push(rec, &saved);
            // czUI fix: 非 WM_PAINT 路径没有 DC, RedrawControl 的
            // GdipCreateFromHDC(NULL) 会静默失败 (开关动画不动的根因) —
            // push 已把 rec->hdc 拷进全局 vb6_UserControl_hDC, 这里直接
            // 用 GetDC 覆盖全局, pop 时恢复旧值
            vb6_UserControl_hDC = NULL;
            cb(ctx);
            vb6_uc_pop(&saved);
            InvalidateRect(rec->hwnd, NULL, FALSE);
            UpdateWindow(rec->hwnd);
        } else {
            cb(ctx);
        }
        return 0;
    }
    return DefWindowProcW(h, msg, wp, lp);
}

// czUI fix: 设计器子控件槽位 — 按当前 UC 上下文存取, 替代被多实例共享的
// 全局句柄变量 (11 个实例只有最后一个的 timer/textbox 生效的根因)。
// 用法: cgen 把设计器子控件句柄变量 emit 成
//   #define vb6_hwnd_txtEmbed (*vb6_UC_DesignSlot("txtEmbed"))
// 无上下文时指向孤儿槽 (NULL), 行为与旧全局一致。
void** vb6_UC_DesignSlot(const char* name) {
    static void* orphan = NULL;
    if (!g_uc_current) return &orphan;
    for (int i = 0; i < g_uc_current->designCount; i++) {
        if (strcmp(g_uc_current->design[i].name, name) == 0)
            return &g_uc_current->design[i].value;
    }
    if (g_uc_current->designCount < VB6_UC_DESIGN_SLOTS) {
        Vb6UcDesignSlot* sl = &g_uc_current->design[g_uc_current->designCount++];
        snprintf(sl->name, sizeof(sl->name), "%s", name);
        sl->value = NULL;
        return &sl->value;
    }
    return &orphan;
}

void* vb6_UC_CreateDesignTimer(void (*cb)(void*), void* ctx) {
    static int clsRegistered = 0;
    if (!clsRegistered) {
        WNDCLASSW wc = {0};
        wc.lpfnWndProc = vb6_uc_timerProc;
        wc.hInstance = GetModuleHandleW(NULL);
        wc.lpszClassName = L"VB6_UC_DesignTimer";
        RegisterClassW(&wc);
        clsRegistered = 1;
    }
    // 消息专用窗口 (不显示), 定时器属性沿用 VB6_TimerEnabled/Interval
    HWND h = CreateWindowExW(0, L"VB6_UC_DesignTimer", L"", WS_OVERLAPPED,
                             0, 0, 0, 0, HWND_MESSAGE, NULL, GetModuleHandleW(NULL), NULL);
    if (!h) return NULL;
    SetPropW(h, L"VB6_TimerCb", (HANDLE)cb);
    SetPropW(h, L"VB6_TimerCtx", (HANDLE)ctx);
    SetTimer(h, 1, 50, NULL);   // 默认 50ms; Interval 属性变化时由 proc 重设
    return h;
}

#ifdef __cplusplus
} // extern "C"
#endif
