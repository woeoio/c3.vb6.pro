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
    /* Fix <vbeclipse> rev18 → 账 #175: UserControl.ScaleWidth/ScaleHeight 按
       **.ctl 声明的 ScaleMode** 交出去 (r->scaleWidth/height 一律存设备像素)。
       rev18 当时按"desc->scaleMode 声明的是 1 (缇)"写了无条件 像素→缇, 那个前提
       是错的: 语料里 9 份 .ctl 有 8 份声明 3=Pixel (第 9 份没写、VB6 默认值也是
       1... 实测见 ai 账 #175), 于是像素型控件的 ScaleWidth 比绘图 DC 大 15 倍,
       Charts 2020 的饼/柱/面积/矩形全画在画布外 (空白)。
       缇那一档 (ucPerspective/ucFolder 那条链) 仍走 vb6_ScalePxToUser 的 mode==1
       分支, 与 rev18 逐字节等价。 */
    int32_t ucMode175 = r->desc ? r->desc->scaleMode : 1;
    vb6_UserControl_ScaleWidth = (int32_t)vb6_ScalePxToUser((double)r->scaleWidth, ucMode175, 0);
    vb6_UserControl_ScaleHeight = (int32_t)vb6_ScalePxToUser((double)r->scaleHeight, ucMode175, 1);
    vb6_UserControl_ScaleMode = ucMode175;
    vb6_UserControl_hDC = r->hdc;
    vb6_UserControl_ContainerHwnd = r->parent;   // 账 #180: 不再折成 int32_t
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

// ============================================================
// 账 #179: **成对**的换入/换出 —— 控件代码不管被谁调都跑在自己的宿主上下文里
//
// 上面那个 vb6_UC_Enter 只在"窗体创建控件"之后换入一次且不弹栈, 于是进程级全局
// (vb6_UserControl_ScaleMode / .hWnd / .hDC / .Font / .Enabled / Extender.* /
//  .BackColor ...) 停在"最近创建的那枚控件"上。实测 (tests/ve_units): 缇型那一枚
// 控件在自己 Initialize 里读 ScaleMode=1, 被容器调同一个方法时读到 3 —— 那是另一枚
// 像素型控件留下的值。同一个坑 #178 用"按实例取"的出口治了 ScaleWidth/TextWidth 两条,
// 但那是逐个成员抄一遍; 这一对是**一处口径**: 发码在每个 .ctl 实例方法的体首
// Push、统一出口尾 Pop, 于是整批成员一次性归位。
// 能这么做的前提: Exit Function/Property 已经走 vb6_proc_exit 统一出口尾
// (cgen_decl_func.cpp / _proc.cpp / _prop.cpp 的"Fix <vbeclipse>: 统一出口"),
// 所以提前返回不会漏弹。
//
// 栈是 RTL 自己的 (生成码只见到两个 void 函数, 不需要 vb6_UCSaved 的定义)。
// 溢出/查不到实例时压一个"原样快照", Pop 时等值还原 ⇒ 与不换入等价, 且配平不破。
// ============================================================
#define VB6_UC_CTX_DEPTH 64
static vb6_UCSaved g_uc_ctxStack[VB6_UC_CTX_DEPTH];
static uint8_t g_uc_ctxReal[VB6_UC_CTX_DEPTH];   // 1 = 真换入过, 需要还原
static int32_t g_uc_ctxDepth = 0;

static void vb6_uc_snapshot(vb6_UCSaved* s) {
    // 把当前全局原样抄一份 (不改变任何值): Pop 时等值还原
    vb6_UCRec* cur = g_uc_current;
    if (cur) { vb6_uc_push(cur, s); return; }     // 已知实例: 重推一次即可自洽
    s->scaleWidth = vb6_UserControl_ScaleWidth;
    s->scaleHeight = vb6_UserControl_ScaleHeight;
    s->scaleMode = vb6_UserControl_ScaleMode;
    s->hDC = vb6_UserControl_hDC;
    s->containerHwnd = vb6_UserControl_ContainerHwnd;
    s->enabled = vb6_UserControl_Enabled;
    s->font = vb6_UserControl_Font;
    s->ambientFont = vb6_Ambient_Font;
    s->extLeft = vb6_Extender_Left;
    s->extTop = vb6_Extender_Top;
    s->extWidth = vb6_Extender_Width;
    s->extHeight = vb6_Extender_Height;
    s->hWnd = vb6_UserControl_hWnd;
    s->autoRedraw = vb6_UserControl_AutoRedraw;
    s->ext = vb6_UserControl_Extender;
    s->current = g_uc_current;
    s->displayName = (void*)vb6_Ambient_DisplayName;
}

void vb6_UC_PushInstance(void* inst) {
    if (g_uc_ctxDepth >= VB6_UC_CTX_DEPTH) return;   // 溢出: 整对空转
    vb6_UCRec* r = inst ? vb6_uc_findByInstance(inst) : NULL;
    int32_t slot = g_uc_ctxDepth++;
    if (r) { vb6_uc_push(r, &g_uc_ctxStack[slot]); g_uc_ctxReal[slot] = 1; }
    else   { vb6_uc_snapshot(&g_uc_ctxStack[slot]); g_uc_ctxReal[slot] = 0; }
}

void vb6_UC_PopInstance(void) {
    if (g_uc_ctxDepth <= 0) return;
    int32_t slot = --g_uc_ctxDepth;
    if (g_uc_ctxReal[slot]) vb6_uc_pop(&g_uc_ctxStack[slot]);
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
    if (!instance) return NULL;
    // 已经就是 HWND 就别再查表 —— vb6_uc_findByInstance 只比指针, 地址空间
    // 碰撞时会把一个真实窗口 HWND 认成某个 UC 实例, 换出那个 rec 里未初始化
    // 的 hwnd 垃圾值, 后续 IsWindow/ShowWindow 全部作用在无效句柄上
    // (实测 frmView* 窗体 Visible=True 静默失效, 停靠面板 0x0)。
    if (IsWindow((HWND)instance)) return instance;
    vb6_UCRec* r = vb6_uc_findByInstance(instance);
    if (r && IsWindow((HWND)r->hwnd)) return (void*)r->hwnd;
    return NULL;
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
    /* 账 #178: 容器→控件的每一次进入都要先把该控件的宿主上下文换进来 (见
       vb6_UC_OwnMethodCall 里那条口径的完整说明)。 */
    vb6_UCRec* r = vb6_uc_findByInstance(inst);
    vb6_UCSaved saved;
    if (r) vb6_uc_push(r, &saved);
    pd->get(inst, outV);
    if (r) vb6_uc_pop(&saved);
    return 1;
}

int32_t vb6_UC_OwnPropSet(void* obj, const wchar_t* name, const void* inV) {
    void* inst = NULL;
    int32_t ok;
    const vb6_UcPropDesc* pd = vb6_uc_ownPropLookup(obj, name, &inst);
    if (!pd || !pd->set || !inst) return 0;
    vb6_UCRec* r = vb6_uc_findByInstance(inst);
    vb6_UCSaved saved;
    if (r) vb6_uc_push(r, &saved);
    ok = pd->set(inst, inV) ? 1 : 0;
    if (r) vb6_uc_pop(&saved);
    return ok;
}

// Fix <vbeclipse> rev21: UserControl **自有 Public Sub 的按名桥**。见
// vb6forms_controls.h 里 vb6_UcMethodDesc 的说明 (为什么缺了它视图窗体停在 0x0)。
// 与 props 桥同一套定位: 先按实例找 rec, 找不到再按宿主 HWND 找 —— 两种形态的
// 调用点都有 (原生 vb6_cls_X* 与被 VB6 包一层的 Variant)。
// ⚠ 形参实参是 **vb6_VARIANT** 数组, 而 cgen 发的 thunk 按 VB6 声明类型拆箱;
//   引擎只负责"按名 + 按实例"路由与**实参个数核对** (Optional 缺省要能对上)。
int32_t vb6_UC_OwnMethodCall(void* obj, const wchar_t* name, int32_t argc,
                             const void* argv[], void* outRet) {
    void* inst = NULL;
    if (!obj || !name) return 0;
    vb6_UCRec* r = vb6_uc_findByInstance(obj);
    if (!r) r = vb6_uc_findByHwnd(obj);
    if (!r || !r->desc || !r->desc->methods || r->desc->methodCount <= 0) return 0;
    for (int32_t i = 0; i < r->desc->methodCount; i++) {
        const vb6_UcMethodDesc* md = &r->desc->methods[i];
        if (!md->name || !md->fn || _wcsicmp(md->name, name) != 0) continue;
        /* 第 12 层留档 (探针已跑完并撤除; **下面三段是已被实测排除的假设, 别再走回头路**):
         *
         *   症状: frmViewProperties / frmViewTasks 停在 0x0。
         *   硬数据: 经本方法桥(`vb6_ComCall` 晚绑定)的 ShowView 恰好 5 次, id 依次
         *   **空 / Views / Help / Diagnose / Events**; frmMain.frm:404-434 配的
         *   ActiveViewId 是 Perspectives / Help / Tasks / ToolBox —— 只有 Help 对上。
         *
         *   ❌ 已排除① `List.cls` 的 Add/Count 口径: 把**原版** List.cls / Folder.cls /
         *      Rectangle.cls 拷进夹具跑 C3, `Contains=True` / `Item(0..2)=Events/Tasks/
         *      Diagnose` / `ActiveViewId` 读回 `Tasks` —— **全对**(夹具在
         *      C:\Users\Administrator\vbework\redim_probe\)。
         *   ❌ 已排除② `ReDim Preserve` 缩小语义: 微软官方文档明说「把数组做得比原本更小,
         *      被删除的元素中的数据将会丢失」⇒ **真缩小**, 我们
         *      `vb6_SafeArrayReDimPreserve1D_T` 的无条件缩小**本来就是对的**。
         *      (曾误以为是"恒等", 别再改。)
         *   ❌ 已排除③ COM 装箱往返: `vb6_ComObject_FromInstance` / `GetInstance`
         *      同一实例(C3_IV_TRACE 下 REJECT=0), 存回取回 ActiveViewId 一致。
         *
         *   ⇒ 剩下的唯一事实: `vb6_Folder_prop_get_ActiveViewId(Folder)` 在
         *   `CreateFolder`(`ucPerspective.ctl:1776`)被调时, 返回的不是配置里写的值。
         *   而该getter 只有两条路能换值: `Views.Contains` 不中 → 清空 → 回落
         *   `Views.Item(0)`。既然 ① 已证明 List/Folder 逻辑本身正确, 那就只剩一种可能:
         *   **传进去的 `Folder` 指针不是 frmMain 配的那个实例** ——
         *   即 `CreateFolder(ByRef Folder As Folder)` 的实参绑定错了。
         *   下一步探针: 在 `vb6_ucPerspective_CreateFolder` 入口打 `*Folder`
         *   (指针 + `m_FolderId` + `m_ActiveViewId`), 与 `Perspective.m_Folders`
         *   里各 Folder 的同一组值对照, 即可定位是"哪个实例被换掉了"。*/
        if (md->argc != argc) return 0;   /* 实参个数对不上 ⇒ 不猜, 交回调用方 */
        /* 账 #178: **容器→控件的每一次进入都必须先把该控件的宿主上下文换进来。**
         * VB6 里控件代码不知道自己是被谁调的: UserControl.ScaleMode / .ScaleWidth /
         * .hDC / .Font / TextWidth 这些宿主成员读的是"当前控件"的状态。本文件里
         * 换入的时机原先只有四类 (窗口消息 paint/mouse/show/size、HostCreate 的
         * init/props、Refresh、composite dump) —— 从容器里直接调控件的公共成员/属性
         * 走的是下面这两个桥, 谁都没换 ⇒ 控件代码读到的是进程级残值。
         * 实测 (tests/ve_units, 两枚同尺寸 UC 只差 ScaleMode):
         *   控件自己的 Initialize 里  I-MODE=1 I-TW=600   (缇, 对)
         *   容器里调 uTw.TW("MMMM")   tw=40               (像素, 因为 ScaleMode 残值 0)
         * 同一段代码两个答案 ⇒ 口径必须收在这里, 而不是让每个控件自己按实例取
         * (rev20 给 ScaleWidth 补的 vb6_UC_ScaleWidthOf(me) 就是被这个坑逼出来的
         * 点状绕法; 换入之后那条路仍然有效, 两者读的是同一个 rec)。*/
        vb6_UCSaved saved178;
        vb6_uc_push(r, &saved178);
        int32_t hit178 = md->fn(r->me, argc, argv, outRet) ? 1 : 0;
        vb6_uc_pop(&saved178);
        return hit178;
    }
    return 0;
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

// ---- Fix <vbeclipse>: 设计器子控件 VB.PictureBox → 每实例一个 STATIC 子窗口 ----
// ucFolder 的 ViewArea 就是一个 PictureBox：视图窗体靠 `SetParent view.hWnd,
// ViewArea.hWnd` 挂到它里面。旧设计子控件循环只建 TextBox/Timer，ViewArea 恒 NULL
// → 视图无处可挂 → 停靠面板全空。WS_CLIPCHILDREN 让子窗体不被容器重绘覆盖。
void* vb6_UC_CreateDesignPicture(int32_t left, int32_t top, int32_t width, int32_t height) {
    HWND parent = (HWND)vb6_UserControl_hWnd;
    if (!parent) return NULL;
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtrW(parent, GWLP_HINSTANCE);
    HWND pic = CreateWindowExW(0, L"STATIC", L"",
                               WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN | SS_NOTIFY,
                               vb6_TwipToX(left), vb6_TwipToY(top),
                               vb6_TwipToX(width), vb6_TwipToY(height),
                               parent, NULL, hInst, NULL);
    return pic;
}

// ---- Fix <vbeclipse> rev32: 设计器子控件里的 VB.Label / VB.Image ----
// ucTab.ctl 的设计面只有 imgIcon(VB.Image) + lblCaption(VB.Label), 而 ucHostInit
// 的发射循环只覆盖 TextBox/Timer/PictureBox/限定名 UC ⇒ 两个槽恒 NULL
// (实测: ucCaption/ucTabStrip/ucFolder 各有 1~3 个 CreateDesign*, ucTab = 0)
// ⇒ `ucTab.ToolTip` setter 里 `lblCaption.ToolTipText = …` 对 NULL 写属性
// ⇒ 0xC0000005, 崩在 ucTabStrip.Add 的 `.ToolTip = ToolTipText`。
// Label 走 SS_LEFT 静态文本; Image 走 SS_NOTIFY(与 Picture 同族, 便于子类化收鼠标)。
void* vb6_UC_CreateDesignLabel(int32_t left, int32_t top, int32_t width, int32_t height) {
    HWND parent = (HWND)vb6_UserControl_hWnd;
    if (!parent) return NULL;
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtrW(parent, GWLP_HINSTANCE);
    return CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT,
                           vb6_TwipToX(left), vb6_TwipToY(top),
                           vb6_TwipToX(width), vb6_TwipToY(height),
                           parent, NULL, hInst, NULL);
}
void* vb6_UC_CreateDesignImage(int32_t left, int32_t top, int32_t width, int32_t height) {
    HWND parent = (HWND)vb6_UserControl_hWnd;
    if (!parent) return NULL;
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtrW(parent, GWLP_HINSTANCE);
    return CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_NOTIFY,
                           vb6_TwipToX(left), vb6_TwipToY(top),
                           vb6_TwipToX(width), vb6_TwipToY(height),
                           parent, NULL, hInst, NULL);
}

// ---- Fix <vbeclipse>: 设计器子控件里"工程内 UserControl" → 每实例一个宿主子窗口 ----
// ucFolder 的 ViewTabs(ucTabStrip)/ViewCaption(ucCaption) 是**另一个 UserControl**。
// 复用 Controls.Add 走的 vb6_UC_HostCreate (它自带 push/pop 宿主上下文, 可安全递归)。
// typeName 不是已登记 UC (如第三方 OCX 子控件) 时 findDesc 返回 NULL → 本函数返回
// NULL, 与"从没建"一致, 不会更糟。parent = 当前 UC 的宿主窗口。
void* vb6_UC_CreateDesignUserControl(const char* typeName, const char* ctrlName,
                                     int32_t left, int32_t top, int32_t width, int32_t height) {
    HWND parent = (HWND)vb6_UserControl_hWnd;
    if (!parent || !typeName) return NULL;
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtrW(parent, GWLP_HINSTANCE);
    return vb6_UC_HostCreate(typeName, left, top, width, height, parent, hInst, ctrlName, -1);
}
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
//   #define vb6_hwnd_txtEmbed (*vb6_UC_DesignSlotOf(me, "txtEmbed"))
// 无上下文时指向孤儿槽 (NULL), 行为与旧全局一致。
//
// Fix VbEclipse: 原签名只吃 name, 依赖全局 g_uc_current。但 cgen 生成的
// UC 实例方法 (vb6_ucFolder_AddView 等) 绝大多数是**直接 C 调用**, 由外部
// 模块 (ucPerspective.CreateFolder) 在任意上下文里发起 —— 此时 g_uc_current
// 已被 pop 成 NULL ⇒ 槽位退化成共享 orphan 槽 (实测恒为 NULL) ⇒
// SetParent(view, NULL) 落空, 视图窗体留在桌面顶层, 停靠面板一片空白。
// 拆成两个入口: Of(inst) 按实例定位 (方法体内有 me, 可靠), 无 inst 的
// 旧入口保留给事件回调等确实只有 hwnd 上下文的场景。
static void** vb6_uc_designSlotIn(vb6_UCRec* r, const char* name) {
    for (int i = 0; i < r->designCount; i++) {
        if (strcmp(r->design[i].name, name) == 0)
            return &r->design[i].value;
    }
    if (r->designCount < VB6_UC_DESIGN_SLOTS) {
        Vb6UcDesignSlot* sl = &r->design[r->designCount++];
        snprintf(sl->name, sizeof(sl->name), "%s", name);
        sl->value = NULL;
        return &sl->value;
    }
    return NULL;   // 槽位耗尽
}

void** vb6_UC_DesignSlotOf(void* inst, const char* name) {
    static void* orphan = NULL;
    vb6_UCRec* r = inst ? vb6_uc_findByInstance(inst) : NULL;
    // 回退到窗口上下文: 事件回调/设计期路径只拿得到 hwnd, 或 inst 尚未登记
    // (HostCreate 里 desc->init(r->me) 之前 rec->me 才刚赋值, 见 uc_host_create.inc)。
    if (!r) r = g_uc_current;
    if (!r) return &orphan;
    void** slot = vb6_uc_designSlotIn(r, name);
    return slot ? slot : &orphan;
}

// Fix <vbeclipse> rev22: 给定一个**设计期子控件窗口**, 反查它属于哪个 UC 实例 +
// 它在设计期槽位里的名字。vb6_ControlMove 改完尺寸后用它决定要不要触发
// `<Ctrl>_Resize` 事件 (VB6 语义: 该控件尺寸变化 ⇒ 其 Resize 事件跑一次)。
// 线性扫 rec 表 —— UC 实例数个位数, 且只在 Move 之后调用, 不在热路径。
// ⚠ 必须要求槽位里的窗口**就是** hwnd: 同一 UC 里可能有多个 HWND, 而 VB6 的事件
//   只认"那个被改了尺寸的控件", 不能凭"同属一个 UC"就全体触发。
int32_t vb6_UC_DesignCtrlOwner(const void* hwnd, void** outInst, const char** outName) {
    if (!hwnd) return 0;
    for (int32_t i = 0; i < g_uc_recCount; i++) {
        vb6_UCRec* r = &g_uc_recs[i];
        for (int j = 0; j < r->designCount; j++) {
            if (r->design[j].value == (void*)hwnd) {
                if (outInst) *outInst = r->me;
                if (outName) *outName = r->design[j].name;
                return 1;
            }
        }
    }
    return 0;
}

// Fix <vbeclipse> rev22: 在**正确**的宿主上下文里跑一次设计期子控件的 Resize 事件。
// 单独开这个口 (而不是让 vb6_forms_ctrl.c 自己 push/pop) 的原因: push/pop 要
// vb6_UCSaved 完整类型, 而把 vb6forms_uc_internal.h include 进 vb6forms_ctrl.c
// 会连带 vb6rtl.h, 三头在同一 TU 里对同一批宿主伪属性的可见性不同 ⇒ 实测直接
// segfault (0xC0000005)。saved 在本单元栈上分配, 天然满足"事件体里 Move 别的
// 控件会再进一层, 共享一块 saved 会被内层 pop 掉外层值"的隔离要求。
// 返回 1 = 事件跑了; 0 = 没跑 (无宿主上下文 / 不是设计期子控件 / 未 ready)。
//
// ⚠ 调用方 (vb6_ControlMove) 必须**已经**判过"不在 UC 上下文里" (g_uc_current
//   为空), 否则这里 push 会与外层 push 打架: 内层 pop 把 g_uc_current 恢复成
//   外层 push 前的值, 外层继续往下走就踩空。
int32_t vb6_UC_RunDesignResize(const void* hwnd) {
    void* inst = NULL;
    const char* ctrlName = NULL;
    if (!hwnd) return 0;
    if (!vb6_UC_DesignCtrlOwner(hwnd, &inst, &ctrlName)) return 0;
    if (!inst) return 0;                              // HostCreate 早期 me 还没赋值
    vb6_UCRec* r = vb6_uc_findByInstance(inst);
    // ready 之前不跑: HostCreate 里槽位是在 ready=1 之前填的 (uc_host_create.inc
    // :30 desc → :102 me → :106 ready), 那段窗口还带着设计期尺寸。
    if (!r || !r->ready || !r->desc || !r->desc->designResize) return 0;
    /* 重入防护靠"**同控件**"而不是"有没有 UC 上下文"。
     * ⚠ 别加 `if (g_uc_current) return 0` —— 那样等于永不触发:
     *   实测 (play78 --arch x86) WM_SIZE 处理尾部 g_uc_current 仍**非空**, 因为
     *   WM_SIZE 是 SetWindowPos 触发的同步 SendMessage, 它在 vb6_uc_pop 之前
     *   就跑完了。而 ViewArea_Resize 恰恰需要在这里被调。
     *   换句话说: "已在 UC 上下文里" 正是本调用点的**常态**, 不是异常。
     * 真正的风险是递归 (事件体 Move 别的控件 → 又发 WM_SIZE → 又调事件),
     * 那个由"同控件不重入 + 深度上限"两层挡住。*/
    static int inEvt = 0;
    static const void* inEvtHwnd = NULL;
    if (inEvt && inEvtHwnd == hwnd) return 0;          // 同控件重入 ⇒ 跳过
    if (inEvt >= 8) return 0;                          // 深度硬上限兜底
    vb6_UCSaved saved;
    const void* prevHwnd = inEvtHwnd;
    int prevIn = inEvt;
    inEvt++; inEvtHwnd = hwnd;
    vb6_uc_push(r, &saved);
    r->desc->designResize(inst, ctrlName);
    vb6_uc_pop(&saved);
    inEvt = prevIn; inEvtHwnd = prevHwnd;
    return 1;
}

// Fix <vbeclipse> rev23: 子控件 Resize 事件要**延到本轮布局收尾后**才跑。
//
// 探针实测 (play78 --arch x86): rev22 把触发挂在 WM_SIZE 尾部。ucFolder 收到 10 次
// WM_SIZE、设计期槽位齐全、`RunDesignResize` 在它身上跑了 30 次 —— 事件确实在跑。
// 但 `vb6_GetControlWidth(ViewArea)` 的读数是**三个阶段混在一起**: 同一 hwnd 先读到
// 1515 再读到 8505 (设计期), 另一个读到 6375 → 8505 → 9330。而 ViewArea 的**最终**
// 尺寸 (窗口树里的 583/253/778) 来自 ucFolder.ctl `UserControl_Resize` 内部那句
// `ViewArea.Move 20, 20, ScaleWidth-30, …` —— 它发生在**整串嵌套 Move 的末尾**。
// ⇒ 事件跑的时候 ViewArea 自己还没拿到最终尺寸, 于是 `ViewArea_Resize` 按当时读到的
//   宽度去 Move 视图窗体 (恒为设计期 8505), 视图窗体停0x0 / 超出容器。
//
// 为什么延后一轮能解决: `SetWindowPos`/`MoveWindow` 触发的 WM_SIZE 是**同步**
// SendMessage, 一整串嵌套 Move 全在同一个调用栈里跑完才返回; 延后触发要等**回到
// 消息循环**, 那时 ViewArea 已是最终尺寸。
//
// 实现: `SetTimer` 一轮即触发 (50ms 的定时器粒度在这里足够, 且不引新结构字段 ——
// rec 是**按值数组**不是指针数组, 加字段会让布局式初始化的跨边界契约再踩一次
// rev22 那个错位坑)。timer proc 里再调一次 `RunDesignResize`。
// Fix <vbeclipse> rev24: 按**子控件 HWND** 排队它所属 UC 的设计期 Resize 事件。
// vb6_ControlMove (vb6forms_ctrl.c) 调它 —— 那个单元不能 include 本头, 故走 extern。
//
// 为什么需要这一层 (宿主 WM_SIZE 那条链不够, 见 vb6forms_ctrl.c 的注释):
//   宿主窗口尺寸在启动后就不再变, 而子控件尺寸**一直在变** —— 变的是子控件自己,
//   它发自己的 WM_SIZE, 不会冒泡到宿主。所以触发点必须在"**该子控件**尺寸变了"。
int32_t vb6_UC_QueueDesignResizeForCtrl(const void* hwnd) {
    void* inst = NULL;
    const char* ctrlName = NULL;
    if (!hwnd) return 0;
    if (vb6_UC_DesignCtrlOwner(hwnd, &inst, &ctrlName) && inst)
        return vb6_UC_QueueDesignResize(inst);
    /* ⚠ Fix <vbeclipse> rev24: **宿主窗口自身**的 Move 也必须排队。
     *   停靠布局里 ucFolder 宿主是被 `l_ucFolder.Move 1, 1, ScaleWidth, ScaleHeight`
     *   (ucPerspective.c:2292/2322) 摆到最终尺寸的 —— 那一发走 host 分派 →
     *   vb6_ControlMove(宿主)。但**宿主自己不在任何 rec 的 design[] 里**
     *   (design[] 只装设计期子控件: ViewTabs/ViewArea/ViewCaption), 所以上面
     *   按 hwnd 反查必然落空。
     *   而它才是**第一因**: 宿主尺寸变 → 其 WM_SIZE → UserControl_Resize →
     *   `ViewArea.Move …` 把 ViewArea 摆到最终尺寸。排队的目的正是让
     *   `ViewArea_Resize` 在**这之后**再跑一次, 那时它读到的才是最终宽度。
     *   探针实测 (play78): 反查失败的恰好是 `569x441 → 320x309` / `138x291` /
     *   `101x...` 这些"把 ViewArea 摆到最终尺寸"那一族, 27 次全部 q=0。*/
    {
        vb6_UCRec* hr = vb6_uc_findByHwnd((HWND)hwnd);
        if (hr) return vb6_UC_QueueDesignResize(hr->me);
    }
    return 0;
}

int32_t vb6_UC_QueueDesignResize(const void* inst) {
    if (!inst) return 0;
    vb6_UCRec* r = vb6_uc_findByInstance(inst);
    if (!r || !r->hwnd || !IsWindow(r->hwnd)) return 0;
    /* ⚠ 为什么用 PostMessage 而不是 SetTimer: 宿主窗口过程里**没有 WM_TIMER 分支**
     * (uc_host_window.c 只处理 NCCREATE/SIZE/SHOWWINDOW/鼠标), DefWindowProc 对
     * WM_TIMER 不会替我们做事 —— 实测 drain 一次都没跑。改成 PostMessage 一个自定义
     * WM_APP, 由 vb6_uc_wndproc 里新增的 case 接住 (与 WM_SIZE 同一条链的思路)。
     *
     * 去重靠窗口属性而不是 rec 字段: rec 是**按值数组**, 加字段会再碰一次 rev22
     * 那个"跨边界布局式初始化的字段顺序"坑。窗口属性是 Win32 自带的关联存储。*/
    if (GetPropW(r->hwnd, VB6_UC_DR_PROP)) return 1;    // 已在队列里 ⇒ 不重复排
    if (!SetPropW(r->hwnd, VB6_UC_DR_PROP, (HANDLE)1)) return 0;
    if (!PostMessageW(r->hwnd, VB6_UC_DR_MSG, 0, 0)) {
        RemovePropW(r->hwnd, VB6_UC_DR_PROP);
        return 0;
    }
    return 1;
}

// WM_APP 消息的处理入口 (由 vb6_uc_wndproc 调用): 布局已收尾, 各设计期子控件尺寸已定。
void vb6_UC_DrainDesignResize(void* hwnd) {
    HWND h = (HWND)hwnd;
    if (!h) return;
    RemovePropW(h, VB6_UC_DR_PROP);                // 先清标志, 允许事件体再排新一轮
    vb6_UCRec* r = vb6_uc_findByHwnd(h);
    if (!r) return;
    for (int32_t i = 0; i < r->designCount; i++)
        if (r->design[i].value) vb6_UC_RunDesignResize(r->design[i].value);
}

void** vb6_UC_DesignSlot(const char* name) {
    return vb6_UC_DesignSlotOf(NULL, name);
}

// ============================================================
// Fix <vbeclipse> rev20: UserControl.ScaleWidth/ScaleHeight 按实例解析
// ------------------------------------------------------------
// 这两个属性在 RTL 里是**进程级全局** (vb6_UserControl_ScaleWidth, 由
// vb6_uc_push 写/ vb6_uc_pop 还原), 于是所有 UC 实例共享一份。
//
// 问题不在 push/pop 本身配对没错, 而在**有一整类调用根本不经过 push**:
// cgen 生成的 UC 实例方法之间是**裸 C 调用**, 不走宿主分派, 所以没有
// 谁替它换入上下文。实测 play78: `ucPerspective.c:518` 等 11 处直接
// `vb6_ucFolder_Refresh(l_ucFolder)`, 而 Refresh (ucFolder.ctl:382) 第一句
// 就是 `ViewArea_Resize` + `UserControl_Resize`, 后者全部按
// `ViewArea.Move 20, 20, ScaleWidth - 30, ScaleHeight - 30` 摆子控件 ——
// 读的是**全局**值。加上 `vb6_UC_Enter` 明确"有意不弹栈"(当前实例语义 =
// 最近进入者), 全局值就永久停在**最后进入的那个实例**上。
//
// 数值证据 (play78): ViewArea 收到的 Move 里 `W=8505` 反复出现,
// 8505 = 8535 - 30, 而 8535 tw = 569 px 恰是 **ucPerspective 的设计期
// 宽度**, 不是 ucFolder 的运行期宽度 ⇒ 跨实例串味, 子控件停在设计期尺寸。
//
// 为什么这里能按实例取值: cgen 在 .ctl 模块里为每个实例方法都发射了形参
// `me` (类实例指针), 且 vb6_uc_findByInstance(me) 就是该实例的 rec —— 与
// 同一个文件里 vb6_UC_DesignSlotOf(me, ...) 的做法一致 (那里也是靠 `me`
// 才正确, 修的是同一个"多实例共享"家族问题)。r->scaleWidth/Height 一律
// 存**像素** (HostCreate 的 vb6_TwipToX / WM_SIZE 的 LOWORD(lParam)),
// 对外按 .ctl 声明的 ScaleMode 交出 (账 #175; rev18 曾无条件按缇)。
//
// 回落: inst 为空 / 查不到 rec 时退回进程级全局, 保持单实例与设计期
// 路径 (vb6_UC_HostCreate 里的 InitProperties 等尚无 me 的场景) 的原行为。
int32_t vb6_UC_ScaleWidthOf(void* inst) {
    vb6_UCRec* r = inst ? vb6_uc_findByInstance(inst) : NULL;
    if (r) return (int32_t)vb6_ScalePxToUser((double)r->scaleWidth,
                                             r->desc ? r->desc->scaleMode : 1, 0);
    return vb6_UserControl_ScaleWidth;
}

int32_t vb6_UC_ScaleHeightOf(void* inst) {
    vb6_UCRec* r = inst ? vb6_uc_findByInstance(inst) : NULL;
    if (r) return (int32_t)vb6_ScalePxToUser((double)r->scaleHeight,
                                             r->desc ? r->desc->scaleMode : 1, 1);
    return vb6_UserControl_ScaleHeight;
}

// 账 #175: 该 HWND 是 UserControl 宿主 → 它 .ctl 声明的 ScaleMode; 否则 0
// (=不是 UC, 由 vb6_ContainerScaleMode 落回窗体的 ScaleMode)。
int32_t vb6_UC_WindowScaleMode(const void* hwnd) {
    if (!hwnd) return 0;
    vb6_UCRec* r = vb6_uc_findByHwnd(hwnd);
    if (!r || !r->desc) return 0;
    return r->desc->scaleMode;
}

/* 账 #177/#178: UserControl.TextWidth / .TextHeight 的**按实例**出口 —— 量出来的是
   设备像素, 折算用的必须是"这一枚控件"声明的 ScaleMode。
   为什么不能读进程级 vb6_UserControl_ScaleMode: 容器直接调控件的公共成员时
   (实测 tests/ve_units 里 frmUnits 调 uTw.TW("MMMM")) 走的是发码**直调**, 不经过
   本文件任何一个 push 点 ⇒ 全局是残值 0 ⇒ 同一句 UserControl.TextWidth 在控件自己的
   Initialize 里给 600 (缇)、被容器调时给 40 (像素)。rev20 给 ScaleWidth 补
   vb6_UC_ScaleWidthOf(me) 治的就是同一个坑, 这里照同一形状办。
   回落与 ScaleWidthOf 一致: 查不到 rec 时用进程级全局。 */
extern int32_t vb6_UC_MeasureTextPx(BSTR text, int wantWidth);

static int32_t vb6_uc_textMeasureOf(void* inst, BSTR text, int wantWidth) {
    vb6_UCRec* r = inst ? vb6_uc_findByInstance(inst) : NULL;
    int32_t px = vb6_UC_MeasureTextPx(text, wantWidth);
    int32_t mode = (r && r->desc) ? r->desc->scaleMode : vb6_UserControl_ScaleMode;
    return (int32_t)vb6_ScalePxToUser((double)px, mode, wantWidth ? 0 : 1);
}

int32_t vb6_UC_TextWidthOf(void* inst, BSTR text) {
    return vb6_uc_textMeasureOf(inst, text, 1);
}

int32_t vb6_UC_TextHeightOf(void* inst, BSTR text) {
    return vb6_uc_textMeasureOf(inst, text, 0);
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
