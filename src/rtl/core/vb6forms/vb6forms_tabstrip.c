// vb6forms_tabstrip.c — VB6 TabStrip 控件 (MSComctlLib.TabStrip) 复刻
//
// 用 comctl32 的 **SysTabControl32** 复刻"只画标签条"的那个控件, **不加载 MSCOMCTL.OCX**。
// 与 SSTab 的区别（见 ai/内置控件/TabStrip 控件（选项卡条）.md）：
//   · **不是容器** —— 页面内容由用户自己在 Click 里切 PictureBox/Frame 的 Visible，
//     所以这里既不接管子控件、也不做"每页显隐"那一套（SSTab 才是容器）；
//   · Tab 可带图标、可多行（MultiRow）、可热跟踪（HotTracking）；
//   · 关键读数是 ClientLeft/Top/Width/Height —— 用户代码按它给页面容器定位。
//
// 属性面（本批）：Tabs（集合：Add/Remove/Clear/Count）、Tab.Caption/Key/Image/Index、
//   SelectedItem（Index/Key/Caption）、ClientLeft/Top/Width/Height、MultiRow/HotTracking/ShowTips。
// 事件面：Click（原生 TCN_SELCHANGE）。
//
// 与 SSTab 的一处刻意分工：SSTab 的页索引是 **0 基**（VB6 的 .Tab 就是 0 基），而
// TabStrip 的 Tabs 集合与 SelectedItem.Index 都是 **1 基** —— 两套口径各自照 VB6，不统一。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6rtl_bstr.h"

#ifdef _WIN32

#define VB6_TS_MAX 32

// 原生通知号（commctrl.h：TCN_FIRST = 0U-550U ⇒ TCN_SELCHANGE = -551）
#define VU_TCN_SELCHANGING (-552)
#define VU_TCN_SELCHANGE   (-551)

typedef struct {
    HWND      hwnd;
    wchar_t** cap;      // 每标签标题
    wchar_t** key;      // 每标签 Key（可空）
    wchar_t** tip;      // 每标签 ToolTipText（可空）
    int*      image;    // 每标签图标下标，-1 = 无
    int       count;
    int       alloc;
    int       cur;      // 当前选中（0 基，-1 = 无）
    int       multiRow;
    int       hotTrack;
    int       showTips;
} Vb6TabStrip;

static Vb6TabStrip g_ts[VB6_TS_MAX];

static Vb6TabStrip* tsFind(void* hwnd) {
    HWND h = (HWND)hwnd;
    int i;
    if (!h) return NULL;
    for (i = 0; i < VB6_TS_MAX; i++)
        if (g_ts[i].hwnd == h) return &g_ts[i];
    return NULL;
}

static wchar_t* tsWcsDup(const wchar_t* s) {
    size_t n;
    wchar_t* p;
    if (!s) return NULL;
    n = wcslen(s);
    p = (wchar_t*)malloc((n + 1) * sizeof(wchar_t));
    if (!p) return NULL;
    memcpy(p, s, (n + 1) * sizeof(wchar_t));
    return p;
}

static void tsFreeRow(Vb6TabStrip* t, int i) {
    if (!t || i < 0 || i >= t->count) return;
    if (t->cap) { free(t->cap[i]); t->cap[i] = NULL; }
    if (t->key) { free(t->key[i]); t->key[i] = NULL; }
    if (t->tip) { free(t->tip[i]); t->tip[i] = NULL; }
}

static int tsReserve(Vb6TabStrip* t, int need) {
    int i, cap;
    if (!t) return 0;
    if (need <= t->alloc) return 1;
    cap = t->alloc ? t->alloc : 8;
    while (cap < need) cap *= 2;
    {
        wchar_t** nc = (wchar_t**)realloc(t->cap, (size_t)cap * sizeof(wchar_t*));
        wchar_t** nk = (wchar_t**)realloc(t->key, (size_t)cap * sizeof(wchar_t*));
        wchar_t** nt = (wchar_t**)realloc(t->tip, (size_t)cap * sizeof(wchar_t*));
        int* ni = (int*)realloc(t->image, (size_t)cap * sizeof(int));
        if (!nc || !nk || !nt || !ni) return 0;
        t->cap = nc; t->key = nk; t->tip = nt; t->image = ni;
    }
    for (i = t->alloc; i < cap; i++) {
        t->cap[i] = NULL; t->key[i] = NULL; t->tip[i] = NULL; t->image[i] = -1;
    }
    t->alloc = cap;
    return 1;
}

// 把 RTL 的平行表整段重灌到 SysTabControl32 上（原生没有"改第 i 项"的便宜路）。
static void tsRepopulate(Vb6TabStrip* t) {
    TCITEMW it;
    int i;
    if (!t || !t->hwnd) return;
    SendMessageW(t->hwnd, TCM_DELETEALLITEMS, 0, 0);
    for (i = 0; i < t->count; i++) {
        ZeroMemory(&it, sizeof(it));
        it.mask = TCIF_TEXT;
        it.pszText = t->cap[i] ? t->cap[i] : L"";
        if (t->image && t->image[i] >= 0) { it.mask |= TCIF_IMAGE; it.iImage = t->image[i]; }
        SendMessageW(t->hwnd, TCM_INSERTITEMW, (WPARAM)i, (LPARAM)&it);
    }
    if (t->cur >= t->count) t->cur = t->count - 1;
    if (t->cur >= 0) SendMessageW(t->hwnd, TCM_SETCURSEL, (WPARAM)t->cur, 0);
}

// ===================== 设计期 =====================
// 样式位在设计期灌一次（cgen 那条 emitTabStripDesignTimeInit 只递观测得到的四枚 bool）。
void vb6_TabStrip_Init(void* hwnd, int32_t multiRow, int32_t hotTrack, int32_t showTips) {
    HWND h = (HWND)hwnd;
    Vb6TabStrip* t = NULL;
    int i;
    if (!h) return;
    t = tsFind(hwnd);
    if (!t) {
        for (i = 0; i < VB6_TS_MAX; i++) if (!g_ts[i].hwnd) { t = &g_ts[i]; break; }
        if (!t) return;
        ZeroMemory(t, sizeof(*t));
        t->hwnd = h;
        t->cur = -1;
    }
    t->multiRow = multiRow ? 1 : 0;
    t->hotTrack = hotTrack ? 1 : 0;
    t->showTips = showTips ? 1 : 0;
    {
        LONG st = GetWindowLongW(h, GWL_STYLE);
        st = multiRow ? (st | TCS_MULTILINE) : (st & ~TCS_MULTILINE);
        st = hotTrack ? (st | TCS_HOTTRACK) : (st & ~TCS_HOTTRACK);
        st = showTips ? (st | TCS_TOOLTIPS) : (st & ~TCS_TOOLTIPS);
        SetWindowLongW(h, GWL_STYLE, st);
    }
}

int32_t vb6_TabStrip_GetMultiRow(void* hwnd)   { Vb6TabStrip* t = tsFind(hwnd); return t ? t->multiRow : 0; }
int32_t vb6_TabStrip_GetHotTracking(void* hwnd){ Vb6TabStrip* t = tsFind(hwnd); return t ? t->hotTrack : 0; }
int32_t vb6_TabStrip_GetShowTips(void* hwnd)   { Vb6TabStrip* t = tsFind(hwnd); return t ? t->showTips : 0; }

void vb6_TabStrip_SetMultiRow(void* hwnd, int32_t v) {
    Vb6TabStrip* t = tsFind(hwnd);
    LONG st;
    if (!t) return;
    t->multiRow = v ? 1 : 0;
    st = GetWindowLongW(t->hwnd, GWL_STYLE);
    st = t->multiRow ? (st | TCS_MULTILINE) : (st & ~TCS_MULTILINE);
    SetWindowLongW(t->hwnd, GWL_STYLE, st);
    InvalidateRect(t->hwnd, NULL, TRUE);
}

void vb6_TabStrip_SetHotTracking(void* hwnd, int32_t v) {
    Vb6TabStrip* t = tsFind(hwnd);
    LONG st;
    if (!t) return;
    t->hotTrack = v ? 1 : 0;
    st = GetWindowLongW(t->hwnd, GWL_STYLE);
    st = t->hotTrack ? (st | TCS_HOTTRACK) : (st & ~TCS_HOTTRACK);
    SetWindowLongW(t->hwnd, GWL_STYLE, st);
}

void vb6_TabStrip_SetShowTips(void* hwnd, int32_t v) {
    Vb6TabStrip* t = tsFind(hwnd);
    if (!t) return;
    t->showTips = v ? 1 : 0;
}

// ===================== Client*（页面容器要用的那块矩形） =====================
// 口径：标签行以下、边框以内。tab 行高按**第一枚标签的真实矩形**量（没有标签时退 24px，
// 那是 comctl32 一个标签行的典型高度）—— 这是用户代码 Move 页面容器时唯一能对的数。
static void tsClientTop(HWND h, int* pTop, int* pLeft) {
    RECT rc;
    int top = 24, left = 2;
    if (!h) { *pTop = 0; *pLeft = 0; return; }
    if (SendMessageW(h, TCM_GETITEMCOUNT, 0, 0) > 0) {
        RECT ir;
        ZeroMemory(&ir, sizeof(ir));
        ir.left = 0; ir.top = 0; ir.right = 0; ir.bottom = 0;
        if (SendMessageW(h, TCM_GETITEMRECT, 0, (LPARAM)&ir)) top = (int)ir.bottom + 2;
    }
    GetClientRect(h, &rc);
    if (top > rc.bottom) top = rc.bottom;
    *pTop = top;
    *pLeft = left;
}

int32_t vb6_TabStrip_GetClientLeft(void* hwnd)   { int t, l; tsClientTop((HWND)hwnd, &t, &l); return l; }
int32_t vb6_TabStrip_GetClientTop(void* hwnd)    { int t, l; tsClientTop((HWND)hwnd, &t, &l); return t; }
int32_t vb6_TabStrip_GetClientWidth(void* hwnd)  {
    RECT rc; int t, l;
    if (!hwnd) return 0;
    GetClientRect((HWND)hwnd, &rc);
    tsClientTop((HWND)hwnd, &t, &l);
    return rc.right - 2 * l;
}
int32_t vb6_TabStrip_GetClientHeight(void* hwnd) {
    RECT rc; int t, l;
    if (!hwnd) return 0;
    GetClientRect((HWND)hwnd, &rc);
    tsClientTop((HWND)hwnd, &t, &l);
    return rc.bottom - t - 2;
}

// ===================== Tabs 集合（1 基，与 VB6 一致） =====================
int32_t vb6_TabStrip_TabsCount(void* hwnd) {
    Vb6TabStrip* t = tsFind(hwnd);
    return t ? t->count : 0;
}

// Tabs.Add [index], [key], [caption], [image]
//   index 是**插入位**（1 基，省略/0 = 追加）—— 与 ListView.ColumnHeaders.Add 同口径。
//   hasIndex/hasKey/hasImage 由 cgen 按"实参是否省略"递，省略时按 0/-1 处理。
void vb6_TabStrip_TabsAdd(void* hwnd, int32_t index, void* keyBstr, void* capBstr,
                          int32_t image, int32_t hasIndex, int32_t hasKey, int32_t hasImage) {
    Vb6TabStrip* t = tsFind(hwnd);
    int at, i;
    if (!t) return;
    if (!tsReserve(t, t->count + 1)) return;
    at = (hasIndex && index >= 1) ? (index - 1) : t->count;
    if (at > t->count) at = t->count;
    if (at < 0) at = 0;
    for (i = t->count; i > at; i--) {
        t->cap[i] = t->cap[i - 1]; t->key[i] = t->key[i - 1];
        t->tip[i] = t->tip[i - 1]; t->image[i] = t->image[i - 1];
    }
    t->cap[at] = capBstr ? tsWcsDup((const wchar_t*)capBstr) : tsWcsDup(L"");
    t->key[at] = (hasKey && keyBstr) ? tsWcsDup((const wchar_t*)keyBstr) : NULL;
    t->tip[at] = NULL;
    t->image[at] = (hasImage && image >= 0) ? image : -1;
    t->count++;
    if (t->cur < 0) t->cur = 0;
    tsRepopulate(t);
}

void vb6_TabStrip_TabsRemove(void* hwnd, int32_t idx1) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i, at;
    if (!t) return;
    at = idx1 - 1;
    if (at < 0 || at >= t->count) return;
    tsFreeRow(t, at);
    for (i = at; i < t->count - 1; i++) {
        t->cap[i] = t->cap[i + 1]; t->key[i] = t->key[i + 1];
        t->tip[i] = t->tip[i + 1]; t->image[i] = t->image[i + 1];
    }
    t->count--;
    if (t->cap[t->count]) { /* 末位已随上移搬走，清指针避免二次释放 */ }
    t->cap[t->count] = NULL; t->key[t->count] = NULL; t->tip[t->count] = NULL; t->image[t->count] = -1;
    if (t->cur >= t->count) t->cur = t->count - 1;
    tsRepopulate(t);
}

void vb6_TabStrip_TabsClear(void* hwnd) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i;
    if (!t) return;
    for (i = 0; i < t->count; i++) tsFreeRow(t, i);
    t->count = 0;
    t->cur = -1;
    SendMessageW(t->hwnd, TCM_DELETEALLITEMS, 0, 0);
}

void* vb6_TabStrip_GetTabCaption(void* hwnd, int32_t idx1) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    if (!t || i < 0 || i >= t->count) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(t->cap[i] ? t->cap[i] : L"");
}

void vb6_TabStrip_SetTabCaption(void* hwnd, int32_t idx1, void* bstr) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    const wchar_t* s = (const wchar_t*)bstr;
    if (!t || i < 0 || i >= t->count) return;
    free(t->cap[i]);
    t->cap[i] = tsWcsDup(s ? s : L"");
    tsRepopulate(t);
}

void* vb6_TabStrip_GetTabKey(void* hwnd, int32_t idx1) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    if (!t || i < 0 || i >= t->count) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(t->key[i] ? t->key[i] : L"");
}

int32_t vb6_TabStrip_GetTabIndexByKey(void* hwnd, void* keyBstr) {
    Vb6TabStrip* t = tsFind(hwnd);
    const wchar_t* k = (const wchar_t*)keyBstr;
    int i;
    if (!t || !k) return 0;
    for (i = 0; i < t->count; i++)
        if (t->key[i] && wcscmp(t->key[i], k) == 0) return i + 1;
    return 0;
}

int32_t vb6_TabStrip_GetSelectedIndex(void* hwnd) {
    Vb6TabStrip* t = tsFind(hwnd);
    LRESULT sel;
    if (!t) return 0;
    // 权威是**原生控件**：用户点标签只改 TCM_GETCURSEL，t->cur 记的是“我们上次设的值”。
    // 早期这里直接答 t->cur + 1 ⇒ 真点完标签读恒 1（实测 .build/probe/verify_tab.ps1
    // 的 T1/T2 两行：clicks 在长、sel 不动）。判决式一律回问原生，顺手把缓存对齐。
    sel = SendMessageW(t->hwnd, TCM_GETCURSEL, 0, 0);
    if (sel < 0) return 0;          // 一页未选 = VB6 的 0
    t->cur = (int)sel;
    return (int32_t)sel + 1;        // 1 基
}

void vb6_TabStrip_SetSelectedIndex(void* hwnd, int32_t idx1) {
    Vb6TabStrip* t = tsFind(hwnd);
    int at = idx1 - 1;
    if (!t) return;
    if (at < 0 || at >= t->count) return;
    t->cur = at;
    SendMessageW(t->hwnd, TCM_SETCURSEL, (WPARAM)at, 0);
}

// 无头判据助手：像用户真点了一下标签那样，让**原生**发一条 TCN_SELCHANGE 到父窗。
// 只发通知不改状态的是"假证据"（UpDown 那一批的教训）；这里先 TCM_SETCURSEL 再发通知，
// 事件处理器跑起来时 SelectedItem 已经是新值 —— 与真手势的顺序一致。
int32_t vb6_TabStrip_SimSelect(void* hwnd, int32_t idx1) {
    Vb6TabStrip* t = tsFind(hwnd);
    HWND parent;
    NMHDR nm;
    int at = idx1 - 1;
    if (!t || at < 0 || at >= t->count) return 0;
    t->cur = at;
    SendMessageW(t->hwnd, TCM_SETCURSEL, (WPARAM)at, 0);
    parent = GetParent(t->hwnd);
    if (parent) {
        nm.hwndFrom = t->hwnd;
        nm.idFrom = (UINT_PTR)GetDlgCtrlID(t->hwnd);
        nm.code = (UINT)VU_TCN_SELCHANGE;
        SendMessageW(parent, WM_NOTIFY, nm.idFrom, (LPARAM)&nm);
    }
    return t->cur + 1;
}

// WM_NOTIFY 换算：原生把 TCN_SELCHANGE 发到父窗，code = **-551**（TCN_FIRST = 0U-550U）。
// 返回 1 = 这条是"选中项真的变了"，调用方该发 Click。
// 同时把 cur 与原生对齐（用户用键盘/鼠标切页时 RTL 那份表要跟上）。
int32_t vb6_TabStrip_OnNotify(void* hwnd, int32_t code, void* lParam) {
    Vb6TabStrip* t = tsFind(hwnd);
    NMHDR* p = (NMHDR*)lParam;
    int sel;
    if (!t || !p) return 0;
    if (code != VU_TCN_SELCHANGE && code != VU_TCN_SELCHANGING) return 0;
    if (p->hwndFrom != t->hwnd) return 0;
    if (code == VU_TCN_SELCHANGING) return 0;      // 未真的变；Click 不发
    sel = (int)SendMessageW(t->hwnd, TCM_GETCURSEL, 0, 0);
    if (sel < 0 || sel >= t->count) return 0;
    t->cur = sel;
    return 1;
}

#endif // _WIN32
