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
// 账 #611 / 台账 §B12（D14）: 控件画字一律问**那位主人**（#204 的字体存储槽 vb6_ControlFont），
// 不许自己 `GetStockObject(DEFAULT_GUI_FONT)` —— 那是这台机器的 GDI 缺省（本机实测 SimSun -12），
// 而同一个 TabStrip 的窗口字体是 MS Sans Serif 8.25pt（-11）⇒ 画出来的标签与量出来的
// 标签尺寸会对不上（自绘气泡测宽尤其明显）。本文件原先两处就是这么长的。
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
    HWND      tipWnd;       // 自建 track tooltip（ShowTips 气泡，见 tsTipShow）
    int       tipShown;     // 正在弹的 1 基页序号；0 = 没弹
    int       tipTracking;  // 已挂 TME_LEAVE
    WNDPROC   prevProc;     // 子类链的下一层（原窗口过程）
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
// 重灌前先收气泡：项会被 TCM_DELETEALLITEMS 清掉，正弹着的那一枚就指向空气了。
static void tsTipHide(Vb6TabStrip* t);

static void tsRepopulate(Vb6TabStrip* t) {
    TCITEMW it;
    int i;
    if (!t || !t->hwnd) return;
    tsTipHide(t);
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

// ===================== ShowTips 气泡（C30-2 收尾之三）=====================
// 为何自建：原生 SysTabControl32 在 TCS_TOOLTIPS 下的气泡文本**恒等于项文本**
// （TCITEM.pszText），拿不到 Tabs(i).ToolTipText 这一枚。而托 comctl32 的 tooltip 去改
// 也不行（实测两轮）：TTM_UPDATETIPTEXT+TTM_TRACKACTIVATE 那条路气泡压根不显；想拿
// SetWindowTextW 打个标记也不行（tooltip 的类过程把 WM_SETTEXT 吃掉了，标题永远是空）。
// ⇒ 干脆自绘一枚无激活弹出窗：类名 VB6_TabTipWnd、标题就是气泡文本（可跨进程 WM_GETTEXT 读），
// 外观用系统提示色（COLOR_INFOBK/COLOR_INFOTEXT）。停在哪个页就弹哪个页的 ToolTipText，
// **空则不弹**（VB6 口径：没写过 ToolTipText 的页没有气泡；不拿标题兜底）。
#define VB6_TS_TIP_CLASSW L"VB6_TabTipWnd"

static LRESULT CALLBACK tsTipProc(HWND h, UINT m, WPARAM wp, LPARAM lp);
static LRESULT CALLBACK tsTipWndProc(HWND h, UINT m, WPARAM wp, LPARAM lp);

static void tsTipHide(Vb6TabStrip* t) {
    if (!t) return;
    if (!t->tipShown) return;
    if (t->tipWnd) ShowWindow(t->tipWnd, SW_HIDE);
    t->tipShown = 0;
}

static void tsTipEnsure(Vb6TabStrip* t) {
    WNDCLASSEXW wc;
    HINSTANCE hi;
    if (!t || !t->hwnd || t->tipWnd) return;
    hi = (HINSTANCE)GetModuleHandleW(NULL);
    ZeroMemory(&wc, sizeof(wc));
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = tsTipWndProc;
    wc.hInstance = hi;
    if (!GetClassInfoExW(hi, VB6_TS_TIP_CLASSW, &wc)) {   // 已注册过就不重复注册
        ZeroMemory(&wc, sizeof(wc));
        wc.cbSize = sizeof(wc);
        wc.lpfnWndProc = tsTipWndProc;
        wc.hInstance = hi;
        wc.hCursor = LoadCursorW(NULL, (LPCWSTR)IDC_ARROW);
        wc.lpszClassName = VB6_TS_TIP_CLASSW;
        if (!RegisterClassExW(&wc)) return;
    }
    t->tipWnd = CreateWindowExW(WS_EX_TOPMOST | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
                                VB6_TS_TIP_CLASSW, L"",
                                WS_POPUP | WS_BORDER,
                                0, 0, 10, 10, (HWND)t->hwnd, NULL, hi, NULL);
}

// 自绘：系统提示色底 + 单行文本（VB6 气泡也是单行）。
static LRESULT CALLBACK tsTipWndProc(HWND h, UINT m, WPARAM wp, LPARAM lp) {
    switch (m) {
    case WM_ERASEBKGND:
        return 1;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        RECT rc;
        HDC dc = BeginPaint(h, &ps);
        wchar_t buf[512];
        int n;
        GetClientRect(h, &rc);
        {
            HBRUSH br = CreateSolidBrush(GetSysColor(COLOR_INFOBK));
            FillRect(dc, &rc, br);
            DeleteObject(br);
        }
        n = GetWindowTextW(h, buf, 512);
        if (n > 0) {
            // 与 WM_SETFONT 发进窗口的那张**同一个**句柄 —— 但只许通过出口 vb6_ControlFont 问
            //（D11：全 src/rtl 里裸问 WM_GETFONT 只许剩出口自己那一行）。它内部先问窗口
            // 再问存储槽（#204），两问皆空才让你用 DC 自带字体。
            HFONT f = vb6_ControlFont(h);
            HFONT of = f ? (HFONT)SelectObject(dc, f) : NULL;
            rc.left += 4; rc.right -= 4; rc.top += 2; rc.bottom -= 2;
            SetBkMode(dc, TRANSPARENT);
            SetTextColor(dc, GetSysColor(COLOR_INFOTEXT));
            DrawTextW(dc, buf, n, &rc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
            if (of) SelectObject(dc, of);
        }
        EndPaint(h, &ps);
        return 0;
    }
    case WM_MOUSEACTIVATE:
        return MA_NOACTIVATE;
    }
    return DefWindowProcW(h, m, wp, lp);
}

static void tsTipShow(Vb6TabStrip* t, int i0) {
    RECT rc, trc;
    POINT pt;
    HDC dc;
    HFONT of;
    int w, ht;
    if (!t || !t->hwnd || i0 < 0 || i0 >= t->count) return;
    if (!t->tip || !t->tip[i0] || !t->tip[i0][0]) { tsTipHide(t); return; }
    tsTipEnsure(t);
    if (!t->tipWnd) return;
    rc.left = rc.top = rc.right = rc.bottom = 0;
    if (!SendMessageW((HWND)t->hwnd, TCM_GETITEMRECT, (WPARAM)i0, (LPARAM)&rc)) return;
    SetWindowTextW(t->tipWnd, t->tip[i0]);
    // 量文本（DT_CALCRECT）后按内容定尺寸；贴在该页左下角，跟 VB6 观感一致。
    // 字体必须与标签**同源**（否则气泡框比文字大一圈/小一圈）—— 走同一个出口 vb6_ControlFont（D11）。
    dc = GetDC(t->tipWnd);
    of = vb6_ControlFont((HWND)t->hwnd);
    of = of ? (HFONT)SelectObject(dc, of) : NULL;
    SetRect(&trc, 0, 0, 480, 0);
    DrawTextW(dc, t->tip[i0], -1, &trc, DT_CALCRECT | DT_SINGLELINE | DT_NOPREFIX);
    w = (trc.right - trc.left) + 8;
    ht = (trc.bottom - trc.top) + 4;
    SelectObject(dc, of);
    ReleaseDC(t->tipWnd, dc);    pt.x = rc.left; pt.y = rc.bottom;
    ClientToScreen((HWND)t->hwnd, &pt);
    SetWindowPos(t->tipWnd, HWND_TOPMOST, pt.x, pt.y, w, ht,
                 SWP_NOACTIVATE | SWP_SHOWWINDOW);
    InvalidateRect(t->tipWnd, NULL, TRUE);
    t->tipShown = i0 + 1;
}

// 命中：逐项比矩形（不用 TCM_HITTEST —— 它要“确实落在某页的绘制区”才算，
// 页间分隔线上会回 -1；自己也占着一个 TCHITTESTINFO 结构，没必要）。
static int tsItemAt(Vb6TabStrip* t, int x, int y) {
    RECT rc;
    int i;
    if (!t || !t->hwnd) return -1;
    for (i = 0; i < t->count; i++) {
        rc.left = rc.top = rc.right = rc.bottom = 0;
        if (!SendMessageW((HWND)t->hwnd, TCM_GETITEMRECT, (WPARAM)i, (LPARAM)&rc)) continue;
        if (x >= rc.left && x < rc.right && y >= rc.top && y < rc.bottom) return i;
    }
    return -1;
}

static void tsTipHover(Vb6TabStrip* t, int x, int y) {
    int i0;
    if (!t) return;
    if (!t->showTips) { tsTipHide(t); return; }
    i0 = tsItemAt(t, x, y);
    if (i0 < 0) { tsTipHide(t); return; }
    if (t->tipShown == i0 + 1) return;   // 同页不重弹（否则每动一像素闪一下）
    tsTipShow(t, i0);
}

// 子类：**自己链**下一层，不占 cgen 的 VB6_OrigProc 属性槽。
// 理由（已实测，账面 #185）：vb6_InstallControlSubclass 见槽上有值就直接 return，
// 而 cgen 是在 Init **之后**才装控件子类的 —— RTL 若先占坑，控件事件会整段哑掉。
static LRESULT CALLBACK tsTipProc(HWND h, UINT m, WPARAM wp, LPARAM lp) {
    Vb6TabStrip* t = tsFind((void*)h);
    WNDPROC prev = t ? t->prevProc : NULL;
    if (t) {
        switch (m) {
        case WM_MOUSEMOVE:
            if (!t->tipTracking) {
                TRACKMOUSEEVENT tme;
                ZeroMemory(&tme, sizeof(tme));
                tme.cbSize = sizeof(tme); tme.dwFlags = TME_LEAVE; tme.hwndTrack = h;
                if (TrackMouseEvent(&tme)) t->tipTracking = 1;
            }
            tsTipHover(t, (int)(short)LOWORD(lp), (int)(short)HIWORD(lp));
            break;
        case WM_MOUSELEAVE:
            t->tipTracking = 0;
            tsTipHide(t);
            break;
        case WM_LBUTTONDOWN:
        case WM_CAPTURECHANGED:
        case WM_DESTROY:
            tsTipHide(t);
            break;
        }
    }
    if (prev) return CallWindowProcW(prev, h, m, wp, lp);
    return DefWindowProcW(h, m, wp, lp);
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
        st = hotTrack ? (st | TCS_HOTTRACK) : (st & ~TCS_HOTTRACK);
        st &= ~TCS_TOOLTIPS;   // 气泡由 RTL 自建（tsTipShow）：原生那条只会说项文本
        SetWindowLongW(h, GWL_STYLE, st);
    }
    if (!t->prevProc) {
        WNDPROC old = (WNDPROC)SetWindowLongPtrW(h, GWLP_WNDPROC, (LONG_PTR)tsTipProc);
        if (old && old != tsTipProc) t->prevProc = old;
    }
    tsTipEnsure(t);   // 气泡窗随控件预建（隐藏态），首次悬停只改文本/位置
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
    {
        LONG st = GetWindowLongW((HWND)t->hwnd, GWL_STYLE);
        st &= ~TCS_TOOLTIPS;   // 同上：这层气泡永远归 RTL 管
        SetWindowLongW((HWND)t->hwnd, GWL_STYLE, st);
    }
    if (!t->showTips) tsTipHide(t);
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
    // C30-2 修：删掉的页在**当前选中页之前**时，选中项要跟着往前顶一格，否则用户手上
    // 那一页会“跳”到下一项（实测：3 页里选中第 3 页，删第 1 页后旧代码把第 3 页变成第 2 页
    // 却仍然选中第 2 个下标 → 选中项换人）。删的就是当前页时**不动下标**：顶上来那一页
    // 就地接管选中（＝ VB6 的“删除后选其下一位”）；删的是末页则钳到新的末项。
    if (at < t->cur) t->cur--;
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

// C30-2: Tabs(i).Key 写 / Tabs(i).Image 读写（VB6 的 Key 空串 = 无键；Image = ImageList 索引，
// -1 = 无图标）。三枚都是“改自存 + 重灌原生项”，不回写控件自己（原生 tab 没有 Key 概念）。
void vb6_TabStrip_SetTabKey(void* hwnd, int32_t idx1, void* bstr) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    const wchar_t* s = (const wchar_t*)bstr;
    if (!t || i < 0 || i >= t->count) return;
    free(t->key[i]);
    t->key[i] = (s && s[0]) ? tsWcsDup(s) : NULL;
    tsRepopulate(t);
}

int32_t vb6_TabStrip_GetTabImage(void* hwnd, int32_t idx1) {    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    if (!t || i < 0 || i >= t->count) return -1;
    return t->image[i];
}

void vb6_TabStrip_SetTabImage(void* hwnd, int32_t idx1, int32_t iImage) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    if (!t || i < 0 || i >= t->count) return;
    t->image[i] = (iImage >= 0) ? iImage : -1;
    tsRepopulate(t);
}

// C30-2: Tabs(i).ToolTipText —— VB6 的 TabStrip 在 ShowTips=True 时气泡里显示的是**这一枚**，
// 而原生 SysTabControl32 的 TCS_TOOLTIPS 只会在气泡里显示**项文本**（TCITEM.pszText）。
// 实测口径：这里只保证**读写往返一致**（业务代码取值/赋值不回错、不失真），不谎称
// “已经能改原生气泡文本”—— 真要让气泡说 ToolTipText 得自建 track tooltip，那是另一刀。
void* vb6_TabStrip_GetTabToolTipText(void* hwnd, int32_t idx1) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    if (!t || i < 0 || i >= t->count) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(t->tip[i] ? t->tip[i] : L"");
}

void vb6_TabStrip_SetTabToolTipText(void* hwnd, int32_t idx1, void* bstr) {
    Vb6TabStrip* t = tsFind(hwnd);
    int i = idx1 - 1;
    const wchar_t* s = (const wchar_t*)bstr;
    if (!t || i < 0 || i >= t->count) return;
    free(t->tip[i]);
    t->tip[i] = (s && s[0]) ? tsWcsDup(s) : NULL;
    tsRepopulate(t);
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
