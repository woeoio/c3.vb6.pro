// vb6forms_statusbar.c — VB6 StatusBar 控件 (ai/内置控件/Toolbar 与 StatusBar 控件.md)
//
// 复刻口径: 用 comctl32 的 msctls_status32 复刻 VB6 StatusBar 的等价行为,
// **不加载 mscomctl.ocx** (本机未注册, 且 32 位 inproc OCX 无法进 x64 进程)。
//
// 属性表 (照文档全量实现):
//   Align      0=sbrAlignNone 1=上 2=下(默认) 3=左 4=右
//   Style      0=sbrNormal(多面板, 默认) 1=sbrSimple(单格, 用 SimpleText)
//   SimpleText 简单模式下的文字
//   Panels     面板集合 (Panel 成员: Key / Text / Width / MinWidth /
//              ToolTipText / AutoSize / Style)
// 方法:
//   Panels.Add([index], [key], [text])   —— 返回 1 基 Index
//   Panels.Remove(index 或 key)
//   Panels.Clear
// AutoSize: 0=sbrFixed 1=sbrSpring 2=sbrContents
// Panel.Style: 0=sbrText 1=sbrCaps 2=sbrNum 5=sbrTime 6=sbrDate
//
// 实现要点:
//   - msctls_status32 无窗口子类化也能画, 但 **时间/日期面板需要自己走时钟**:
//     SDK 10.0.19041.0 的 commctrl.h 里**没有** SBT_CAPS/SBT_NUMLOCK/SBT_TIME/SBT_DATE
//     (实测只剩 SBT_TOOLTIPS/OWNERDRAW/NOBORDERS/POPOUT/RTLREADING/NOTABPARSING),
//     所以这四个"系统自动显示"的面板得由这里自己算文本; 时钟用 SetTimer + 子类化。
//   - comctl32 只认**下标**, VB6 的 Key/Text 是字符串 → 额外维护一份平行表,
//     下标与面板顺序严格同步 (Add 插到中间 / Remove 后搬动)。
//   - VB6 的面板宽度是**缇**吗? 不是 —— StatusBar 的 Width/MinWidth 在 VB6 里就是像素
//     (设计期属性页按像素给), 这里直接当像素用。
//   - 布局口径: 固定宽面板各占 Width; 弹簧面板(MinWidth 为下限)瓜分剩余空间;
//     随内容面板(sbrContents)按文字实测宽。全部换算成 SB_SETPARTS 的右边界数组。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#include <time.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6rtl_bstr.h"

#ifdef _WIN32

// ===================== VB6 枚举常量 =====================
#define VB6_SBR_ALIGN_NONE  0
#define VB6_SBR_ALIGN_TOP   1
#define VB6_SBR_ALIGN_BOT   2
#define VB6_SBR_ALIGN_LEFT  3
#define VB6_SBR_ALIGN_RIGHT 4

#define VB6_SBR_NORMAL   0
#define VB6_SBR_SIMPLE   1

#define VB6_SBR_FIXED     0
#define VB6_SBR_SPRING    1
#define VB6_SBR_CONTENTS  2

#define VB6_SBR_TEXT  0
#define VB6_SBR_CAPS  1
#define VB6_SBR_NUM   2
#define VB6_SBR_TIME  5
#define VB6_SBR_DATE  6

#define VB6_SB_TIMER_ID 0x7FF1

// ===================== 实例与条目 =====================

typedef struct {
    wchar_t* key;        // 未命名则 NULL
    wchar_t* text;       // **用户给的** Text (设计期或运行期, 原样留住)
    wchar_t* shown;      // **实际显示**的文本: 0/Caps/Num/Time/Date 面板的现算值,
                         // 普通面板为 NULL (表示"就用 text")。
                         // 两者必须分开 —— 早期版本把自动值盖进 text, 于是
                         // `Style=sbrNum` 的面板设计期 Text="Tip" 读回来变成空串
                         // (.frm 里 `Panels(2) = "Tip"` 就是设计期 Text, 不能吞)。
    wchar_t* tip;        // ToolTipText (仅存起来, msctls_status32 的 SBARS_TOOLTIPS
                         // 需要 SB_SETTEXTLENGTH 那套 tooltip, 这里不实现激活)
    int      width;      // Width (像素, <=0 = 未指定)
    int      minWidth;   // MinWidth (弹簧面板的下限)
    int      autoSize;   // 0 fixed / 1 spring / 2 contents
    int      style;      // 0 text / 1 caps / 2 num / 5 time / 6 date
} Vb6PanelEntry;

typedef struct {
    HWND           hwnd;   // msctls_status32 窗口
    Vb6PanelEntry* ents;
    int            count;
    int            cap;
    int            style;       // StatusBar.Style (0/1)
    int            align;       // StatusBar.Align
    wchar_t*       simpleText;  // Style=1 时显示
    UINT_PTR       timerId;     // 时间/日期面板的时钟; 0 = 没装
} Vb6StatusBar;

// 从 hwnd 找回实例。生成的 C 里槽就是 HWND, 但窗口创建后我们还要挂一堆面板状态,
// 所以**槽与 HWND 双写**: 槽里放 Vb6StatusBar*, 同时用 Prop 存一份给子类化回调用。
static void vb6_SbStore(Vb6StatusBar* sb) {
    if (!sb || !sb->hwnd) return;
    SetPropW(sb->hwnd, L"VB6_StatusBar", (HANDLE)sb);
}

static Vb6StatusBar* vb6_SbFromHwnd(void* hwnd) {
    if (!hwnd) return NULL;
    return (Vb6StatusBar*)GetPropW((HWND)hwnd, L"VB6_StatusBar");
}

static void SbEntClear(Vb6PanelEntry* e) {
    if (!e) return;
    free(e->key);   e->key = NULL;
    free(e->text);  e->text = NULL;
    free(e->shown); e->shown = NULL;
    free(e->tip);   e->tip = NULL;
}

static void SbEntsFree(Vb6StatusBar* sb) {
    if (!sb) return;
    for (int i = 0; i < sb->count; i++) SbEntClear(&sb->ents[i]);
    free(sb->ents);
    sb->ents = NULL; sb->count = 0; sb->cap = 0;
}

static int SbReserve(Vb6StatusBar* sb, int need) {
    if (need <= sb->cap) return 0;
    int ncap = sb->cap ? sb->cap : 8;
    while (ncap < need) ncap *= 2;
    Vb6PanelEntry* ne = (Vb6PanelEntry*)realloc(sb->ents, (size_t)ncap * sizeof(Vb6PanelEntry));
    if (!ne) return -1;
    for (int i = sb->cap; i < ncap; i++) {
        ne[i].key = NULL; ne[i].text = NULL; ne[i].shown = NULL; ne[i].tip = NULL;
        ne[i].width = 0; ne[i].minWidth = 0; ne[i].autoSize = VB6_SBR_FIXED;
        ne[i].style = VB6_SBR_TEXT;
    }
    sb->ents = ne; sb->cap = ncap;
    return 0;
}

static int SbFindKey(Vb6StatusBar* sb, const wchar_t* key) {
    if (!key || !key[0]) return -1;
    for (int i = 0; i < sb->count; i++) {
        if (sb->ents[i].key && _wcsicmp(sb->ents[i].key, key) == 0) return i;
    }
    return -1;
}

// ===================== 文本测算 =====================
// sbrContents 面板要按实测文字宽排版; msctls_status32 只告诉你右边界,
// 所以这里自己拿控件字体测一遍。
static int SbMeasureText(HWND hw, const wchar_t* txt) {
    if (!txt || !txt[0]) return 0;
    HDC hdc = GetDC(hw);
    if (!hdc) return 0;
    // 账 #205: 状态条是 RTL 自己注册的窗口类, 裸问窗口永远拿不到那张字体 ——
    // 实测同一串文字在 8pt 与 20pt 两枚状态条上量出**同一个宽**(140)，改了字号也不动。
    // 字体只从 vb6_ControlFont 那一处问（#200 立的出口，#204 把创建期下发的那张存进同槽位）。
    HFONT font = vb6_ControlFont(hw);
    HFONT old = font ? (HFONT)SelectObject(hdc, font) : NULL;
    SIZE sz = {0, 0};
    GetTextExtentPoint32W(hdc, txt, (int)wcslen(txt), &sz);
    if (old) SelectObject(hdc, old);
    ReleaseDC(hw, hdc);
    return sz.cx;
}

// ===================== 面板文本 =====================
// Style=1/2/5/6 是"系统自己显示", 显示的文本不由用户给。这里现算一份塞进
// ents[i].**shown**, **绝不动 e->text** —— 用户给的文本要原样留着读回来。
static void SbFillAutoText(Vb6StatusBar* sb, int i) {
    Vb6PanelEntry* e = &sb->ents[i];
    int kind = e->style;
    if (kind != VB6_SBR_CAPS && kind != VB6_SBR_NUM &&
        kind != VB6_SBR_TIME && kind != VB6_SBR_DATE) return;
    free(e->shown); e->shown = NULL;

    wchar_t buf[64];
    switch (kind) {
        case VB6_SBR_CAPS: {
            // Caps Lock 开 = 显示 "CAPS", 关 = 空串 (VB6 就是这样, 不是 "sbrCaps" 字眼)。
            int on = (GetKeyState(VK_CAPITAL) & 1) != 0;
            if (on) wcsncpy(buf, L"CAPS", 32);
            else    buf[0] = L'\0';
            buf[32] = L'\0';
            break;
        }
        case VB6_SBR_NUM: {
            int on = (GetKeyState(VK_NUMLOCK) & 1) != 0;
            const wchar_t* p = on ? L"NUM" : L"";
            wcsncpy(buf, p, 32); buf[32] = L'\0';
            break;
        }
        case VB6_SBR_TIME: {
            SYSTEMTIME st; GetLocalTime(&st);
            swprintf(buf, 64, L"%02d:%02d:%02d", st.wHour, st.wMinute, st.wSecond);
            break;
        }
        case VB6_SBR_DATE:
        default: {
            SYSTEMTIME st; GetLocalTime(&st);
            // VB6 的 sbrDate 用地区短日期, 这里用"Y/M/D"这个与区域无关的稳定形态,
            // 免得夹具断言跟着系统区域设置跑。
            swprintf(buf, 64, L"%04d/%02d/%02d", st.wYear, st.wMonth, st.wDay);
            break;
        }
    }
    if (buf[0]) e->shown = _wcsdup(buf);
}

// 面板"当前该显示什么": 自动显示优先, 否则用用户文本。
static const wchar_t* SbDisplayText(Vb6PanelEntry* e) {
    if (e->shown) return e->shown;
    return e->text ? e->text : L"";
}

// 把某面板的当前文本刷进 msctls_status32。
static void SbPushText(Vb6StatusBar* sb, int i) {
    SendMessageW(sb->hwnd, SB_SETTEXTW, (WPARAM)(i + 1),
                 (LPARAM)SbDisplayText(&sb->ents[i]));
}

// ===================== 布局 =====================
// 算出每个面板的像素宽, 累加成 SB_SETPARTS 的右边界数组。
// 返回需要的数组长度 (<= count), 失败返回 -1。
static int SbLayout(Vb6StatusBar* sb, int* outOffsets, int outCap) {
    if (!sb->count) return 0;
    if (outCap < sb->count) return -1;

    RECT rc;
    GetClientRect(sb->hwnd, &rc);
    int clientW = rc.right - rc.left;
    if (clientW < 0) clientW = 0;

    // 第一遍: 各面板"先要"的宽度。弹簧面板先按 MinWidth 占位。
    int* want = (int*)calloc((size_t)sb->count, sizeof(int));
    int  springs = 0;
    int  fixedTotal = 0;
    for (int i = 0; i < sb->count; i++) {
        Vb6PanelEntry* e = &sb->ents[i];
        if (e->autoSize == VB6_SBR_SPRING) {
            want[i] = e->minWidth > 0 ? e->minWidth : 0;
            springs++;
        } else if (e->autoSize == VB6_SBR_CONTENTS) {
            want[i] = SbMeasureText(sb->hwnd, SbDisplayText(e));
        } else {
            want[i] = e->width > 0 ? e->width : SbMeasureText(sb->hwnd, SbDisplayText(e));
            fixedTotal += want[i];
        }
    }

    // 弹簧面板瓜分剩余空间。
    int remaining = clientW - fixedTotal;
    if (remaining < 0) remaining = 0;
    int perSpring = springs ? remaining / springs : 0;
    int leftover   = springs ? remaining - perSpring * springs : 0;

    int acc = 0;
    for (int i = 0; i < sb->count; i++) {
        Vb6PanelEntry* e = &sb->ents[i];
        int take = want[i];
        if (e->autoSize == VB6_SBR_SPRING) {
            take = want[i] + perSpring;
            if (leftover > 0) { take += 1; leftover--; }   // 除不尽的零头给第一个弹簧
            if (take < (e->minWidth > 0 ? e->minWidth : 0)) take = e->minWidth > 0 ? e->minWidth : 0;
        }
        acc += take;
        outOffsets[i] = acc;
    }
    free(want);
    return sb->count;
}

static void SbApplyParts(Vb6StatusBar* sb) {
    if (!sb->hwnd || sb->count <= 0) return;
    int* offs = (int*)calloc((size_t)sb->count, sizeof(int));
    if (!offs) return;
    int n = SbLayout(sb, offs, sb->count);
    if (n > 0) SendMessageW(sb->hwnd, SB_SETPARTS, (WPARAM)n, (LPARAM)offs);
    free(offs);
}

// 有没有面板要跟着时钟走。
static int SbHasClock(Vb6StatusBar* sb) {
    for (int i = 0; i < sb->count; i++) {
        int s = sb->ents[i].style;
        if (s == VB6_SBR_TIME || s == VB6_SBR_DATE) return 1;
    }
    return 0;
}

static void SbRefreshClock(Vb6StatusBar* sb) {
    for (int i = 0; i < sb->count; i++) {
        int s = sb->ents[i].style;
        if (s != VB6_SBR_TIME && s != VB6_SBR_DATE) continue;
        SbFillAutoText(sb, i);
        SbPushText(sb, i);
    }
}

static void SbUpdateClock(Vb6StatusBar* sb) {
    if (!sb) return;
    int need = SbHasClock(sb);
    if (need && !sb->timerId) {
        sb->timerId = SetTimer(sb->hwnd, VB6_SB_TIMER_ID, 1000, NULL);
    } else if (!need && sb->timerId) {
        KillTimer(sb->hwnd, sb->timerId);
        sb->timerId = 0;
    }
}

// ===================== 自注册窗口类 (comctl32 不注册时) =====================
// comctl32 在本机**两个版本都不注册** msctls_status32: v5.82 (无清单) 与 v6.0
// (挂 v6 清单) 实测都一样 —— 用任何 ICC_ 标志位调 InitCommonControlsEx 都补不上,
// 连 dwICC=0xFFFFFFFF 全开也不行, CreateWindowExW 于是静默失败
// (GetLastError()==1407, 类未注册)。所以这里自己注册一个**同名的真**状态栏类:
// SB_* 的消息契约照 comctl32 实现, 绘制也自己来 —— 反正 VB6 StatusBar 的语义
// (Key/Text/AutoSize/时间日期面板) 本来就是 comctl32 没有的那部分。
// 若 comctl32 哪天真注册了 (别的机器 / 别的 comctl32 版本), 就让它赢, 这里跳过。
// SDK 10.0.19041.0 的 winuser.h / commctrl.h 也没有这两个 (BDR_/BF_ 与
// SB_GETMINHEIGHT 都没给全), 按 comctl32 的既有约定补上。
#ifndef BF_ADJUSTRECT
#define BF_ADJUSTRECT    0x1000   // 画完边框再把矩形收进去
#endif
#ifndef SB_GETMINHEIGHT
#define SB_GETMINHEIGHT  (WM_USER+5)   // WM_USER+5 在 SB_* 里是空号 (SDK 只定义到 +20)
#endif

#define C3_SBT_LEFT      0x0000
#define C3_SBT_CENTER    0x0002
#define C3_SBT_RIGHT     0x0003
#define C3_SBT_CAPS      0x0004   // SDK 10.0.19041.0 的 commctrl.h 没有这几个
#define C3_SBT_NUMLOCK   0x0008
#define C3_SBT_RECCOUNT  0x0010

typedef struct {
    int       cap;      // 下面几个数组的实际容量 (只增不缩)
    int       count;    // 当前格数 (由 SB_SETPARTS 决定)
    int*      rights;   // 每格右边界 (像素, 相对客户区左边)
    wchar_t** texts;
    int*      flags;    // SBT_* 每格格式
    HICON*    icons;
    int       minHeight; // SB_SETMINHEIGHT
    int       simple;    // SB_SIMPLE 打开
} Vb6SbWin;

typedef struct {
    UINT   uProcessId;
    HWND   hwndFrom;
    LPARAM wParam;
    LPARAM lParam;
    UINT   uFlags;
    DWORD  dwLength;
} C3SbGetText;

static Vb6SbWin* C3SbWin(HWND hw) {
    return (Vb6SbWin*)GetPropW(hw, L"VB6_SBWin");
}

// 索引可能越过当前格数 (SB_SETTEXT 直接给 1 基下标), 这里按需**精确**扩容,
// 绝不碰 count —— count 的权威是 SB_SETPARTS。
static int C3SbWinReserve(Vb6SbWin* w, int need) {
    if (need <= w->cap) return 0;
    int n = w->cap ? w->cap : 8;
    while (n < need) n += 8;
    int* r = (int*)realloc(w->rights, (size_t)n * sizeof(int));
    if (!r) return -1;
    w->rights = r;
    wchar_t** t = (wchar_t**)realloc(w->texts, (size_t)n * sizeof(wchar_t*));
    if (!t) return -1;
    w->texts = t;
    int* f = (int*)realloc(w->flags, (size_t)n * sizeof(int));
    if (!f) return -1;
    w->flags = f;
    HICON* ic = (HICON*)realloc(w->icons, (size_t)n * sizeof(HICON));
    if (!ic) return -1;
    w->icons = ic;
    for (int i = w->cap; i < n; i++) {
        w->texts[i] = NULL; w->flags[i] = C3_SBT_LEFT; w->icons[i] = NULL;
    }
    w->cap = n;
    return 0;
}

static void C3SbWinFree(Vb6SbWin* w) {
    if (!w) return;
    for (int i = 0; i < w->cap; i++) {
        free(w->texts[i]);
        if (w->icons[i]) DestroyIcon(w->icons[i]);
    }
    free(w->rights); free(w->texts); free(w->flags); free(w->icons); free(w);
}

static void C3SbPaint(HWND hw) {
    Vb6SbWin* w = C3SbWin(hw);
    RECT rc;
    GetClientRect(hw, &rc);
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hw, &ps);
    if (!hdc) return;
    HBRUSH bk = GetSysColorBrush(COLOR_BTNFACE);
    // 账 #205: 画的那一遍与量的那一遍必须问同一处，否则格宽按一张字画按另一张。
    HFONT font = vb6_ControlFont(hw);
    HFONT oldFont = font ? (HFONT)SelectObject(hdc, font) : NULL;
    int h = rc.bottom - rc.top;
    int n = w ? w->count : 0;
    if (w && w->simple) n = 1;
    for (int i = 0; i < n; i++) {
        RECT pr;
        pr.left  = (i == 0) ? 0 : w->rights[i - 1];
        pr.right = (i + 1 < n) ? w->rights[i] : rc.right;
        pr.top = 0; pr.bottom = h;
        if (pr.right <= pr.left) continue;
        int flags = w->flags[i];
        int noBorders = (flags & SBT_NOBORDERS) || (w->simple != 0);
        if (!noBorders)
            DrawEdge(hdc, &pr, BDR_SUNKENOUTER | BDR_RAISEDINNER,
                     BF_RECT | BF_ADJUSTRECT);
        FillRect(hdc, &pr, bk);
        InflateRect(&pr, -2, -2);
        if (w->icons[i] && !w->simple)
            DrawIconEx(hdc, pr.left, pr.top, w->icons[i], 16, 16, 0, NULL, DI_NORMAL);
        const wchar_t* t = w->texts[i];
        if (t && *t) {
            UINT dt = DT_SINGLELINE | DT_VCENTER | DT_NOPREFIX | DT_END_ELLIPSIS;
            if (flags & C3_SBT_CENTER)                                dt |= DT_CENTER;
            else if ((flags & 0x000F) >= C3_SBT_RIGHT)                dt |= DT_RIGHT;
            else                                                      dt |= DT_LEFT;
            DrawTextW(hdc, t, (int)wcslen(t), &pr, dt);
        }
    }
    if (oldFont) SelectObject(hdc, oldFont);
    EndPaint(hw, &ps);
}

static LRESULT CALLBACK vb6_StatusBarWndProc(HWND hw, UINT msg,
                                              WPARAM wp, LPARAM lp) {
    Vb6SbWin* w = C3SbWin(hw);
    switch (msg) {
    case WM_NCCREATE:
        SetPropW(hw, L"VB6_SBWin", (HANDLE)calloc(1, sizeof(Vb6SbWin)));
        return DefWindowProcW(hw, msg, wp, lp);
    case WM_NCDESTROY:
        C3SbWinFree(C3SbWin(hw));
        RemovePropW(hw, L"VB6_SBWin");
        return DefWindowProcW(hw, msg, wp, lp);
    case WM_PAINT:
        C3SbPaint(hw);
        return 0;
    case WM_ERASEBKGND:
        return 1;
    case WM_SIZE:
        InvalidateRect(hw, NULL, TRUE);
        return DefWindowProcW(hw, msg, wp, lp);
    case WM_SETFONT:
        InvalidateRect(hw, NULL, TRUE);
        return DefWindowProcW(hw, msg, wp, lp);
    case SB_SETTEXTW: {
        int idx = (int)(intptr_t)LOWORD(wp);
        int flags = (int)(intptr_t)HIWORD(wp);
        const wchar_t* t = (const wchar_t*)lp;
        if (!w) return 0;
        // SB_SIMPLEID (0xff) = 简单模式那唯一一格
        if (idx == SB_SIMPLEID) { idx = 0; flags |= SBT_NOBORDERS; }
        if (idx < 0 || C3SbWinReserve(w, idx + 1) != 0) return 0;
        free(w->texts[idx]);
        w->texts[idx] = (t && *t) ? _wcsdup(t) : NULL;
        w->flags[idx] = flags;
        InvalidateRect(hw, NULL, TRUE);
        return 0;
    }
    case SB_GETTEXTLENGTHW: {
        int idx = (int)(intptr_t)LOWORD(wp);
        if (!w || idx < 0 || idx >= w->count || !w->texts[idx]) return 0;
        return (LRESULT)(wcslen(w->texts[idx]) + 1);
    }
    case SB_GETTEXTW: {
        int idx = (int)(intptr_t)LOWORD(wp);
        C3SbGetText* g = (C3SbGetText*)lp;
        if (!w || !g || idx < 0 || idx >= w->count) return 0;
        size_t need = w->texts[idx] ? wcslen(w->texts[idx]) + 1 : 1;
        g->dwLength = (DWORD)need;
        if (g->uProcessId == (UINT)(uintptr_t)GetCurrentProcessId() && g->lParam) {
            if (need > 1) memcpy((void*)g->lParam, w->texts[idx], need * sizeof(wchar_t));
            else *(wchar_t*)g->lParam = L'\0';
        }
        return (LRESULT)need;
    }
    case SB_SETPARTS: {
        int n = (int)(intptr_t)wp;
        const int* arr = (const int*)lp;
        if (!w) return 0;
        if (n <= 0) {
            w->count = 0;                       // Panels.Clear
        } else if (arr) {
            if (C3SbWinReserve(w, n) != 0) return 0;
            for (int i = 0; i < n; i++) w->rights[i] = arr[i];
            w->count = n;
        }
        // arr==NULL 且 n>0: 只是一次"重排"通知, 格数与文本都不动 (RemovePanel 这么用)。
        InvalidateRect(hw, NULL, TRUE);
        return 0;
    }
    case SB_GETPARTS: {
        int n = (int)(intptr_t)wp;
        int* arr = (int*)lp;
        if (!arr) return 0;
        int m = w ? w->count : 0;
        for (int i = 0; i < n && i < m; i++) arr[i] = w->rights[i];
        return (LRESULT)m;
    }
    case SB_SETMINHEIGHT: {
        if (!w) return 0;
        w->minHeight = (int)(intptr_t)lp;
        if (w->minHeight > 0) {
            RECT rc; GetWindowRect(hw, &rc);
            SetWindowPos(hw, NULL, 0, 0, rc.right - rc.left, w->minHeight,
                         SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
        }
        return 0;
    }
    case SB_GETMINHEIGHT:
        return w ? (LRESULT)w->minHeight : 0;
    case SB_SIMPLE:
        if (!w) return 0;
        w->simple = lp ? 1 : 0;
        InvalidateRect(hw, NULL, TRUE);
        return 0;
    case SB_ISSIMPLE:
        return w ? (LRESULT)w->simple : 0;
    case SB_SETICON: {
        int idx = (int)(intptr_t)LOWORD(wp);
        HICON ic = (HICON)lp;
        if (!w) return 0;
        if (idx < 0 || C3SbWinReserve(w, idx + 1) != 0) return 0;
        if (w->icons[idx]) DestroyIcon(w->icons[idx]);
        w->icons[idx] = ic;      // 所有权转给窗口, NCDESTROY 里统一 Destroy
        InvalidateRect(hw, NULL, TRUE);
        return 0;
    }
    case SB_GETICON: {
        int idx = (int)(intptr_t)LOWORD(wp);
        if (!w || idx < 0 || idx >= w->cap) return 0;
        return (LRESULT)(uintptr_t)w->icons[idx];
    }
    case SB_GETBORDERS: {
        int* b = (int*)lp;
        if (!b) return 0;
        b[0] = 2; b[1] = 2; b[2] = 2;   // 上 / 下 / 左 (右同左)
        return TRUE;
    }
    case SB_GETRECT: {
        RECT* r = (RECT*)lp;
        if (!r) return 0;
        GetClientRect(hw, r);
        return TRUE;
    }
    default:
        return DefWindowProcW(hw, msg, wp, lp);
    }
}

void vb6_StatusBar_RegisterClass(void) {
    static int done = 0;
    if (done) return;
    WNDCLASSW wc;
    memset(&wc, 0, sizeof(wc));
    // 已经注册 (别的机器 / 别的 comctl32 版本) 就别抢 —— 让真货赢。
    if (GetClassInfoW(NULL, L"msctls_status32", &wc)) { done = 1; return; }
    wc.style = CS_GLOBALCLASS | CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = vb6_StatusBarWndProc;
    wc.hCursor = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = L"msctls_status32";
    RegisterClassW(&wc);
    done = 1;
}

// ===================== 子类化: 时钟 + 时钟刷新 =====================
static LRESULT CALLBACK vb6_StatusBarSub(HWND hw, UINT msg, WPARAM wp, LPARAM lp) {
    Vb6StatusBar* sb = vb6_SbFromHwnd((void*)hw);
    if (msg == WM_TIMER && sb && sb->timerId && wp == VB6_SB_TIMER_ID) {
        SbRefreshClock(sb);   // 先自己刷, 再放过去 (默认过程不会重画时间面板)
        SbApplyParts(sb);
    }
    void* orig = vb6_GetOriginalWndProc((void*)hw);
    if (orig)
        return CallWindowProcW((WNDPROC)orig, hw, msg, wp, lp);
    return DefWindowProcW(hw, msg, wp, lp);
}

// ===================== 对外接口 =====================
// 设计期初始化: 窗口创建后由生成代码调用一次 (style/align 传 -1 表示用 VB6 默认)。
void vb6_StatusBar_Init(void* hwnd, int32_t style, int32_t align) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    Vb6StatusBar* sb = (Vb6StatusBar*)calloc(1, sizeof(Vb6StatusBar));
    if (!sb) return;
    sb->hwnd = hw;
    sb->style = (style == 1) ? VB6_SBR_SIMPLE : VB6_SBR_NORMAL;
    sb->align = (align >= 0) ? (int)align : VB6_SBR_ALIGN_BOT;
    sb->simpleText = NULL;

    LONG st = GetWindowLongPtrW(hw, GWL_STYLE);
    st &= ~(CCS_TOP | CCS_BOTTOM | CCS_LEFT | CCS_RIGHT | CCS_NORESIZE);
    switch (sb->align) {
        case VB6_SBR_ALIGN_TOP:  st |= CCS_TOP;  break;
        case VB6_SBR_ALIGN_LEFT: st |= CCS_LEFT; break;
        case VB6_SBR_ALIGN_RIGHT: st |= CCS_RIGHT; break;
        default: st |= CCS_BOTTOM; break;   // 2=下 (VB6 默认) / 0=无 / 4=右 → 都按"停靠"
    }
    SetWindowLongPtrW(hw, GWL_STYLE, st);
    SetWindowPos(hw, NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);

    vb6_SbStore(sb);
    vb6_InstallControlSubclass(hwnd, (void*)vb6_StatusBarSub);
    if (sb->style == VB6_SBR_SIMPLE) {
        const wchar_t* t = sb->simpleText ? sb->simpleText : L"";
        SendMessageW(hw, SB_SETTEXTW, (WPARAM)1, (LPARAM)t);
    }
}

void vb6_StatusBar_Destroy(void* hwnd) {
    if (!hwnd) return;
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb) return;
    if (sb->timerId) { KillTimer(sb->hwnd, sb->timerId); sb->timerId = 0; }
    SbEntsFree(sb);
    free(sb->simpleText); sb->simpleText = NULL;
    free(sb);
    SetPropW((HWND)hwnd, L"VB6_StatusBar", NULL);
}

// --- StatusBar 自身属性 ---
int32_t vb6_StatusBar_GetStyle(void* hwnd) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    return sb ? sb->style : VB6_SBR_NORMAL;
}

void vb6_StatusBar_SetStyle(void* hwnd, int32_t val) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb) return;
    sb->style = (val == 1) ? VB6_SBR_SIMPLE : VB6_SBR_NORMAL;
    if (sb->style == VB6_SBR_SIMPLE) {
        const wchar_t* t = sb->simpleText ? sb->simpleText : L"";
        SendMessageW(sb->hwnd, SB_SETTEXTW, (WPARAM)1, (LPARAM)t);
    }
}

void* vb6_StatusBar_GetSimpleText(void* hwnd) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb || !sb->simpleText) return NULL;
    return (void*)vb6_BSTR_FromStr(sb->simpleText);
}

void vb6_StatusBar_SetSimpleText(void* hwnd, const wchar_t* text) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb) return;
    free(sb->simpleText);
    sb->simpleText = (text && text[0]) ? _wcsdup(text) : NULL;
    if (sb->style == VB6_SBR_SIMPLE) {
        const wchar_t* t = sb->simpleText ? sb->simpleText : L"";
        SendMessageW(sb->hwnd, SB_SETTEXTW, (WPARAM)1, (LPARAM)t);
    }
}

int32_t vb6_StatusBar_GetAlign(void* hwnd) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    return sb ? sb->align : VB6_SBR_ALIGN_BOT;
}

void vb6_StatusBar_SetAlign(void* hwnd, int32_t val) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb) return;
    sb->align = (int)val;
    LONG st = GetWindowLongPtrW(sb->hwnd, GWL_STYLE);
    st &= ~(CCS_TOP | CCS_BOTTOM | CCS_LEFT | CCS_RIGHT | CCS_NORESIZE);
    switch (sb->align) {
        case VB6_SBR_ALIGN_TOP:  st |= CCS_TOP;  break;
        case VB6_SBR_ALIGN_LEFT: st |= CCS_LEFT; break;
        case VB6_SBR_ALIGN_RIGHT: st |= CCS_RIGHT; break;
        default: st |= CCS_BOTTOM; break;
    }
    SetWindowLongPtrW(sb->hwnd, GWL_STYLE, st);
    SetWindowPos(sb->hwnd, NULL, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
}

// --- Panels 集合 ---
int32_t vb6_StatusBar_GetPanelsCount(void* hwnd) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    return sb ? (int32_t)sb->count : 0;
}

// C29-4: WM_NOTIFY 换算 —— PanelClick/PanelDblClick 的事件源。
// 原生 status bar 的点击通知: NM_CLICK(-2) / NM_DBLCLK(-6), **1 基面板号放在
// NMHDR.idFrom**。返回 1 基下标 (0 = 与本控件无关/不是点击通知), cgen 侧用
// `> 0` 兜住, 再拿它调 vb6_StatusBar_PanelAt 造 Panel 对象回调。
int32_t vb6_StatusBar_OnNotify(void* hwnd, int32_t code, void* lParam) {
    if (!hwnd || !lParam) return 0;
    NMHDR* nm = (NMHDR*)lParam;
    if (nm->hwndFrom != (HWND)hwnd) return 0;   // 父窗会收到所有子控件的 WM_NOTIFY
    if (code == -2 /*NM_CLICK*/ || code == -6 /*NM_DBLCLK*/)
        return (int32_t)nm->idFrom;
    return 0;
}

// C29-4 判据辅助: **程序化**构造一次真实点击 (NM_CLICK/NM_DBLCLK 发给父窗的
// WM_NOTIFY)。无头环境点不了鼠标, 而直接调 handler 会绕开派发链 —— 只有从
// 真窗口消息进来才验得到"派发分支 + OnNotify 换算 + PanelAt 造对象"整条链。
// 判据专用 (写进 ai/029 边界), 不对应任何 VB6 语义。
void vb6_StatusBar_SimClick(void* hwnd, int32_t panelIdx, int32_t dblClick) {
    NMHDR nm;
    if (!hwnd || panelIdx < 1) return;
    memset(&nm, 0, sizeof(nm));
    nm.hwndFrom = (HWND)hwnd;
    nm.idFrom = (UINT_PTR)panelIdx;
    nm.code = dblClick ? -6 : -2;
    HWND parent = GetParent((HWND)hwnd);
    if (!parent) parent = (HWND)hwnd;
    SendMessageW(parent, WM_NOTIFY, 0, (LPARAM)&nm);
}

// Panels.Add([index], [key], [text]) —— 返回 VB6 语义的 1 基 Index, 插入失败返回 0。
int32_t vb6_StatusBar_AddPanel(void* hwnd, int32_t index, const wchar_t* key,
                               const wchar_t* text) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb) return 0;
    int pos = (index > 0) ? (int)index - 1 : sb->count;
    if (pos < 0) pos = 0;
    if (pos > sb->count) pos = sb->count;
    if (SbReserve(sb, sb->count + 1) != 0) return 0;
    // memmove 之后, a[pos] 与 a[pos+1] 是**同一份内容的两个槽** —— 而 a[pos+1] 里装的
    // 正是旧 a[pos] 挪过去后"继续活着"的那个面板。所以这一路**一个指针都不能 free**,
    // 只能在下面把空出来的 a[pos] 覆盖掉。早先这里写过 SbEntClear(&ents[pos]), 那一下
    // 把 a[pos+1] 的指针也一起释放了 → 悬垂; 等到 ClearPanels 再 free 就是双重释放
    // → 堆损坏 (症状: 程序在 ClearPanels 里 0xC000041D, 离案发地很远)。
    // 往下搬的 RemovePanel 是镜像那一半, 它靠尾部 memset 兜住 —— 两路缺一不可。
    if (pos < sb->count) {
        int mv = sb->count - pos;
        memmove(&sb->ents[pos + 1], &sb->ents[pos], (size_t)mv * sizeof(Vb6PanelEntry));
    }
    {
        Vb6PanelEntry* e = &sb->ents[pos];
        memset(e, 0, sizeof(Vb6PanelEntry));
        e->key = (key && key[0]) ? _wcsdup(key) : NULL;
        e->text = (text && text[0]) ? _wcsdup(text) : NULL;
        e->width = 0; e->minWidth = 0;
        e->autoSize = VB6_SBR_FIXED; e->style = VB6_SBR_TEXT;
    }
    sb->count++;

    // 类型面板的文本要现算; 普通面板文本为空就先给空串, 免得 SB_SETTEXT 收 NULL。
    if (sb->ents[pos].style != VB6_SBR_TEXT) SbFillAutoText(sb, pos);
    SbPushText(sb, pos);
    SbApplyParts(sb);
    SbUpdateClock(sb);
    return (int32_t)(pos + 1);
}

void vb6_StatusBar_RemovePanel(void* hwnd, const wchar_t* keyOrIndex, int32_t index) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb || sb->count <= 0) return;

    int idx = -1;
    if (keyOrIndex && keyOrIndex[0]) {
        idx = SbFindKey(sb, keyOrIndex);
        if (idx < 0) idx = _wtoi(keyOrIndex);   // 纯数字字符串按下标
    }
    if (idx < 0 && index > 0) idx = (int)index - 1;
    if (idx < 0 || idx >= sb->count) return;

    SbEntClear(&sb->ents[idx]);
    if (idx + 1 < sb->count) {
        memmove(&sb->ents[idx], &sb->ents[idx + 1],
                (size_t)(sb->count - idx - 1) * sizeof(Vb6PanelEntry));
        // 尾部那份已经被搬到 ents[idx] 了, 只留空壳 —— **绝不能再 SbEntClear 一遍**,
        // 否则 key/text 被双重 free、堆被毁 (ImageList 的 IlRemoveAt 同款坑)。
        memset(&sb->ents[sb->count - 1], 0, sizeof(Vb6PanelEntry));
    }
    sb->count--;

    // 面板删了以后下标整体前移, msctls_status32 那边要重排。
    SendMessageW(sb->hwnd, SB_SETPARTS, (WPARAM)sb->count, (LPARAM)NULL);
    for (int i = 0; i < sb->count; i++) SbPushText(sb, i);
    SbApplyParts(sb);
    SbUpdateClock(sb);
}

void vb6_StatusBar_ClearPanels(void* hwnd) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb) return;
    SbEntsFree(sb);
    SendMessageW(sb->hwnd, SB_SETPARTS, 0, (LPARAM)NULL);
    SbUpdateClock(sb);
}

// --- 面板取值 (传入的都是 VB6 的 1 基 Index) ---
static Vb6PanelEntry* SbAt(Vb6StatusBar* sb, int32_t index) {
    if (!sb) return NULL;
    int i = (int)index - 1;
    if (i < 0 || i >= sb->count) return NULL;
    return &sb->ents[i];
}

void* vb6_StatusBar_GetPanelText(void* hwnd, int32_t index) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e) return NULL;
    // sbrTime/sbrDate 这类面板 VB6 读回来就是它显示的那个值; sbrNum/sbrCaps 则
    // 保留用户写进去的 Text (设计期 .frm 里 `Panels(2) = "Tip"` 就是这份)。
    const wchar_t* t = SbDisplayText(e);
    if (!t || !t[0]) return NULL;
    return (void*)vb6_BSTR_FromStr(t);
}

void* vb6_StatusBar_GetPanelKey(void* hwnd, int32_t index) {
    Vb6PanelEntry* e = SbAt(vb6_SbFromHwnd(hwnd), index);
    if (!e || !e->key) return NULL;
    return (void*)vb6_BSTR_FromStr(e->key);
}

void* vb6_StatusBar_GetPanelKeyByKey(void* hwnd, const wchar_t* key) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb || !key || !key[0]) return NULL;
    int i = SbFindKey(sb, key);
    if (i < 0 || !sb->ents[i].key) return NULL;
    return (void*)vb6_BSTR_FromStr(sb->ents[i].key);
}

void* vb6_StatusBar_GetPanelTextByKey(void* hwnd, const wchar_t* key) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb || !key || !key[0]) return NULL;
    int i = SbFindKey(sb, key);
    if (i < 0) return NULL;
    const wchar_t* t = SbDisplayText(&sb->ents[i]);
    if (!t || !t[0]) return NULL;
    return (void*)vb6_BSTR_FromStr(t);
}

void vb6_StatusBar_SetPanelKey(void* hwnd, int32_t index, const wchar_t* key) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e || !key) return;
    // Key 只存不显示, 但要写回 text 让 SbLayout 重新分槽; 空串 = 取消 Key。
    free(e->key);
    e->key = (key[0]) ? _wcsdup(key) : NULL;
    SbPushText(sb, (int)index - 1);
}

void vb6_StatusBar_SetPanelText(void* hwnd, int32_t index, const wchar_t* text) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e) return;
    free(e->text);
    e->text = (text && text[0]) ? _wcsdup(text) : NULL;
    SbPushText(sb, (int)index - 1);
    if (e->autoSize == VB6_SBR_CONTENTS
        || (e->autoSize == VB6_SBR_FIXED && e->width <= 0))
        SbApplyParts(sb);
}

int32_t vb6_StatusBar_GetPanelIndexByKey(void* hwnd, const wchar_t* key) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    if (!sb || !key || !key[0]) return 0;
    int i = SbFindKey(sb, key);
    return (i >= 0) ? (int32_t)(i + 1) : 0;   // 出口仍是 1 基
}

int32_t vb6_StatusBar_GetPanelWidth(void* hwnd, int32_t index) {
    Vb6PanelEntry* e = SbAt(vb6_SbFromHwnd(hwnd), index);
    return e ? e->width : 0;
}

void vb6_StatusBar_SetPanelWidth(void* hwnd, int32_t index, int32_t val) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e) return;
    e->width = (int)val;
    if (e->autoSize == VB6_SBR_FIXED) SbApplyParts(sb);
}

int32_t vb6_StatusBar_GetPanelMinWidth(void* hwnd, int32_t index) {
    Vb6PanelEntry* e = SbAt(vb6_SbFromHwnd(hwnd), index);
    return e ? e->minWidth : 0;
}

void vb6_StatusBar_SetPanelMinWidth(void* hwnd, int32_t index, int32_t val) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e) return;
    e->minWidth = (int)val;
    SbApplyParts(sb);
}

int32_t vb6_StatusBar_GetPanelAutoSize(void* hwnd, int32_t index) {
    Vb6PanelEntry* e = SbAt(vb6_SbFromHwnd(hwnd), index);
    return e ? e->autoSize : VB6_SBR_FIXED;
}

void vb6_StatusBar_SetPanelAutoSize(void* hwnd, int32_t index, int32_t val) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e) return;
    e->autoSize = (int)val;
    if (e->autoSize == VB6_SBR_FIXED && e->width > 0) SbApplyParts(sb);
}

int32_t vb6_StatusBar_GetPanelStyle(void* hwnd, int32_t index) {
    Vb6PanelEntry* e = SbAt(vb6_SbFromHwnd(hwnd), index);
    return e ? e->style : VB6_SBR_TEXT;
}

void vb6_StatusBar_SetPanelStyle(void* hwnd, int32_t index, int32_t val) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e) return;
    e->style = (int)val;
    SbFillAutoText(sb, (int)index - 1);
    SbPushText(sb, (int)index - 1);
    SbApplyParts(sb);
    SbUpdateClock(sb);
}

void* vb6_StatusBar_GetPanelToolTip(void* hwnd, int32_t index) {
    Vb6PanelEntry* e = SbAt(vb6_SbFromHwnd(hwnd), index);
    if (!e || !e->tip) return NULL;
    return (void*)vb6_BSTR_FromStr(e->tip);
}

void vb6_StatusBar_SetPanelToolTip(void* hwnd, int32_t index, const wchar_t* text) {
    Vb6StatusBar* sb = vb6_SbFromHwnd(hwnd);
    Vb6PanelEntry* e = SbAt(sb, index);
    if (!e) return;
    free(e->tip);
    e->tip = (text && text[0]) ? _wcsdup(text) : NULL;
}

#endif  /* _WIN32 */
