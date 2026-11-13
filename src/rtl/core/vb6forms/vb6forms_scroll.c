// vb6forms_scroll.c - vb6forms 模块拆分: 滚动条 + 定时器属性 + 文本选择 + 默认/取消按钮
// 由 vb6forms.c 按控件/窗体功能家族拆分而来 (纯搬移, 零行为改动)

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include <stdio.h>
#include <stdarg.h>

#include <stdlib.h>   /* malloc, free */
#include <oleauto.h>  /* SysAllocString, BSTR */
#include <olectl.h>   /* IPicture, OleLoadPicture, OLE_HANDLE */

// ============================================================
// 残差 #6: VScrollBar/HScrollBar 按 VB6 观感自绘
// ============================================================
// VB6 的 VScrollBar 不是原生 SCROLLBAR，而是 ThunderRT6VScrollBar（VB 自绘）。实测
// (.temp/sb_all.ps1：真 VB6 运行起来后对 6 枚 33x31 控件逐个 PrintWindow)：
//   240,240,240 = 960 / 255,255,255 = 63 / whiteRows=[15]（6 枚全等，且与 Value 无关）
//   ⇒ 240(COLOR_BTNFACE) 底 + 左边 1px 白 + 高/2 处一整行白线。
// C3 这边的原生 SCROLLBAR 在 33x31 时**一个像素都不画**：@120dpi 系统箭头高
// SM_CYVSCROLL=21，竖排两根箭头要 42px > 31px ⇒ 系统认为放不下按钮，只走
// WM_ERASEBKGND，用类背景刷（COLOR_SCROLLBAR=纯白）铺满 ⇒ 33x31 区 1023/1023 纯白、
// 无箭头、无 3D 边。同一枚控件拉到 33x200 能正常画出箭头/槽、拉到 100x31 也正常
// （.temp/sb_size.ps1 对照），所以根因是"尺寸不足 + 类背景刷是白"，不是 z-order/
// Enable/Value/range（range 改成 0..100/page10 也照样纯白，.temp/sb_range.ps1）。
// 处理：创建后子类化，按上面的实测观感自绘（不改尺寸、不改事件通道）。

// 残差 #9: hover 时增减区浮现上下两枚三角 (2026-11-08 真 VB6 取证)
// 实测 (.temp/vb6ref/shoot_hover.ps1 → .temp/vb6ref/shots/*_num_metrics.json, 120dpi, 33x31):
//   光标不在控件上 (park) / 移开后 (leave): 全区只有 240 底 + 左边 1px 白 + 中线 1px 白, ink=0
//   光标在控件上:  上下两枚三角**同时**出现;
//     被悬停的那半 (hot)    = 底宽 10 / 高 8, 三角色 102,102,102
//     另一半     (normal) = 底宽  8 / 高 6, 三角色 133,133,133
//   几何 (33x31): 上三角 y4..11(hot)/y5..10(normal) 下三角 y20..27(hot)/y21..26(normal),
//     各自在所在半区内水平居中、垂直居中; x12..21(hot)/x13..20(normal)。
// 实现: 子类过程接 WM_MOUSEMOVE + TrackMouseEvent(TME_LEAVE) → WM_MOUSELEAVE, 把 hot 半区
//       记在窗口属性上并重绘; 三角尺寸按上面的实测比例随 DPI 缩放 (96dpi ⇒ 8x6 / 6x5)。
#define VB6_SBAR_TRACK_PROP L"VB6_SBar_Track"
#define VB6_SBAR_HOT_PROP   L"VB6_SBar_Hot"

static void vb6_ScrollBarTri(HDC dc, int x, int y, int w, int h, int dir, COLORREF col) {
    /* dir: 0=上 1=下 2=左 3=右; (x,y,w,h) 是三角的**精确包围盒**。
       逐行 FillRect 而不是 Polygon: GDI 的 Polygon 填充会吃掉最右/最下一像素,
       实测 10x8 只能落 8x6 (真 VB6 是 10x8), 包围盒对不齐。 */
    HBRUSH br = CreateSolidBrush(col);
    HBRUSH old;
    int i, span, off;
    if (w < 1) w = 1;
    if (h < 1) h = 1;
    old = (HBRUSH)SelectObject(dc, br);
    if (dir == 0 || dir == 1) {
        for (i = 0; i < h; i++) {
            RECT r;
            int k = (dir == 0) ? i : (h - 1 - i);   /* k: 0=尖  h-1=底 */
            span = (h > 1) ? (1 + (w - 1) * k / (h - 1)) : w;
            if (span > w) span = w;
            off = (w - span) / 2;
            r.left = x + off; r.right = r.left + span;
            r.top = y + i;    r.bottom = r.top + 1;
            FillRect(dc, &r, br);
        }
    } else {
        for (i = 0; i < w; i++) {
            RECT r;
            int k = (dir == 2) ? i : (w - 1 - i);   /* k: 0=尖  w-1=底 */
            span = (w > 1) ? (1 + (h - 1) * k / (w - 1)) : h;
            if (span > h) span = h;
            off = (h - span) / 2;
            r.top = y + off;  r.bottom = r.top + span;
            r.left = x + i;   r.right = r.left + 1;
            FillRect(dc, &r, br);
        }
    }
    SelectObject(dc, old);
    DeleteObject(br);
}

static void vb6_ScrollBarPaintArrows(HDC dc, HWND hwnd) {
    int hot = (int)(INT_PTR)GetPropW(hwnd, VB6_SBAR_HOT_PROP);
    RECT rc;
    int dpi, hotW, hotH, nrmW, nrmH, vertical, size, mid;
    if (hot == 0) return;
    GetClientRect(hwnd, &rc);
    dpi = GetDeviceCaps(dc, LOGPIXELSY);
    if (dpi < 96) dpi = 96;
    /* 120dpi: 10x8 / 8x6 —— 与真 VB6 实拍同尺寸; 其它 DPI 等比 (96dpi ⇒ 8x6 / 6x5) */
    hotW = (10 * dpi + 60) / 120;
    hotH = ( 8 * dpi + 60) / 120;
    nrmW = ( 8 * dpi + 60) / 120;
    nrmH = ( 6 * dpi + 60) / 120;
    vertical = (GetWindowLongW(hwnd, GWL_STYLE) & SBS_VERT) != 0;
    if (vertical) {
        size = rc.bottom - rc.top;
        if (size < 3) return;
        mid = rc.top + size / 2;
        if (mid - rc.top > 0) {                    /* 上半: 向上三角 */
            int half = mid - rc.top;
            int hh = (hot == 1) ? hotH : nrmH;
            int ww = (hot == 1) ? hotW : nrmW;
            int top, left;
            if (hh > half) hh = half;
            if (ww > rc.right - rc.left) ww = rc.right - rc.left;
            top  = (rc.top + (mid - 1) + 2 - hh) / 2;   /* 半区内垂直居中 (实测同式) */
            left = (rc.right - rc.left - ww + 1) / 2;
            vb6_ScrollBarTri(dc, left, top, ww, hh, 0,
                             (hot == 1) ? RGB(102,102,102) : RGB(133,133,133));
        }
        if (rc.bottom - mid - 1 > 0) {             /* 下半: 向下三角 */
            int half = rc.bottom - mid - 1;
            int hh = (hot == 2) ? hotH : nrmH;
            int ww = (hot == 2) ? hotW : nrmW;
            int top, left;
            if (hh > half) hh = half;
            if (ww > rc.right - rc.left) ww = rc.right - rc.left;
            top  = ((mid + 1) + (rc.bottom - 1) + 2 - hh) / 2;
            left = (rc.right - rc.left - ww + 1) / 2;
            vb6_ScrollBarTri(dc, left, top, ww, hh, 1,
                             (hot == 2) ? RGB(102,102,102) : RGB(133,133,133));
        }
    } else {
        size = rc.right - rc.left;
        if (size < 3) return;
        mid = rc.left + size / 2;
        if (mid - rc.left > 0) {                   /* 左半: 向左三角 (无 VB6 样本, 镜像) */
            int half = mid - rc.left;
            int ww = (hot == 1) ? hotW : nrmW;
            int hh = (hot == 1) ? hotH : nrmH;
            int left, top;
            if (ww > half) ww = half;
            if (hh > rc.bottom - rc.top) hh = rc.bottom - rc.top;
            left = (rc.left + (mid - 1) + 2 - ww) / 2;
            top  = (rc.bottom - rc.top - hh + 1) / 2;
            vb6_ScrollBarTri(dc, left, top, ww, hh, 2,
                             (hot == 1) ? RGB(102,102,102) : RGB(133,133,133));
        }
        if (rc.right - mid - 1 > 0) {              /* 右半: 向右三角 */
            int half = rc.right - mid - 1;
            int ww = (hot == 2) ? hotW : nrmW;
            int hh = (hot == 2) ? hotH : nrmH;
            int left, top;
            if (ww > half) ww = half;
            if (hh > rc.bottom - rc.top) hh = rc.bottom - rc.top;
            left = ((mid + 1) + (rc.right - 1) + 2 - ww) / 2;
            top  = (rc.bottom - rc.top - hh + 1) / 2;
            vb6_ScrollBarTri(dc, left, top, ww, hh, 3,
                             (hot == 2) ? RGB(102,102,102) : RGB(133,133,133));
        }
    }
}

static void vb6_ScrollBarPaintTo(HDC dc, HWND hwnd) {
    RECT rc;
    GetClientRect(hwnd, &rc);
    HBRUSH bg = CreateSolidBrush(GetSysColor(COLOR_BTNFACE));
    FillRect(dc, &rc, bg);
    DeleteObject(bg);
    HBRUSH wh = (HBRUSH)GetStockObject(WHITE_BRUSH);
    if (GetWindowLongW(hwnd, GWL_STYLE) & SBS_VERT) {
        RECT left  = { rc.left, rc.top, rc.left + 1, rc.bottom };      /* 左边 1px 白 */
        int  mid   = rc.top + (rc.bottom - rc.top) / 2;
        RECT wline = { rc.left + 1, mid, rc.right, mid + 1 };           /* 中线 1px 白 */
        FillRect(dc, &left, wh);
        FillRect(dc, &wline, wh);
    } else {
        /* 横向没有真 VB6 样本，按竖向那枚的镜像落笔（上边 1px + 宽/2 处 1px） */
        RECT top   = { rc.left, rc.top, rc.right, rc.top + 1 };
        int  mid   = rc.left + (rc.right - rc.left) / 2;
        RECT wline = { mid, rc.top + 1, mid + 1, rc.bottom };
        FillRect(dc, &top, wh);
        FillRect(dc, &wline, wh);
    }
    vb6_ScrollBarPaintArrows(dc, hwnd);
}

// 残差 #7: 命中测试 —— 外观对了不等于交互通了。
// 原生 SCROLLBAR 在 33x31 下系统判定"放不下两根 21px 箭头"(SM_CYVSCROLL@120dpi=21),
// 于是既不画箭头、也不建 thumb; 它自己的命中测试因此**整块无热区**, 点下去什么都不发
// (实测 6 枚控件的 GetScrollPos 在点击前后一模一样, Text1/_Change/重绘三项全无变化)。
// VB6 的 VScrollBar 是自绘控件: 上下半区各算一个箭头, 点一次 Value ±SmallChange。
// 这里在子类过程里直接接管命中, 更新原生 pos 后向**父窗**发 WM_VSCROLL —— 事件通道
// 与用户拖拽走的那条完全相同 (生成的 WndProc 仍按 LOWORD(wp) 派发 _Change)。
static void vb6_ScrollBarStep(HWND hw, int dir) {
    int min = vb6_GetScrollMin(hw);
    int max = vb6_GetScrollMax(hw);
    /* 变量名别用 small: rpcndr.h 里有 `#define small char` (Windows 头的老坑) */
    int step = vb6_GetSmallChange(hw);
    if (step <= 0) step = 1;
    int oldPos = GetScrollPos(hw, SB_CTL);
    int winMax = max - min;
    if (winMax < 0) winMax = 0;
    int newPos = oldPos + dir * step;
    if (newPos < 0) newPos = 0;
    if (newPos > winMax) newPos = winMax;
    if (newPos == oldPos) return;          /* 到头了: VB6 也不发 Change */
    SetScrollPos(hw, SB_CTL, newPos, TRUE);
    HWND parent = GetParent(hw);
    if (parent && !vb6_Form_IsLoading(hw))
        SendMessageW(parent, WM_VSCROLL,
                     MAKEWPARAM(dir < 0 ? SB_LINEUP : SB_LINEDOWN,
                                (WORD)(SHORT)(newPos + min)),
                     (LPARAM)hw);
}

static LRESULT CALLBACK vb6_ScrollBarSubclassProc(HWND hwnd, UINT msg,
                                                 WPARAM wp, LPARAM lp) {
    WNDPROC orig = (WNDPROC)GetPropW(hwnd, L"VB6_SBar_OrigProc");
    switch (msg) {
    case WM_LBUTTONDOWN: {
        RECT rc;
        GetClientRect(hwnd, &rc);
        int vertical = (GetWindowLongW(hwnd, GWL_STYLE) & SBS_VERT) != 0;
        int pos  = vertical ? (int)(short)HIWORD(lp) : (int)(short)LOWORD(lp);
        int size = vertical ? (rc.bottom - rc.top) : (rc.right - rc.left);
        SetFocus(hwnd);
        vb6_ScrollBarStep(hwnd, pos < size / 2 ? -1 : +1);
        return 0;   /* 吞掉: 别让原生 proc 进"无箭头空转"的 tracking loop */
    }
    case WM_ERASEBKGND: {
        RECT rc;
        GetClientRect(hwnd, &rc);
        HBRUSH bg = CreateSolidBrush(GetSysColor(COLOR_BTNFACE));
        FillRect((HDC)wp, &rc, bg);
        DeleteObject(bg);
        return 1;
    }
    /* 残差 #9: 增减区 hover 才显三角 —— 进/出都要重绘。
       WM_MOUSEMOVE 里首次进入时登记 TrackMouseEvent(TME_LEAVE), 由 WM_MOUSELEAVE 收尾;
       离开不依赖"再收一条 mousemove 发现不在了", 那是拖到别的控件上才成立的假象。 */
    case WM_MOUSEMOVE: {
        RECT rc;
        int vertical, size, p, hotNew, hotOld;
        GetClientRect(hwnd, &rc);
        vertical = (GetWindowLongW(hwnd, GWL_STYLE) & SBS_VERT) != 0;
        p = vertical ? (int)(short)HIWORD(lp) : (int)(short)LOWORD(lp);
        size = vertical ? (rc.bottom - rc.top) : (rc.right - rc.left);
        hotNew = (p < size / 2) ? 1 : 2;
        hotOld = (int)(INT_PTR)GetPropW(hwnd, VB6_SBAR_HOT_PROP);
        if (hotNew != hotOld) {
            SetPropW(hwnd, VB6_SBAR_HOT_PROP, (HANDLE)(INT_PTR)hotNew);
            InvalidateRect(hwnd, NULL, FALSE);
        }
        if (!GetPropW(hwnd, VB6_SBAR_TRACK_PROP)) {
            TRACKMOUSEEVENT tme;
            tme.cbSize = sizeof(tme);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = hwnd;
            tme.dwHoverTime = 0;
            if (TrackMouseEvent(&tme))
                SetPropW(hwnd, VB6_SBAR_TRACK_PROP, (HANDLE)(INT_PTR)1);
        }
        break;
    }
    case WM_MOUSELEAVE:
        RemovePropW(hwnd, VB6_SBAR_TRACK_PROP);
        if (GetPropW(hwnd, VB6_SBAR_HOT_PROP)) {
            RemovePropW(hwnd, VB6_SBAR_HOT_PROP);
            InvalidateRect(hwnd, NULL, FALSE);
        }
        return 0;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC dc = BeginPaint(hwnd, &ps);
        vb6_ScrollBarPaintTo(dc, hwnd);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_PRINTCLIENT:
        vb6_ScrollBarPaintTo((HDC)wp, hwnd);
        return 0;
    case WM_PRINT:
        if (lp & PRF_CLIENT) vb6_ScrollBarPaintTo((HDC)wp, hwnd);
        return 0;
    default:
        break;
    }
    return orig ? CallWindowProcW(orig, hwnd, msg, wp, lp)
                : DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_InstallScrollBarSubclass(void* hwnd) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    if (GetPropW(hw, L"VB6_SBar_OrigProc")) return;
    WNDPROC orig = (WNDPROC)SetWindowLongPtrW(hw, GWLP_WNDPROC,
                                              (LONG_PTR)vb6_ScrollBarSubclassProc);
    if (orig) SetPropW(hw, L"VB6_SBar_OrigProc", (HANDLE)orig);
    InvalidateRect(hw, NULL, TRUE);
}

// ============================================================
// P13.11: ScrollBar properties (HScrollBar/VScrollBar)
// ============================================================

int vb6_GetScrollMin(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ScrollMin");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;
}

void vb6_SetScrollMin(void* hwnd, int min) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ScrollMin", (HANDLE)(INT_PTR)min);
    // Apply to Win32 scrollbar: Win32 min is always 0, we scale
    // Store VB6 Min/Max, set Win32 range to 0..(Max-Min)
    int max = vb6_GetScrollMax(hwnd);
    int winMax = max - min;
    if (winMax < 0) winMax = 0;
    HWND hw = (HWND)hwnd;
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);
    if (wcsicmp(className, L"SCROLLBAR") == 0) {
        SetScrollRange(hw, SB_CTL, 0, winMax, TRUE);
    } else {
        // Form scrollbars
        SetScrollRange(hw, SB_HORZ, 0, winMax, TRUE);
    }
    // Adjust current value
    int val = vb6_GetScrollValue(hwnd);
    int winVal = val - min;
    if (wcsicmp(className, L"SCROLLBAR") == 0) {
        SetScrollPos(hw, SB_CTL, winVal, TRUE);
    }
}

int vb6_GetScrollMax(void* hwnd) {
    if (!hwnd) return 32767;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ScrollMax");
    if (hProp) return (int)(INT_PTR)hProp;
    return 32767;  // VB6 default
}

void vb6_SetScrollMax(void* hwnd, int max) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ScrollMax", (HANDLE)(INT_PTR)max);
    int min = vb6_GetScrollMin(hwnd);
    int winMax = max - min;
    if (winMax < 0) winMax = 0;
    HWND hw = (HWND)hwnd;
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);
    if (wcsicmp(className, L"SCROLLBAR") == 0) {
        SetScrollRange(hw, SB_CTL, 0, winMax, TRUE);
    }
}

int vb6_GetScrollValue(void* hwnd) {
    if (!hwnd) return 0;
    HWND hw = (HWND)hwnd;
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);
    int winPos;
    if (wcsicmp(className, L"SCROLLBAR") == 0) {
        winPos = GetScrollPos(hw, SB_CTL);
    } else {
        winPos = GetScrollPos(hw, SB_HORZ);
    }
    int min = vb6_GetScrollMin(hwnd);
    return winPos + min;  // Convert from Win32 (0-based) to VB6
}

void vb6_SetScrollValue(void* hwnd, int value) {
    if (!hwnd) return;
    int min = vb6_GetScrollMin(hwnd);
    int winPos = value - min;
    if (winPos < 0) winPos = 0;
    HWND hw = (HWND)hwnd;
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);
    if (wcsicmp(className, L"SCROLLBAR") == 0) {
        int oldWin = GetScrollPos(hw, SB_CTL);
        SetScrollPos(hw, SB_CTL, winPos, TRUE);
        int newWin = GetScrollPos(hw, SB_CTL);
        /* 实测（.temp/evprobe，真 VB6）：Value **真变化**只发 Change（不发 Scroll），
           同值赋值静默；设计期灌值不发（vb6_Form_IsLoading 闸，见 vb6forms_prop.h）。
           发 SB_ENDSCROLL(8)：窗体 WndProc 把 4/5 之外的码一律当 Change 派发，
           与用户拖拽/点轨道那条原生通道同一分法。 */
        if (newWin != oldWin && !vb6_Form_IsLoading(hw)) {
            HWND parent = GetParent(hw);
            if (parent)
                SendMessageW(parent,
                             (GetWindowLongW(hw, GWL_STYLE) & SBS_VERT) ? WM_VSCROLL : WM_HSCROLL,
                             MAKEWPARAM(SB_ENDSCROLL, (WORD)(SHORT)(newWin + min)),
                             (LPARAM)hw);
        }
    }
}

int vb6_GetLargeChange(void* hwnd) {
    if (!hwnd) return 1;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_LargeChange");
    if (hProp) return (int)(INT_PTR)hProp;
    return 1;
}

void vb6_SetLargeChange(void* hwnd, int change) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_LargeChange", (HANDLE)(INT_PTR)change);
    // Win32 scroll info page size
    HWND hw = (HWND)hwnd;
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);
    SCROLLINFO si;
    si.cbSize = sizeof(si);
    si.fMask = SIF_PAGE;
    si.nPage = (UINT)change;
    if (wcsicmp(className, L"SCROLLBAR") == 0) {
        SetScrollInfo(hw, SB_CTL, &si, TRUE);
    }
}

int vb6_GetSmallChange(void* hwnd) {
    if (!hwnd) return 1;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_SmallChange");
    if (hProp) return (int)(INT_PTR)hProp;
    return 1;
}

void vb6_SetSmallChange(void* hwnd, int change) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_SmallChange", (HANDLE)(INT_PTR)change);
    // SmallChange is handled in WM_HSCROLL/WM_VSCROLL handler, just store
}

// ============================================================
// P13.12: Timer properties
// ============================================================

// P13.12 + C29-T: Timer properties
//
// 改之前这两个 setter 只往句柄上 SetPropW，而 Timer 是无窗口控件 ⇒ 句柄恒 NULL ⇒
// `Timer1.Enabled = True` / `Timer1.Interval = 100` 全是静默空转（实测：运行期开启后
// tick 数还是 0；改 Interval 后速率纹丝不动；关掉之后还在烧）。现在 setter 存住读数的
// 同时真的驱动引擎（起 / 停 / 按新周期重排）。
//
// Enabled 的存储照 CommonDialog 那条 normalize：VB6 的 True 是 -1，直存 val+1 会变成 0，
// 而 SetPropW 存 (HANDLE)0 与"从没设过"不可分辨 —— 于是 `Enabled = False` 读回 True。

int vb6_GetTimerInterval(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_TimerInterval");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;
}

void vb6_SetTimerInterval(void* hwnd, int interval) {
    if (!hwnd) return;
    if (interval < 0) interval = 0;
    if (interval > 65535) interval = 65535;      // VB6 口径
    SetPropW((HWND)hwnd, L"VB6_TimerInterval", (HANDLE)(INT_PTR)interval);
    vb6_TimerSetPeriod(hwnd, interval);
}

int vb6_GetTimerEnabled(void* hwnd) {
    if (!hwnd) return -1;  // VB6 默认: True
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_TimerEnabled");
    if (hProp) return ((int)(INT_PTR)hProp - 1) ? -1 : 0;
    return -1;  // True
}

void vb6_SetTimerEnabled(void* hwnd, int enabled) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_TimerEnabled", (HANDLE)(INT_PTR)((enabled ? 1 : 0) + 1));
    vb6_TimerSetEnabled(hwnd, enabled);
}

// ============================================================
// P13.14: TextBox selection properties
// ============================================================

int vb6_GetSelStart(void* hwnd) {
    if (!hwnd) return 0;
    DWORD start = 0, end = 0;
    SendMessageW((HWND)hwnd, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
    return (int)start;
}

void vb6_SetSelStart(void* hwnd, int start) {
    if (!hwnd) return;
    if (start < 0) start = 0;
    DWORD oldStart = 0, oldEnd = 0;
    SendMessageW((HWND)hwnd, EM_GETSEL, (WPARAM)&oldStart, (LPARAM)&oldEnd);
    // Keep selection length, move start
    DWORD selLen = oldEnd - oldStart;
    SendMessageW((HWND)hwnd, EM_SETSEL, (WPARAM)start, (LPARAM)(start + selLen));
}

int vb6_GetSelLength(void* hwnd) {
    if (!hwnd) return 0;
    DWORD start = 0, end = 0;
    SendMessageW((HWND)hwnd, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
    return (int)(end - start);
}

void vb6_SetSelLength(void* hwnd, int length) {
    if (!hwnd) return;
    if (length < 0) length = 0;
    DWORD start = 0, end = 0;
    SendMessageW((HWND)hwnd, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
    SendMessageW((HWND)hwnd, EM_SETSEL, (WPARAM)start, (LPARAM)(start + length));
}

void* vb6_GetSelText(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    HWND hw = (HWND)hwnd;
    DWORD start = 0, end = 0;
    SendMessageW(hw, EM_GETSEL, (WPARAM)&start, (LPARAM)&end);
    if (start >= end) return SysAllocString(L"");
    int textLen = (int)SendMessageW(hw, WM_GETTEXTLENGTH, 0, 0);
    if (textLen <= 0) return SysAllocString(L"");
    WCHAR* buf = (WCHAR*)malloc((textLen + 1) * sizeof(WCHAR));
    if (!buf) return SysAllocString(L"");
    SendMessageW(hw, WM_GETTEXT, (WPARAM)(textLen + 1), (LPARAM)buf);
    /* Extract selection from full text */
    int selLen = (int)(end - start);
    if ((int)start + selLen > textLen) selLen = textLen - (int)start;
    if (selLen < 0) selLen = 0;
    WCHAR* selBuf = (WCHAR*)malloc((selLen + 1) * sizeof(WCHAR));
    if (!selBuf) { free(buf); return SysAllocString(L""); }
    memcpy(selBuf, buf + start, selLen * sizeof(WCHAR));
    selBuf[selLen] = L'\0';
    BSTR result = SysAllocString(selBuf);
    free(selBuf);
    free(buf);
    return result;
}

void vb6_SetSelText(void* hwnd, void* bstrText) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    // Replace current selection with new text
    if (!bstrText) {
        // Replace with empty = delete selection
        SendMessageW(hw, EM_REPLACESEL, TRUE, (LPARAM)L"");
    } else {
        SendMessageW(hw, EM_REPLACESEL, TRUE, (LPARAM)(BSTR)bstrText);
    }
}

// ============================================================
// P13.15: CommandButton Default/Cancel
// ============================================================

int vb6_GetDefaultButton(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_DefaultButton");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;
}

void vb6_SetDefaultButton(void* hwnd, int isDefault) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_DefaultButton", (HANDLE)(INT_PTR)isDefault);
    // Visual feedback: default button has BS_DEFPUSHBUTTON style
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    if (isDefault) {
        style |= BS_DEFPUSHBUTTON;
        style &= ~BS_PUSHBUTTON;
    } else {
        style &= ~BS_DEFPUSHBUTTON;
        style |= BS_PUSHBUTTON;
    }
    SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
    // Redraw
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

int vb6_GetCancelButton(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_CancelButton");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;
}

void vb6_SetCancelButton(void* hwnd, int isCancel) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_CancelButton", (HANDLE)(INT_PTR)isCancel);
    // Cancel button is a logical property only; no visual change
    // The parent form's message loop should check this on VK_ESCAPE
}
