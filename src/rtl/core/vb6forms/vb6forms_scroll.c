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
