// vb6forms_widget.c - vb6forms 模块拆分: 控件子类化 + 窗体属性 + Label/Button/Menu/Shape 属性 + 绘图属性
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
// P18-F: 控件子类化基础设施
// ============================================================

// 通用控件子类化安装 (属性名 = 本层自己的槽位)
// 账 #185 的口径: **每一层子类用自己的属性名**存它的原始窗口过程 (本层 = VB6_OrigProc,
// RTL 自绘的 PictureBox/Image = VB6_ImageOrigProc, Frame = VB6_GBox_OrigProc,
// 图形按钮 = VB6_GfxBtn_OrigProc, SSTab = VB6_SSTab_OrigProc)。两层同名 ⇒ 后装的那层
// 看见"只装一次"那一问就直接返回, 于是它的事件臂一条也不响 (实测 PictureBox/Image 全灭)。
void vb6_InstallControlSubclass(void* hwnd, void* subclassProc) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    /* Only install once */
    if (GetPropW(hw, L"VB6_OrigProc")) return;
    WNDPROC origProc = (WNDPROC)SetWindowLongPtrW(hw, GWLP_WNDPROC, (LONG_PTR)subclassProc);
    if (origProc) SetPropW(hw, L"VB6_OrigProc", (HANDLE)origProc);
}

// 获取原始窗口过程 (子类化Proc内调用CallWindowProc用)
void* vb6_GetOriginalWndProc(void* hwnd) {
    if (!hwnd) return NULL;
    return (void*)GetPropW((HWND)hwnd, L"VB6_OrigProc");
}

// 移除控件子类化 (WM_DESTROY时调用)
void vb6_RemoveControlSubclass(void* hwnd) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    WNDPROC origProc = (WNDPROC)GetPropW(hw, L"VB6_OrigProc");
    if (origProc) {
        SetWindowLongPtrW(hw, GWLP_WNDPROC, (LONG_PTR)origProc);
        RemovePropW(hw, L"VB6_OrigProc");
    }
    RemovePropW(hw, L"VB6_MouseTracked");
}

// 启动鼠标跟踪 (TrackMouseEvent封装, 用于MouseEnter/MouseLeave)
void vb6_StartMouseTracking(void* hwnd) {
    if (!hwnd) return;
    TRACKMOUSEEVENT tme;
    tme.cbSize = sizeof(tme);
    tme.dwFlags = TME_LEAVE | TME_HOVER;
    tme.hwndTrack = (HWND)hwnd;
    tme.dwHoverTime = HOVER_DEFAULT;
    TrackMouseEvent(&tme);
}

// ============================================================
// Fix 142: 容器控件把子控件的 WM_COMMAND 转发到窗体 WndProc
// ------------------------------------------------------------
// 子控件(按钮/复选框/文本框等)的通知消息发往其「直接父窗口」(即容器),
// 而生成的窗体 WndProc 只处理自己收到的 WM_COMMAND, 因此容器内子控件的
// Click/Change 事件永远不会派发 (cmdButtonInPic_Click 形同虚设).
// 这里用 comctl32 的 SetWindowSubclass 叠加安装一个转发子类: 收到
// WM_COMMAND 时把消息原样 SendMessage 给顶层窗口(窗体), 由窗体 WndProc
// 按控件 ID 派发. SetWindowSubclass 支持在同一 HWND 上叠加多个子类,
// 与生成代码使用的 vb6_InstallControlSubclass (SetWindowLongPtr 单实例)
// 互不冲突, 且可链式调用原始过程.
// ============================================================
typedef LRESULT (CALLBACK *VB6_SUBCLASSPROC_142)(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
typedef BOOL    (WINAPI   *VB6_SETWNDSUBCLASS_142)(HWND, VB6_SUBCLASSPROC_142, UINT_PTR, DWORD_PTR);
typedef LRESULT (WINAPI   *VB6_DEFSUBCLASSPROC_142)(HWND, UINT, WPARAM, LPARAM);

static VB6_DEFSUBCLASSPROC_142 g_defSubclassProc142 = NULL;
static VB6_SETWNDSUBCLASS_142  g_setWindowSubclass142 = NULL;

static LRESULT CALLBACK vb6_CmdForwardSubclass142(HWND hwnd, UINT msg, WPARAM wp,
                                                  LPARAM lp, UINT_PTR idSub,
                                                  DWORD_PTR refData) {
    (void)idSub; (void)refData;
    if (msg == WM_COMMAND) {
        HWND root = GetAncestor(hwnd, GA_ROOT);
        if (root && root != hwnd) {
            SendMessageW(root, WM_COMMAND, wp, lp);
        }
    }
    if (g_defSubclassProc142) return g_defSubclassProc142(hwnd, msg, wp, lp);
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_ForwardChildCommands(void* containerHwnd) {
    if (!containerHwnd) return;
    if (!g_setWindowSubclass142) {
        HMODULE h = GetModuleHandleW(L"comctl32.dll");
        if (!h) h = LoadLibraryW(L"comctl32.dll");
        if (!h) return;
        g_setWindowSubclass142 = (VB6_SETWNDSUBCLASS_142)GetProcAddress(h, (LPCSTR)410);
        if (!g_setWindowSubclass142)
            g_setWindowSubclass142 = (VB6_SETWNDSUBCLASS_142)GetProcAddress(h, "SetWindowSubclass");
        g_defSubclassProc142 = (VB6_DEFSUBCLASSPROC_142)GetProcAddress(h, (LPCSTR)413);
        if (!g_defSubclassProc142)
            g_defSubclassProc142 = (VB6_DEFSUBCLASSPROC_142)GetProcAddress(h, "DefSubclassProc");
    }
    if (g_setWindowSubclass142)
        g_setWindowSubclass142((HWND)containerHwnd, vb6_CmdForwardSubclass142, 0x56424346, 0);
}

// ============================================================
// P20-40: Form属性 (KeyPreview/WindowState/ControlBox/MaxButton/MinButton)
// ============================================================

int32_t vb6_GetKeyPreview(void* hwnd) {
    if (!hwnd) return 0;
    return GetPropW((HWND)hwnd, L"VB6_KeyPreview") ? -1 : 0;
}

void vb6_SetKeyPreview(void* hwnd, int32_t val) {
    if (!hwnd) return;
    if (val) SetPropW((HWND)hwnd, L"VB6_KeyPreview", (HANDLE)1);
    else RemovePropW((HWND)hwnd, L"VB6_KeyPreview");
}

int32_t vb6_GetWindowState(void* hwnd) {
    if (!hwnd) return 0;
    WINDOWPLACEMENT wp;
    wp.length = sizeof(wp);
    if (GetWindowPlacement((HWND)hwnd, &wp)) {
        switch (wp.showCmd) {
            case SW_SHOWMINIMIZED: return 1;
            case SW_SHOWMAXIMIZED: return 2;
            default: return 0;
        }
    }
    return 0;
}

void vb6_SetWindowState(void* hwnd, int32_t val) {
    if (!hwnd) return;
    switch (val) {
        case 0: ShowWindow((HWND)hwnd, SW_SHOWNORMAL); break;
        case 1: ShowWindow((HWND)hwnd, SW_MINIMIZE); break;
        case 2: ShowWindow((HWND)hwnd, SW_SHOWMAXIMIZED); break;
    }
}

int32_t vb6_GetScaleWidth(void* hwnd) {
    if (!hwnd) return 0;
    RECT rc;
    if (GetClientRect((HWND)hwnd, &rc)) {
        /* VB6 ScaleWidth: 客户区宽度, 单位 = 该窗口自己的 ScaleMode (账 #175)。
           Fix 184: 缇那一档仍走同一个 DPI 源 (vb6_ScalePxToUser 的 mode==1 分支
           就是 vb6_XToTwipX), 所以窗体 (缺省 1=缇) 的读数与改动前逐字节相同。 */
        return (int32_t)vb6_ScalePxToUser((double)(rc.right - rc.left),
                                          vb6_WindowScaleModeSelf(hwnd), 0);
    }
    return 0;
}

int32_t vb6_GetScaleHeight(void* hwnd) {
    if (!hwnd) return 0;
    RECT rc;
    if (GetClientRect((HWND)hwnd, &rc)) {
        /* VB6 ScaleHeight: 客户区高度, 单位同上 (账 #175)。 */
        return (int32_t)vb6_ScalePxToUser((double)(rc.bottom - rc.top),
                                          vb6_WindowScaleModeSelf(hwnd), 1);
    }
    return 0;
}

int32_t vb6_GetControlBox(void* hwnd) {
    if (!hwnd) return -1;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    return (style & WS_SYSMENU) ? -1 : 0;
}

void vb6_SetControlBox(void* hwnd, int32_t val) {
    if (!hwnd) return;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    if (val) style |= WS_SYSMENU;
    else style &= ~WS_SYSMENU;
    SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
    SetWindowPos((HWND)hwnd, NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
}

int32_t vb6_GetMaxButton(void* hwnd) {
    if (!hwnd) return -1;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    return (style & WS_MAXIMIZEBOX) ? -1 : 0;
}

void vb6_SetMaxButton(void* hwnd, int32_t val) {
    if (!hwnd) return;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    if (val) style |= WS_MAXIMIZEBOX;
    else style &= ~WS_MAXIMIZEBOX;
    SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
    SetWindowPos((HWND)hwnd, NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
}

int32_t vb6_GetMinButton(void* hwnd) {
    if (!hwnd) return -1;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    return (style & WS_MINIMIZEBOX) ? -1 : 0;
}

void vb6_SetMinButton(void* hwnd, int32_t val) {
    if (!hwnd) return;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    if (val) style |= WS_MINIMIZEBOX;
    else style &= ~WS_MINIMIZEBOX;
    SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
    SetWindowPos((HWND)hwnd, NULL, 0, 0, 0, 0, SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
}

// ============================================================
// P20-39: Label属性 (AutoSize/WordWrap/BackStyle)
// ============================================================

int32_t vb6_GetLabelAutoSize(void* hwnd) {
    if (!hwnd) return 0;
    return GetPropW((HWND)hwnd, L"VB6_AutoSize") ? -1 : 0;
}

void vb6_SetLabelAutoSize(void* hwnd, int32_t val) {
    if (!hwnd) return;
    if (val) {
        SetPropW((HWND)hwnd, L"VB6_AutoSize", (HANDLE)1);
        HDC hdc = GetDC((HWND)hwnd);
        // Fix 190: 标签文本是 Unicode (窗口层已切 W), 度量必须用 W 版 API ——
        // GetWindowTextA 读 Unicode 窗口会把非 ASCII 降级成 '?', GetTextExtentPoint32A
        // 再按单字节量宽, 中文标签的 AutoSize 宽度会算错 (截字/留白)。
        wchar_t text[1024] = {0};
        GetWindowTextW((HWND)hwnd, text, 1024);
        HFONT hFont = vb6_ControlFont((HWND)hwnd);   // 账 #204: 同一处问法 —— STATIC 自己不应这一问，裸问恒 NULL 就等于按 DC 默认字体算宽度
        HFONT hOld = (HFONT)SelectObject(hdc, hFont);
        SIZE sz;
        GetTextExtentPoint32W(hdc, text, (int)wcslen(text), &sz);
        SelectObject(hdc, hOld);
        ReleaseDC((HWND)hwnd, hdc);
        SetWindowPos((HWND)hwnd, NULL, 0, 0, sz.cx + 4, sz.cy + 2, SWP_NOMOVE | SWP_NOZORDER);
    } else {
        RemovePropW((HWND)hwnd, L"VB6_AutoSize");
    }
}

int32_t vb6_GetLabelWordWrap(void* hwnd) {
    if (!hwnd) return 0;
    return GetPropW((HWND)hwnd, L"VB6_WordWrap") ? -1 : 0;
}

void vb6_SetLabelWordWrap(void* hwnd, int32_t val) {
    if (!hwnd) return;
    if (val) SetPropW((HWND)hwnd, L"VB6_WordWrap", (HANDLE)1);
    else RemovePropW((HWND)hwnd, L"VB6_WordWrap");
}

int32_t vb6_GetLabelBackStyle(void* hwnd) {
    if (!hwnd) return 1;
    return GetPropW((HWND)hwnd, L"VB6_BackStyle0") ? 0 : 1;
}

void vb6_SetLabelBackStyle(void* hwnd, int32_t val) {
    if (!hwnd) return;
    if (val == 0) {
        SetPropW((HWND)hwnd, L"VB6_BackStyle0", (HANDLE)1);
        LONG exStyle = GetWindowLongW((HWND)hwnd, GWL_EXSTYLE);
        exStyle |= WS_EX_TRANSPARENT;
        SetWindowLongW((HWND)hwnd, GWL_EXSTYLE, exStyle);
        InvalidateRect((HWND)hwnd, NULL, TRUE);
    } else {
        RemovePropW((HWND)hwnd, L"VB6_BackStyle0");
        LONG exStyle = GetWindowLongW((HWND)hwnd, GWL_EXSTYLE);
        exStyle &= ~WS_EX_TRANSPARENT;
        SetWindowLongW((HWND)hwnd, GWL_EXSTYLE, exStyle);
        InvalidateRect((HWND)hwnd, NULL, TRUE);
    }
}
