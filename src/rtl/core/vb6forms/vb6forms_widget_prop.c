// vb6forms_widget_prop.c - vb6forms 模块拆分: Button/Menu/Shape/PictureBox 属性读写
// 由 vb6forms_widget.c 拆出 (2026-09-17, 纯搬移, 零行为改动)

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
// P20-38: Button Style属性 (0=Standard, 1=Graphical)
// ============================================================

int32_t vb6_GetButtonStyle(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ButtonStyle");
    if (hProp) return (int32_t)(INT_PTR)hProp;
    return 0;  // Default: Standard
}

void vb6_SetButtonStyle(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ButtonStyle", (HANDLE)(INT_PTR)val);
    if (val == 1) {
        // Graphical style: make button flat/owner-draw visual
        LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
        style |= BS_BITMAP;
        SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
        InvalidateRect((HWND)hwnd, NULL, TRUE);
    } else {
        // Standard style: remove bitmap style
        LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
        style &= ~BS_BITMAP;
        SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
        InvalidateRect((HWND)hwnd, NULL, TRUE);
    }
}

// ============================================================
// P20-36: Menu属性 (Caption/Checked/Enabled/Visible)
// ============================================================

void* vb6_GetMenuCaption(void* menuHandle, int menuId) {
    if (!menuHandle) return SysAllocString(L"");
    MENUITEMINFOW mii;
    memset(&mii, 0, sizeof(mii));
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STRING;
    mii.dwTypeData = NULL;
    if (!GetMenuItemInfoW((HMENU)menuHandle, (UINT)menuId, FALSE, &mii)) {
        return SysAllocString(L"");
    }
    if (mii.cch == 0) return SysAllocString(L"");
    WCHAR* buf = (WCHAR*)malloc((mii.cch + 1) * sizeof(WCHAR));
    if (!buf) return SysAllocString(L"");
    mii.dwTypeData = buf;
    mii.cch++;
    GetMenuItemInfoW((HMENU)menuHandle, (UINT)menuId, FALSE, &mii);
    BSTR result = SysAllocString(buf);
    free(buf);
    return result;
}

void vb6_SetMenuCaption(void* menuHandle, int menuId, void* bstrCaption) {
    if (!menuHandle || !bstrCaption) return;
    MENUITEMINFOW mii;
    memset(&mii, 0, sizeof(mii));
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STRING;
    mii.dwTypeData = (BSTR)bstrCaption;
    SetMenuItemInfoW((HMENU)menuHandle, (UINT)menuId, FALSE, &mii);
    // Redraw menu bar
    HWND hwndOwner = NULL;
    if (menuHandle) {
        MENUBARINFO mbi;
        mbi.cbSize = sizeof(mbi);
        // Try to find the owner window - not easily available from HMENU
        // Just let the next paint refresh it
    }
}

int32_t vb6_GetMenuChecked(void* menuHandle, int menuId) {
    if (!menuHandle) return 0;
    MENUITEMINFOW mii;
    memset(&mii, 0, sizeof(mii));
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STATE;
    if (!GetMenuItemInfoW((HMENU)menuHandle, (UINT)menuId, FALSE, &mii)) return 0;
    return (mii.fState & MFS_CHECKED) ? -1 : 0;
}

void vb6_SetMenuChecked(void* menuHandle, int menuId, int32_t val) {
    if (!menuHandle) return;
    CheckMenuItem((HMENU)menuHandle, (UINT)menuId, val ? MF_CHECKED : MF_UNCHECKED);
}

int32_t vb6_GetMenuEnabled(void* menuHandle, int menuId) {
    if (!menuHandle) return -1;
    MENUITEMINFOW mii;
    memset(&mii, 0, sizeof(mii));
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STATE;
    if (!GetMenuItemInfoW((HMENU)menuHandle, (UINT)menuId, FALSE, &mii)) return -1;
    return (mii.fState & MFS_DISABLED) ? 0 : -1;
}

void vb6_SetMenuEnabled(void* menuHandle, int menuId, int32_t val) {
    if (!menuHandle) return;
    EnableMenuItem((HMENU)menuHandle, (UINT)menuId, val ? MF_ENABLED : (MF_DISABLED | MF_GRAYED));
}

int32_t vb6_GetMenuVisible(void* menuHandle, int menuId) {
    if (!menuHandle) return -1;
    MENUITEMINFOW mii;
    memset(&mii, 0, sizeof(mii));
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_STATE;
    if (!GetMenuItemInfoW((HMENU)menuHandle, (UINT)menuId, FALSE, &mii)) return -1;
    return (mii.fState & MFS_DEFAULT) ? 0 : -1;  // Use MS_HIDEMENU concept
}

void vb6_SetMenuVisible(void* menuHandle, int menuId, int32_t val) {
    if (!menuHandle) return;
    // Win32 doesn't have a native "hide menu item" - we remove/insert
    // For simplicity, use MF_BYCOMMAND with DeleteMenu/InsertMenu
    // But that's too destructive. Store visibility as property.
    // Use RemoveMenu/DeleteMenu only in offline scenario
    // For now: just toggle the state flag
    if (!val) {
        DeleteMenu((HMENU)menuHandle, (UINT)menuId, MF_BYCOMMAND);
    }
    // Note: making a menu item reappear after deletion is not possible
    // without recreating it. This is a known limitation.
}

// ============================================================
// P20-34: Shape属性 (Shape/BorderWidth/BorderStyle/FillStyle)
// ============================================================

int32_t vb6_GetShapeType(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ShapeType");
    if (hProp) return (int32_t)(INT_PTR)hProp;
    return 0;  // Default: Rectangle
}

void vb6_SetShapeType(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ShapeType", (HANDLE)(INT_PTR)val);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

int32_t vb6_GetShapeBorderWidth(void* hwnd) {
    if (!hwnd) return 1;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ShapeBorderWidth");
    if (hProp) return (int32_t)(INT_PTR)hProp;
    return 1;  // Default: 1
}

void vb6_SetShapeBorderWidth(void* hwnd, int32_t val) {
    if (!hwnd) return;
    if (val < 1) val = 1;
    if (val > 8192) val = 8192;
    SetPropW((HWND)hwnd, L"VB6_ShapeBorderWidth", (HANDLE)(INT_PTR)val);
    /* C29-1a: 手册那条 VB6 规则 —— BorderWidth 不是 1 而 BorderStyle 又不是 0(透明) 或
     * 6(内实线) 时, BorderStyle 被强制回 1(实线): GDI 的 PS_DASH/PS_DOT 在笔宽大于 1 时
     * 本来就画成实线, 属性读数跟着走, 不出现"读回虚线、画出实线"。 */
    if (val != 1) {
        int32_t bsC29 = vb6_GetShapeBorderStyle(hwnd);
        if (bsC29 != 0 && bsC29 != 6) vb6_SetShapeBorderStyle(hwnd, 1);
    }
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

int32_t vb6_GetShapeBorderStyle(void* hwnd) {
    if (!hwnd) return 1;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ShapeBorderStyle");
    // C29-1a: 存的是 val+1 —— SetPropW 把 (HANDLE)0 存进去后 GetPropW 回 NULL, 与"从没设过"
    // 完全不可分辨, 于是 VB6 里合法的 0 (=Transparent) 会被读成默认值 1。
    if (hProp) return (int32_t)(INT_PTR)hProp - 1;
    return 1;  // Default: Solid
}

void vb6_SetShapeBorderStyle(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ShapeBorderStyle", (HANDLE)(INT_PTR)(val + 1));
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

int32_t vb6_GetShapeFillStyle(void* hwnd) {
    if (!hwnd) return 1;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ShapeFillStyle");
    // C29-1a: 存的是 val+1 —— SetPropW 把 (HANDLE)0 存进去后 GetPropW 回 NULL, 与"从没设过"
    // 完全不可分辨, 于是 VB6 里合法的 0 (=Transparent) 会被读成默认值 1。
    if (hProp) return (int32_t)(INT_PTR)hProp - 1;
    return 1;  // Default: Transparent
}

void vb6_SetShapeFillStyle(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ShapeFillStyle", (HANDLE)(INT_PTR)(val + 1));
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// ============================================================
// P20-42: PictureBox图形属性 (AutoRedraw/ScaleMode/CurrentX/Y)
// ============================================================

int32_t vb6_GetAutoRedraw(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_AutoRedraw");
    return hProp ? -1 : 0;
}

void vb6_SetAutoRedraw(void* hwnd, int32_t val) {
    if (!hwnd) return;
    if (val) {
        SetPropW((HWND)hwnd, L"VB6_AutoRedraw", (HANDLE)1);
        // AutoRedraw: 在内存中创建持久位图
        // 获取客户区大小
        RECT rc;
        GetClientRect((HWND)hwnd, &rc);
        HDC hdc = GetDC((HWND)hwnd);
        HDC memDC = CreateCompatibleDC(hdc);
        HBITMAP memBmp = CreateCompatibleBitmap(hdc, rc.right, rc.bottom);
        HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, memBmp);
        // 用背景色填充 (VB6 语义: 持久位图初始 = BackColor; 此前硬编码
        // COLOR_BTNFACE, 窗体设了 BackColor 时首帧闪一帧错的底色)
        HBRUSH bgBrush = CreateSolidBrush((COLORREF)vb6_GetControlBackColor(hwnd));
        if (bgBrush) { FillRect(memDC, &rc, bgBrush); DeleteObject(bgBrush); }
        // 存储
        SetPropW((HWND)hwnd, L"VB6_AutoRedrawDC", (HANDLE)memDC);
        SetPropW((HWND)hwnd, L"VB6_AutoRedrawBmp", (HANDLE)memBmp);
        ReleaseDC((HWND)hwnd, hdc);
    } else {
        RemovePropW((HWND)hwnd, L"VB6_AutoRedraw");
        // 清理内存位图
        HANDLE hMemDC = GetPropW((HWND)hwnd, L"VB6_AutoRedrawDC");
        HANDLE hMemBmp = GetPropW((HWND)hwnd, L"VB6_AutoRedrawBmp");
        if (hMemDC) { DeleteDC((HDC)hMemDC); RemovePropW((HWND)hwnd, L"VB6_AutoRedrawDC"); }
        if (hMemBmp) { DeleteObject((HBITMAP)hMemBmp); RemovePropW((HWND)hwnd, L"VB6_AutoRedrawBmp"); }
    }
}

/* Fix <c3-menu3d>: AutoSize 把控件从设计期 25x25 放大到图片本征尺寸 (127x127) 之后,
 * ARDC 仍是 WM_CREATE 时按旧客户区建的小位图 —— 镜像 127x127 图标照样被 25x25 裁掉,
 * 格子只剩 BackColor。按新客户区重建记忆位图 (与 SetAutoRedraw 同口径)。 */
void vb6_AutoRedrawRefit(void* hwnd) {
    if (!hwnd) return;
    if (!GetPropW((HWND)hwnd, L"VB6_AutoRedrawDC")) return;
    vb6_SetAutoRedraw(hwnd, 0);
    vb6_SetAutoRedraw(hwnd, 1);
}

int32_t vb6_GetScaleMode(void* hwnd) {
    if (!hwnd) return 1;  // Default: Twips
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ScaleMode");
    if (hProp) return (int32_t)(INT_PTR)hProp;
    return 1;  // Twips
}

void vb6_SetScaleMode(void* hwnd, int32_t val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ScaleMode", (HANDLE)(INT_PTR)val);
}

float vb6_GetCurrentX(void* hwnd) {
    if (!hwnd) return 0.0f;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_CurrentX");
    if (hProp) {
        float val;
        memcpy(&val, &hProp, sizeof(float));
        return val;
    }
    return 0.0f;
}

void vb6_SetCurrentX(void* hwnd, float val) {
    if (!hwnd) return;
    HANDLE hProp;
    memcpy(&hProp, &val, sizeof(float));
    SetPropW((HWND)hwnd, L"VB6_CurrentX", hProp);
}

float vb6_GetCurrentY(void* hwnd) {
    if (!hwnd) return 0.0f;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_CurrentY");
    if (hProp) {
        float val;
        memcpy(&val, &hProp, sizeof(float));
        return val;
    }
    return 0.0f;
}

void vb6_SetCurrentY(void* hwnd, float val) {
    if (!hwnd) return;
    HANDLE hProp;
    memcpy(&hProp, &val, sizeof(float));
    SetPropW((HWND)hwnd, L"VB6_CurrentY", hProp);
}
