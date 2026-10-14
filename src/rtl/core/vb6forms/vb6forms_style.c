// vb6forms_style.c - vb6forms 模块拆分: 对齐 / 制表 / 提示 / 鼠标指针 / 边框
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
// P13.5: Alignment
// ============================================================

int vb6_GetAlignment(void* hwnd) {
    if (!hwnd) return 0;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    if (style & ES_CENTER) return 2;
    if (style & ES_RIGHT) return 1;
    return 0;  // Left
}

void vb6_SetAlignment(void* hwnd, int align) {
    if (!hwnd) return;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    style &= ~(ES_LEFT | ES_CENTER | ES_RIGHT);
    switch (align) {
        case 1: style |= ES_RIGHT; break;
        case 2: style |= ES_CENTER; break;
        default: style |= ES_LEFT; break;
    }
    SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// ============================================================
// P13.5b: Align (PictureBox/Frame 停靠)
// ============================================================

// 0=vbAlignNone, 1=vbAlignTop, 2=vbAlignBottom, 3=vbAlignLeft, 4=vbAlignRight
// 停靠语义: 贴到父窗口客户区对应边; 固定维度(Left/Right 用宽, Top/Bottom 用高)
// 保持不变, 伸展维度撑满客户区. (VB6 还会顶开其它控件, 此处不重排其它控件.)
int vb6_GetControlAlign(void* hwnd) {
    if (!hwnd) return 0;
    return (int)(INT_PTR)GetPropW((HWND)hwnd, L"VB6_Align");
}

void vb6_SetControlAlign(void* hwnd, int align) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_Align", (HANDLE)(INT_PTR)align);
    if (align < 1 || align > 4) return;  /* 0 = vbAlignNone: 不改变位置 */
    HWND parent = GetParent((HWND)hwnd);
    if (!parent) return;
    RECT prc;
    if (!GetClientRect(parent, &prc)) return;
    int clientW = (int)(prc.right - prc.left);
    int clientH = (int)(prc.bottom - prc.top);
    RECT crc;
    GetWindowRect((HWND)hwnd, &crc);
    int x = 0, y = 0;
    int w = (int)(crc.right - crc.left);
    int h = (int)(crc.bottom - crc.top);
    switch (align) {
        case 1: w = clientW; break;                     /* Top: 撑满宽, 贴顶 */
        case 2: y = clientH - h; w = clientW; break;    /* Bottom */
        case 3: h = clientH; break;                     /* Left: 撑满高, 贴左 */
        case 4: x = clientW - w; h = clientH; break;    /* Right */
    }
    SetWindowPos((HWND)hwnd, NULL, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE);
}

// ============================================================
// P13.6: TabIndex/TabStop
// ============================================================

int vb6_GetTabIndex(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_TabIndex");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;
}

void vb6_SetTabIndex(void* hwnd, int index) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_TabIndex", (HANDLE)(INT_PTR)index);
}

int vb6_GetTabStop(void* hwnd) {
    if (!hwnd) return -1;  // Default: True
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    return (style & WS_TABSTOP) ? -1 : 0;  // VB6: True=-1
}

void vb6_SetTabStop(void* hwnd, int tabstop) {
    if (!hwnd) return;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    if (tabstop) {
        style |= WS_TABSTOP;
    } else {
        style &= ~WS_TABSTOP;
    }
    SetWindowLongW((HWND)hwnd, GWL_STYLE, style);
}

// ============================================================
// P20-12: CausesValidation
// ============================================================

int vb6_GetCausesValidation(void* hwnd) {
    if (!hwnd) return -1;  // Default: True (VB6 convention)
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_CausesValidation");
    if (hProp) return (int)(INT_PTR)hProp;
    return -1;  // Default True
}

void vb6_SetCausesValidation(void* hwnd, int causes) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_CausesValidation", (HANDLE)(INT_PTR)causes);
}

// ============================================================
// P13.8: ToolTipText
// ============================================================

// Helper: ToolTip control management
static HWND vb6_GetToolTipCtrl(void) {
    // Use a shared tooltip control (lazy init)
    static HWND s_hwndTT = NULL;
    if (!s_hwndTT) {
        s_hwndTT = CreateWindowExW(0, TOOLTIPS_CLASSW, NULL,
            WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
            CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
            NULL, NULL, vb6_hInstance, NULL);
    }
    return s_hwndTT;
}

wchar_t* vb6_GetToolTipText(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ToolTipText");
    if (!hProp) return SysAllocString(L"");
    return SysAllocString((LPCWSTR)hProp);
}

void vb6_SetToolTipText(void* hwnd, void* bstrText) {
    if (!hwnd) return;
    // Free old tooltip text stored in property
    HANDLE hOld = GetPropW((HWND)hwnd, L"VB6_ToolTipText");
    if (hOld) {
        SysFreeString((BSTR)hOld);
        RemovePropW((HWND)hwnd, L"VB6_ToolTipText");
    }
    if (bstrText) {
        BSTR bs = (BSTR)bstrText;
        BSTR copy = SysAllocString(bs);
        SetPropW((HWND)hwnd, L"VB6_ToolTipText", (HANDLE)copy);
        
        // Register with tooltip control
        HWND hwndTT = vb6_GetToolTipCtrl();
        if (hwndTT) {
            TOOLINFOW ti;
            memset(&ti, 0, sizeof(ti));
            ti.cbSize = sizeof(ti);
            ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
            ti.hwnd = GetParent((HWND)hwnd);
            ti.uId = (UINT_PTR)hwnd;
            ti.lpszText = copy;
            // Try update first, then add
            SendMessageW(hwndTT, TTM_UPDATETIPTEXTW, 0, (LPARAM)&ti);
            if (SendMessageW(hwndTT, TTM_GETTOOLINFO, 0, (LPARAM)&ti) == 0) {
                // Tool not yet added
                SendMessageW(hwndTT, TTM_ADDTOOLW, 0, (LPARAM)&ti);
            }
        }
    }
}

// C29-SL-j 判据证人（账 #148）：**这枚控件的 ToolTipText 到底进没进 tooltip 宿主**。
// 上面那条 vb6_SetToolTipText 走的是"存一份拷贝 + 往共享宿主登记工具"两步；存的那一步
// 一直有判据读（SE/CP 那几组读的就是 GetPropW 回来的串），**登记那一步以前没人验过**。
// 这里改问宿主：TTM_GETTEXT 能把工具文本读回来 = 登记成立。
// 为什么要在产物里验而不是探针：裸编的 C 探针不嵌 Common-Controls 6.0 的 manifest ⇒ 走 v5，
// 而 v5 里 TTM_ADDTOOLW 直接返回失败（实测 .build/slprobe/slmeasure15.c），拿它定罪会冤枉产品。
int vb6_ToolTipRegistered(void* hwnd) {
    HWND host;
    TOOLINFOW ti;
    wchar_t buf[128];
    LRESULT rc;
    if (!hwnd) return 0;
    host = vb6_GetToolTipCtrl();
    if (!host) return 0;
    memset(&ti, 0, sizeof(ti));
    ti.cbSize = sizeof(ti);
    ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
    ti.hwnd = GetParent((HWND)hwnd);
    ti.uId = (UINT_PTR)hwnd;
    ti.lpszText = buf;
    buf[0] = 0;
    rc = SendMessageW(host, TTM_GETTEXTW, 0, (LPARAM)&ti);
    return (rc > 0 || buf[0]) ? -1 : 0;   // VB6: True = -1
}

// ============================================================
// P13.9: Tag
// ============================================================
wchar_t* vb6_GetControlTag(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_Tag");
    if (!hProp) return SysAllocString(L"");
    return SysAllocString((LPCWSTR)hProp);
}

void vb6_SetControlTag(void* hwnd, void* bstrTag) {
    if (!hwnd) return;
    // Free old tag
    HANDLE hOld = GetPropW((HWND)hwnd, L"VB6_Tag");
    if (hOld) {
        SysFreeString((BSTR)hOld);
        RemovePropW((HWND)hwnd, L"VB6_Tag");
    }
    BSTR copy = bstrTag ? SysAllocString((BSTR)bstrTag) : SysAllocString(L"");
    SetPropW((HWND)hwnd, L"VB6_Tag", (HANDLE)copy);
}

// ============================================================
// P13.7: MousePointer/MouseIcon
// ============================================================

// VB6 MousePointer values to Windows cursor mapping
static LPCWSTR vb6_MousePointerToCursor(int pointer) {
    switch (pointer) {
        case 0:  return NULL;              // vbDefault - use class cursor
        case 1:  return IDC_ARROW;         // vbArrow
        case 2:  return IDC_CROSS;         // vbCrosshair
        case 3:  return IDC_IBEAM;         // vbIbeam
        case 4:  return IDC_ICON;          // vbIconPointer (obsolete)
        case 5:  return IDC_SIZE;          // vbSizePointer
        case 6:  return IDC_SIZENESW;      // vbSizeNESW
        case 7:  return IDC_SIZENS;        // vbSizeNS
        case 8:  return IDC_SIZENWSE;      // vbSizeNWSE
        case 9:  return IDC_SIZEWE;        // vbSizeEW
        case 10: return IDC_UPARROW;       // vbUpArrow
        case 11: return IDC_WAIT;          // vbHourglass
        case 12: return IDC_NO;            // vbNoDrop
        case 13: return IDC_APPSTARTING;   // vbArrowHourglass
        case 14: return IDC_HELP;          // vbArrowQuestion
        case 15: return IDC_SIZEALL;       // vbSizeAll
        case 99: return NULL;              // vbCustom - use MouseIcon
        default: return NULL;
    }
}

int vb6_GetMousePointer(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_MousePointer");
    if (hProp) return (int)(INT_PTR)hProp;
    return 0;  // Default
}

void vb6_SetMousePointer(void* hwnd, int pointer) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_MousePointer", (HANDLE)(INT_PTR)pointer);
    if (pointer == 99) {
        // Custom: use MouseIcon cursor if set
        HANDLE hIcon = GetPropW((HWND)hwnd, L"VB6_MouseIcon");
        if (hIcon) {
            SetClassLongPtrW((HWND)hwnd, GCLP_HCURSOR, (LONG_PTR)hIcon);
        }
    } else if (pointer == 0) {
        // Default: restore class cursor
        SetClassLongPtrW((HWND)hwnd, GCLP_HCURSOR, (LONG_PTR)LoadCursorW(NULL, IDC_ARROW));
    } else {
        LPCWSTR cursorName = vb6_MousePointerToCursor(pointer);
        if (cursorName) {
            HCURSOR hCur = LoadCursorW(NULL, cursorName);
            if (hCur) SetClassLongPtrW((HWND)hwnd, GCLP_HCURSOR, (LONG_PTR)hCur);
        }
    }
}

void* vb6_GetMouseIcon(void* hwnd) {
    if (!hwnd) return NULL;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_MouseIcon");
    return hProp;  // HCURSOR handle
}

void vb6_SetMouseIcon(void* hwnd, void* hCursor) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_MouseIcon", (HANDLE)hCursor);
    // If MousePointer is 99 (Custom), apply immediately
    int mp = vb6_GetMousePointer(hwnd);
    if (mp == 99 && hCursor) {
        SetClassLongPtrW((HWND)hwnd, GCLP_HCURSOR, (LONG_PTR)hCursor);
    }
}

// ============================================================
// P13.10: BorderStyle
// ============================================================

int vb6_GetBorderStyle(void* hwnd) {
    if (!hwnd) return 0;
    HWND hw = (HWND)hwnd;
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);

    // C29-RT-a: RICHEDIT50W（RichTextBox）与 Edit 同一条边框路 —— 它的边框就是 WS_EX_CLIENTEDGE。
    // 不走这条就落到下面那个"存属性"的兜底分支，而 SetPropW(0) 等于把属性**删掉**
    //（GetPropW 回 NULL ⇒ 读回来永远是默认 1），BorderStyle=None 就永远设不上。
    if (wcsicmp(className, L"Edit") == 0 || wcsicmp(className, L"RICHEDIT50W") == 0) {
        // TextBox / RichTextBox: 0=None(No border), 1=Fixed Single
        LONG style = GetWindowLongW(hw, GWL_EXSTYLE);
        return (style & WS_EX_CLIENTEDGE) ? 1 : 0;
    }
    // Form/ComboBox/ListBox: store as property
    // 账 #107: 这里存的是 val+1，不是 val —— SetPropW(hw, name, (HANDLE)0) 等于把属性**删掉**
    //（RemoveProp 的语义），于是 BorderStyle = 0 (None) 对所有非 Edit / 非 Static 控件都设不上：
    // 写进去当场消失，GetPropW 回 NULL ⇒ 读回来是下面那段"按类名/样式猜默认值"的结果
    //（ListBox 实测读回 2 —— LISTBOX 的 WS_BORDER 被那把 WS_OVERLAPPEDWINDOW 尺当成了 CAPTION）。
    // 同一族在 Fix 187（BackColor 的黑色，用独立哨兵）与 CommonDialog（"整数一律存 val+1"）里
    // 各修过一次，这里取后者。
    HANDLE hProp = GetPropW(hw, L"VB6_BorderStyle");
    if (hProp) return (int)(INT_PTR)hProp - 1;
    // Form default is 2 (Sizable)
    if (wcsicmp(className, L"VB6_Form") == 0 || 
        GetWindowLongW(hw, GWL_STYLE) & WS_OVERLAPPEDWINDOW) {
        return 2;
    }
    return 1;  // Default Fixed Single for most controls
}

void vb6_SetBorderStyle(void* hwnd, int style) {
    if (!hwnd) return;
    HWND hw = (HWND)hwnd;
    WCHAR className[256] = {0};
    GetClassNameW(hw, className, 256);

    if (wcsicmp(className, L"Edit") == 0 || wcsicmp(className, L"RICHEDIT50W") == 0) {
        // TextBox / RichTextBox: only 0 or 1
        LONG exStyle = GetWindowLongW(hw, GWL_EXSTYLE);
        if (style == 0) {
            exStyle &= ~WS_EX_CLIENTEDGE;
        } else {
            exStyle |= WS_EX_CLIENTEDGE;
        }
        SetWindowLongW(hw, GWL_EXSTYLE, exStyle);
        SetWindowPos(hw, NULL, 0, 0, 0, 0,
            SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
    } else if (wcsicmp(className, L"Static") == 0) {
        // Label (STATIC): 0=None, 1=Fixed Single → 切换 WS_BORDER
        LONG gstyle = GetWindowLongW(hw, GWL_STYLE);
        if (style == 0) {
            gstyle &= ~WS_BORDER;
        } else {
            gstyle |= WS_BORDER;
        }
        SetWindowLongW(hw, GWL_STYLE, gstyle);
        SetWindowPos(hw, NULL, 0, 0, 0, 0,
            SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER);
        SetPropW(hw, L"VB6_BorderStyle", (HANDLE)(INT_PTR)(style + 1));
        InvalidateRect(hw, NULL, TRUE);
    } else {
        // Store as property for other controls
        SetPropW(hw, L"VB6_BorderStyle", (HANDLE)(INT_PTR)(style + 1));
    }
}
