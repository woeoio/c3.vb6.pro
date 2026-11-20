// vb6forms_list.c - vb6forms 模块拆分: ListBox/ComboBox + TextBox 特有属性
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
// ============================================================
// P13.3: ListBox/ComboBox properties
// ============================================================

// Helper: check if HWND is a ListBox
static int vb6_IsListBox(void* hwnd) {
    if (!hwnd) return 0;
    wchar_t cls[64];
    GetClassNameW((HWND)hwnd, cls, 64);
    // VB6 ListBox creates "ListBox" class, our RTL uses "LISTBOX"
    return (wcsicmp(cls, L"LISTBOX") == 0 || wcsicmp(cls, L"ListBox") == 0);
}

int vb6_GetListCount(void* hwnd) {
    if (!hwnd) return 0;
    if (vb6_IsListBox(hwnd)) {
        return (int)SendMessageW((HWND)hwnd, LB_GETCOUNT, 0, 0);
    }
    return (int)SendMessageW((HWND)hwnd, CB_GETCOUNT, 0, 0);
}

int vb6_GetListIndex(void* hwnd) {
    if (!hwnd) return -1;
    if (vb6_IsListBox(hwnd)) {
        return (int)SendMessageW((HWND)hwnd, LB_GETCURSEL, 0, 0);
    }
    return (int)SendMessageW((HWND)hwnd, CB_GETCURSEL, 0, 0);
}

void vb6_SetListIndex(void* hwnd, int index) {
    if (!hwnd) return;
    const int isLb = vb6_IsListBox(hwnd);
    const int prev = (int)SendMessageW((HWND)hwnd, isLb ? LB_GETCURSEL : CB_GETCURSEL, 0, 0);
    SendMessageW((HWND)hwnd, isLb ? LB_SETCURSEL : CB_SETCURSEL, (WPARAM)index, 0);
    if (prev == index) return;
    /* VB6 语义: **程序化**赋 ListIndex 也会引发 Click（用户手册明写，项目就在靠它: prjNewAniGif
     * frmUCtest.Form_Load 只有一句 `List1.ListIndex = 6`，靠它进 LoadSampleGif 加载样图）。
     * 原来这里只发 LB_SETCURSEL —— 那是纯数据设置，Win32 不会反向通知父窗，于是 Click 永远
     * 不触发，控件停在设计期画面。这里自己把同一条通知补发给父窗：生成侧的 WM_COMMAND 分派
     * 按 (控件 id, LBN_SELCHANGE/CBN_SELCHANGE=1) 找对应处理器，与用户点选时是同一条路。
     * 只在**选择真的变了**时发，避免处理器回写同一 index 时自锁。
     *
     * 用 Post 而不是 Send: Form_Load 里赋 ListIndex 时窗体还在初始化(生成的 WndProc 此时
     * 由 vb6_formLoading_* 挡事件)，同步 Send 会把处理器直接拖进初始化中途 —— prjNewAniGif
     * 里 LoadSampleGif 于是拿着半初始化的控件跑，进程随即自行退出(exit 0)。VB6 的事件本来
     * 就是从消息循环派发的，Post 才等价: 处理器在 Form_Load 返回、事件闸门放开之后才跑。 */
    {
        HWND parent = GetParent((HWND)hwnd);
        const int id = GetDlgCtrlID((HWND)hwnd);
        if (parent && id > 0) {
            PostMessageW(parent, WM_COMMAND,
                         MAKEWPARAM((WORD)id, (WORD)(isLb ? LBN_SELCHANGE : CBN_SELCHANGE)),
                         (LPARAM)(HWND)hwnd);
        }
    }
}

void* vb6_GetListItem(void* hwnd, int index) {
    if (!hwnd) return SysAllocString(L"");
    int len;
    WCHAR* buf;
    BSTR bstr;
    if (vb6_IsListBox(hwnd)) {
        len = (int)SendMessageW((HWND)hwnd, LB_GETTEXTLEN, (WPARAM)index, 0);
        if (len == LB_ERR) return SysAllocString(L"");
        buf = (WCHAR*)malloc((len + 1) * sizeof(WCHAR));
        if (!buf) return SysAllocString(L"");
        SendMessageW((HWND)hwnd, LB_GETTEXT, (WPARAM)index, (LPARAM)buf);
        bstr = SysAllocString(buf);
        free(buf);
    } else {
        len = (int)SendMessageW((HWND)hwnd, CB_GETLBTEXTLEN, (WPARAM)index, 0);
        if (len == CB_ERR) return SysAllocString(L"");
        buf = (WCHAR*)malloc((len + 1) * sizeof(WCHAR));
        if (!buf) return SysAllocString(L"");
        SendMessageW((HWND)hwnd, CB_GETLBTEXT, (WPARAM)index, (LPARAM)buf);
        bstr = SysAllocString(buf);
        free(buf);
    }
    return (void*)bstr;
}

/* 2026-11-15 (VB6 语法): `AddItem item[, index]` 的第二枚实参是**插入位置**,
 * 不是装饰 —— VB6 里 `cboFrame.AddItem "所有帧", 0` 把这一项插到**第一位**
 * (aoGIF 工程就靠它把"所有帧"放在顶部, 随后 `ListIndex = 0` 默认选中)。
 * 原因: 此处旧实现只收 (hwnd, item) 两枚, 发码把 index 丢掉 ⇒ 一律追加到末尾:
 * 列表顺序反了(所有帧跑最后), 而 `ListIndex = 0` 选中的是第一条帧 —— 观感就是
 * "默认显示值不对"。index < 0 或未给 = 追加 (与原口径同)。 */
void vb6_AddItemAt(void* hwnd, void* bstrItem, int32_t index) {
    if (!hwnd) return;
    BSTR bs = (BSTR)bstrItem;
    if (!bs) {
        bs = SysAllocString(L"");
    }
    int isLB = vb6_IsListBox(hwnd);
    LRESULT idx;
    if (index < 0) {
        idx = isLB ? SendMessageW((HWND)hwnd, LB_ADDSTRING, 0, (LPARAM)bs)
                   : SendMessageW((HWND)hwnd, CB_ADDSTRING, 0, (LPARAM)bs);
    } else {
        /* VB6: index 必须落在 0..ListCount —— 越界是错误 381。这里的口径是
         * 交给控件的 INSERTSTRING: 超尾时 Windows 自己等价于追加 (比静默丢弃好)。 */
        idx = isLB ? SendMessageW((HWND)hwnd, LB_INSERTSTRING, (WPARAM)index, (LPARAM)bs)
                   : SendMessageW((HWND)hwnd, CB_INSERTSTRING, (WPARAM)index, (LPARAM)bs);
    }
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
        int after = (int)SendMessageW((HWND)hwnd, isLB ? LB_GETCOUNT : CB_GETCOUNT, 0, 0);
        fprintf(stderr, "[C3_LIST] AddItem hwnd=%p isLB=%d at=%d ret=%ld countAfter=%d text='%ls'\n",
                hwnd, isLB, (int)index, (long)idx, after, bs ? bs : L"(null)");
    }
    if (bstrItem == NULL) {
        SysFreeString(bs);
    }
}

void vb6_AddItem(void* hwnd, void* bstrItem) {
    vb6_AddItemAt(hwnd, bstrItem, -1);
}

void vb6_RemoveItem(void* hwnd, int index) {
    if (!hwnd) return;
    if (vb6_IsListBox(hwnd)) {
        SendMessageW((HWND)hwnd, LB_DELETESTRING, (WPARAM)index, 0);
    } else {
        SendMessageW((HWND)hwnd, CB_DELETESTRING, (WPARAM)index, 0);
    }
}

void vb6_ClearList(void* hwnd) {
    if (!hwnd) return;
    if (vb6_IsListBox(hwnd)) {
        SendMessageW((HWND)hwnd, LB_RESETCONTENT, 0, 0);
    } else {
        SendMessageW((HWND)hwnd, CB_RESETCONTENT, 0, 0);
    }
}

// P14.4.3: ListBox/ComboBox extended properties
void vb6_SetListItem(void* hwnd, int index, void* bstrItem) {
    if (!hwnd) return;
    if (index < 0) return;
    const wchar_t* wstr = bstrItem ? (const wchar_t*)bstrItem : L"";
    if (vb6_IsListBox(hwnd)) {
        SendMessageW((HWND)hwnd, LB_DELETESTRING, (WPARAM)index, 0);
        SendMessageW((HWND)hwnd, LB_INSERTSTRING, (WPARAM)index, (LPARAM)wstr);
    } else {
        SendMessageW((HWND)hwnd, CB_DELETESTRING, (WPARAM)index, 0);
        SendMessageW((HWND)hwnd, CB_INSERTSTRING, (WPARAM)index, (LPARAM)wstr);
    }
}

int vb6_GetSelected(void* hwnd, int index) {
    if (!hwnd || index < 0) return 0;
    if (vb6_IsListBox(hwnd)) {
        return (int)SendMessageW((HWND)hwnd, LB_GETSEL, (WPARAM)index, 0);
    }
    return 0;  /* ComboBox: use ListIndex instead */
}

void vb6_SetSelected(void* hwnd, int index, int selected) {
    if (!hwnd || index < 0) return;
    if (vb6_IsListBox(hwnd)) {
        SendMessageW((HWND)hwnd, LB_SETSEL, (WPARAM)(selected ? 1 : 0), (LPARAM)index);
    }
}

int32_t vb6_GetItemData(void* hwnd, int index) {
    if (!hwnd || index < 0) return 0;
    if (vb6_IsListBox(hwnd)) {
        return (int32_t)SendMessageW((HWND)hwnd, LB_GETITEMDATA, (WPARAM)index, 0);
    } else {
        return (int32_t)SendMessageW((HWND)hwnd, CB_GETITEMDATA, (WPARAM)index, 0);
    }
}

void vb6_SetItemData(void* hwnd, int index, int32_t data) {
    if (!hwnd || index < 0) return;
    if (vb6_IsListBox(hwnd)) {
        SendMessageW((HWND)hwnd, LB_SETITEMDATA, (WPARAM)index, (LPARAM)data);
    } else {
        SendMessageW((HWND)hwnd, CB_SETITEMDATA, (WPARAM)index, (LPARAM)data);
    }
}

int vb6_GetNewIndex(void* hwnd) {
    if (!hwnd) return -1;
    if (vb6_IsListBox(hwnd)) {
        return (int)SendMessageW((HWND)hwnd, LB_GETCOUNT, 0, 0) - 1;
    } else {
        return (int)SendMessageW((HWND)hwnd, CB_GETCOUNT, 0, 0) - 1;
    }
}

// ============================================================
// P13.4: TextBox-specific properties
// ============================================================

int vb6_GetMultiLine(void* hwnd) {
    if (!hwnd) return 0;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    return (style & ES_MULTILINE) ? -1 : 0;  // VB6: True=-1
}

void vb6_SetMultiLine(void* hwnd, int multiline) {
    // Note: ES_MULTILINE cannot be changed after creation in Win32
    // This is a VB6 design-time only property; we store it but it has no effect
    if (!hwnd) return;
    // Store in window property for consistency
    SetPropW((HWND)hwnd, L"VB6_MultiLine", (HANDLE)(INT_PTR)(multiline ? -1 : 0));
}

int vb6_GetScrollBars(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE hProp = GetPropW((HWND)hwnd, L"VB6_ScrollBars");
    if (hProp) return (int)(INT_PTR)hProp;
    // Infer from style
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    int result = 0;
    if (style & WS_HSCROLL) result |= 1;
    if (style & WS_VSCROLL) result |= 2;
    return result;
}

void vb6_SetScrollBars(void* hwnd, int scrollbars) {
    // ScrollBars is design-time only in VB6; store as property
    if (!hwnd) return;
    SetPropW((HWND)hwnd, L"VB6_ScrollBars", (HANDLE)(INT_PTR)scrollbars);
}

int vb6_GetMaxLength(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, EM_GETLIMITTEXT, 0, 0);
}

void vb6_SetMaxLength(void* hwnd, int maxlength) {
    if (!hwnd) return;
    if (maxlength <= 0) maxlength = 0x7FFFFFFE;  // VB6: 0 = unlimited
    SendMessageW((HWND)hwnd, EM_SETLIMITTEXT, (WPARAM)maxlength, 0);
}

void* vb6_GetPasswordChar(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    WCHAR ch = (WCHAR)SendMessageW((HWND)hwnd, EM_GETPASSWORDCHAR, 0, 0);
    if (ch == 0) return SysAllocString(L"");
    WCHAR buf[2] = { ch, 0 };
    return SysAllocString(buf);
}

void vb6_SetPasswordChar(void* hwnd, void* bstrChar) {
    if (!hwnd) return;
    WCHAR ch = 0;
    if (bstrChar) {
        BSTR bs = (BSTR)bstrChar;
        if (SysStringLen(bs) > 0) ch = bs[0];
    }
    SendMessageW((HWND)hwnd, EM_SETPASSWORDCHAR, (WPARAM)ch, 0);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

int vb6_GetLocked(void* hwnd) {
    if (!hwnd) return 0;
    LONG style = GetWindowLongW((HWND)hwnd, GWL_STYLE);
    // Locked = read-only in Win32 terms
    return (style & ES_READONLY) ? -1 : 0;  // VB6: True=-1
}

void vb6_SetLocked(void* hwnd, int locked) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, EM_SETREADONLY, (WPARAM)(locked ? TRUE : FALSE), 0);
}
