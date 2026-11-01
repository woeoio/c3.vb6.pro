// vb6forms_webview.c - vb6forms 模块拆分: WebView + Drive/Dir/FileListBox
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
// WebView2宿主 (P7.9)
// ============================================================
// 使用WebView2 COM API (ICoreWebView2) 替代VB6的SHDocVw.WebBrowser
// 动态加载WebView2Loader.dll避免编译时依赖

#include <objbase.h>
#include <exdisp.h>

// WebView2接口前向声明 (避免需要WebView2.h头文件)
// 使用动态CoCreateInstance方式

typedef struct vb6_WebViewInfo {
    void* hwnd;                          // 宿主窗口 (static控件)
    void* controller;                    // ICoreWebView2Controller
    void* webview;                       // ICoreWebView2
    int ready;                           // 1=就绪, 0=初始化中, -1=失败
    vb6_WebViewEventCallback onDocComplete;
    char currentUrl[2048];               // 当前URL
} vb6_WebViewInfo;

// 全局WebView信息表 (最多16个WebView实例)
#define VB6_WEBVIEW_MAX 16
static vb6_WebViewInfo g_webViews[VB6_WEBVIEW_MAX];
static int g_webViewCount = 0;

// 从HWND查找WebViewInfo
static vb6_WebViewInfo* vb6_FindWebViewInfo(void* hwnd) {
    for (int i = 0; i < g_webViewCount; i++) {
        if (g_webViews[i].hwnd == hwnd) return &g_webViews[i];
    }
    return NULL;
}

// WebView2控制器创建完成回调 (简化版, 使用CoCreateInstance)
// 由于真正的WebView2需要异步回调, 这里使用简化实现:
// 创建一个static控件作为占位, 在Navigate时用ShellExecute打开浏览器
// 完整实现需要链接WebView2Loader.lib

void* vb6_CreateWebView(void* hParent, int x, int y, int width, int height, const char* controlName) {
    // 创建一个static控件作为WebView的宿主区域
    // Fix 190: 窗口层 Unicode —— 占位 STATIC 用 W 版创建, 名字按 UTF-8 解码
    wchar_t* wname = vb6_u8ToWideDup(controlName);
    HWND hwnd = CreateWindowExW(0, L"STATIC", wname ? wname : L"",
        WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN,
        x, y, width, height,
        (HWND)hParent, NULL, (HINSTANCE)vb6_GetAppInstance(), NULL);
    free(wname);
    vb6_GeomOwnerStamp(hwnd, vb6_GeomAuthoredOwner(hParent));   /* 账 #305 299: 与 vb6_CreateControl 同一口径 */

    if (hwnd && g_webViewCount < VB6_WEBVIEW_MAX) {
        vb6_WebViewInfo* info = &g_webViews[g_webViewCount++];
        memset(info, 0, sizeof(*info));
        info->hwnd = (void*)hwnd;
        info->ready = 1;  // 简化版: 直接标记为就绪
        SetWindowTextW(hwnd, L"WebView2 Placeholder");
    }
    return (void*)hwnd;
}

int vb6_WebViewNavigate(void* hwnd, const char* url) {
    vb6_WebViewInfo* info = vb6_FindWebViewInfo(hwnd);
    if (!info) return -2;
    if (info->ready != 1) return -1;

    // 保存URL
    if (url) {
        strncpy(info->currentUrl, url, sizeof(info->currentUrl) - 1);
        info->currentUrl[sizeof(info->currentUrl) - 1] = 0;
    }

    // 简化实现: 使用ShellExecute打开默认浏览器
    // 完整WebView2实现应使用ICoreWebView2::Navigate()
    if (url && url[0]) {
        ShellExecuteA(NULL, "open", url, NULL, NULL, SW_SHOWNORMAL);
    }

    // 触发DocumentComplete回调
    if (info->onDocComplete) {
        info->onDocComplete(hwnd, url);
    }

    return 0;
}

void* vb6_WebViewGetUrl(void* hwnd) {
    vb6_WebViewInfo* info = vb6_FindWebViewInfo(hwnd);
    if (!info || !info->currentUrl[0]) return NULL;
    // 返回BSTR
    int len = (int)strlen(info->currentUrl);
    BSTR bstr = SysAllocStringByteLen(info->currentUrl, len);
    return (void*)bstr;
}

int vb6_WebViewIsReady(void* hwnd) {
    vb6_WebViewInfo* info = vb6_FindWebViewInfo(hwnd);
    return info ? info->ready : -2;
}

void vb6_WebViewResize(void* hwnd, int width, int height) {
    if (!hwnd) return;
    MoveWindow((HWND)hwnd, 0, 0, width, height, TRUE);
}

int vb6_WebViewExecuteScript(void* hwnd, const char* script) {
    vb6_WebViewInfo* info = vb6_FindWebViewInfo(hwnd);
    if (!info) return -2;
    if (info->ready != 1) return -1;
    // 简化版: 不执行, 完整实现应使用ICoreWebView2::ExecuteScript()
    return 0;
}

void vb6_WebViewSetDocumentCompleteCallback(void* hwnd, vb6_WebViewEventCallback callback) {
    vb6_WebViewInfo* info = vb6_FindWebViewInfo(hwnd);
    if (info) info->onDocComplete = callback;
}

// P20-50: WebBrowser navigation method stubs
int vb6_WebViewGoBack(void* hwnd) {
    (void)hwnd;
    return -1;  // Not implemented: needs ICoreWebView2::GoBack()
}

int vb6_WebViewGoForward(void* hwnd) {
    (void)hwnd;
    return -1;  // Not implemented: needs ICoreWebView2::GoForward()
}

int vb6_WebViewRefresh(void* hwnd) {
    if (!hwnd) return -1;
    vb6_WebViewInfo* info = vb6_FindWebViewInfo(hwnd);
    if (info && info->currentUrl[0]) {
        ShellExecuteA(NULL, "open", info->currentUrl, NULL, NULL, SW_SHOWNORMAL);
        return 0;
    }
    return -1;
}


// ============================================================
// P20-37: DriveListBox / DirListBox / FileListBox
// ============================================================

// Forward declarations (Refresh called by SetPath/SetPattern)
void vb6_DirListBoxRefresh(void* hwnd);
void vb6_FileListBoxRefresh(void* hwnd);

// DriveListBox: COMBOBOX populated with drive letters
BSTR vb6_DriveListBoxDrive(void* hwnd) {
    if (!hwnd) return SysAllocString(L"C");
    int sel = (int)SendMessageW((HWND)hwnd, CB_GETCURSEL, 0, 0);
    if (sel == CB_ERR) sel = 0;
    wchar_t buf[8] = L"C:";
    SendMessageW((HWND)hwnd, CB_GETLBTEXT, sel, (LPARAM)buf);
    // VB6 口径: 下拉列表里显示 "C:\", 但 Drive 属性**不带**尾反斜杠 (读数就是 "C:")。
    // 列表项保留原样, 只在这一刀出口上截掉。
    if (buf[0] && buf[1] == ':') {
        if (buf[2] == L'\\' && buf[3] == L'\0') buf[2] = L'\0';
        BSTR bstr = SysAllocString(buf);
        return bstr;
    }
    return SysAllocString(L"C:");
}

void vb6_DriveListBoxSetDrive(void* hwnd, BSTR drive) {
    if (!hwnd || !drive || SysStringLen(drive) < 1) return;
    wchar_t ch = towupper(drive[0]);
    int count = (int)SendMessageW((HWND)hwnd, CB_GETCOUNT, 0, 0);
    for (int i = 0; i < count; i++) {
        wchar_t buf[8] = {0};
        SendMessageW((HWND)hwnd, CB_GETLBTEXT, i, (LPARAM)buf);
        if (buf[0] && towupper(buf[0]) == ch) {
            SendMessageW((HWND)hwnd, CB_SETCURSEL, i, 0);
            return;
        }
    }
}

void vb6_DriveListBoxRefresh(void* hwnd) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, CB_RESETCONTENT, 0, 0);
    wchar_t drives[256] = {0};
    GetLogicalDriveStringsW(256, drives);
    wchar_t* p = drives;
    int sel = 0;
    wchar_t curDrive[4] = {0};
    // Get current drive
    GetCurrentDirectoryW(4, curDrive);
    while (*p) {
        SendMessageW((HWND)hwnd, CB_ADDSTRING, 0, (LPARAM)p);
        if (towupper(p[0]) == towupper(curDrive[0])) {
            sel = (int)SendMessageW((HWND)hwnd, CB_GETCOUNT, 0, 0) - 1;
        }
        p += wcslen(p) + 1;
    }
    SendMessageW((HWND)hwnd, CB_SETCURSEL, sel, 0);
}

// DirListBox: LISTBOX populated with directory entries
BSTR vb6_DirListBoxPath(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_DirPath");
    if (h) return SysAllocString((wchar_t*)h);
    return SysAllocString(L"");
}

void vb6_DirListBoxSetPath(void* hwnd, BSTR path) {
    if (!hwnd) return;
    // Store path as property
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_DirPath");
    if (h) { HeapFree(GetProcessHeap(), 0, h); RemovePropW((HWND)hwnd, L"VB6_DirPath"); }
    if (path && SysStringLen(path) > 0) {
        int len = SysStringLen(path) + 1;
        wchar_t* buf = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, len * sizeof(wchar_t));
        wcscpy_s(buf, len, path);
        SetPropW((HWND)hwnd, L"VB6_DirPath", buf);
    }
    vb6_DirListBoxRefresh(hwnd);
}

void vb6_DirListBoxRefresh(void* hwnd) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, LB_RESETCONTENT, 0, 0);
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_DirPath");
    if (!h) return;
    wchar_t* dirPath = (wchar_t*)h;
    wchar_t searchPath[MAX_PATH];
    _snwprintf_s(searchPath, MAX_PATH, _TRUNCATE, L"%s\\*.*", dirPath);
    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW(searchPath, &fd);
    if (hFind == INVALID_HANDLE_VALUE) return;
    do {
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            if (fd.cFileName[0] == '.' && (fd.cFileName[1] == 0 || (fd.cFileName[1] == '.' && fd.cFileName[2] == 0))) continue;
            // C29-1b: VB6 的 DirListBox 把每一项列成 [名字] —— 方括号就是"这是目录"的记号,
            // 双击下钻也靠它认 (旧实现直接列裸名字, 于是既看不出层级也没法下钻)。
            wchar_t entry[MAX_PATH + 8];
            _snwprintf_s(entry, MAX_PATH + 8, _TRUNCATE, L"[%s]", fd.cFileName);
            SendMessageW((HWND)hwnd, LB_ADDSTRING, 0, (LPARAM)entry);
        }
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);
}

/* C29-1b: 双击一项就下钻一层。返回 1 表示 Path 真的变了 (调用方据此才发 Dir1_Change,
 * 与 VB6 "Path 变了才触发 Change" 的口径一致); 选中项不是 [名字] (卷标行、空列表、
 * 双击到非目录) 时返回 0, 什么都不动。 */
int32_t vb6_DirListBoxDescendSelected(void* hwnd) {
    if (!hwnd) return 0;
    int sel = (int)SendMessageW((HWND)hwnd, LB_GETCURSEL, 0, 0);
    if (sel == LB_ERR) return 0;
    wchar_t text[MAX_PATH + 8] = {0};
    if ((int)SendMessageW((HWND)hwnd, LB_GETTEXT, sel, (LPARAM)text) == LB_ERR) return 0;
    size_t n = wcslen(text);
    if (n < 3 || text[0] != L'[' || text[n - 1] != L']') return 0;
    text[n - 1] = 0;
    wchar_t* name = text + 1;
    BSTR cur = vb6_DirListBoxPath(hwnd);
    if (!cur || SysStringLen(cur) == 0) { SysFreeString(cur); return 0; }
    wchar_t next[MAX_PATH * 2];
    size_t curLen = wcslen(cur);
    if (curLen && (cur[curLen - 1] == L'\\' || cur[curLen - 1] == L'/'))
        _snwprintf_s(next, MAX_PATH * 2, _TRUNCATE, L"%s%s", cur, name);
    else
        _snwprintf_s(next, MAX_PATH * 2, _TRUNCATE, L"%s\\%s", cur, name);
    SysFreeString(cur);
    BSTR nb = SysAllocString(next);
    vb6_DirListBoxSetPath(hwnd, nb);
    SysFreeString(nb);
    return 1;
}

// FileListBox: LISTBOX populated with file entries
BSTR vb6_FileListBoxPath(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_FilePath");
    if (h) return SysAllocString((wchar_t*)h);
    return SysAllocString(L"");
}

void vb6_FileListBoxSetPath(void* hwnd, BSTR path) {
    if (!hwnd) return;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_FilePath");
    if (h) { HeapFree(GetProcessHeap(), 0, h); RemovePropW((HWND)hwnd, L"VB6_FilePath"); }
    if (path && SysStringLen(path) > 0) {
        int len = SysStringLen(path) + 1;
        wchar_t* buf = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, len * sizeof(wchar_t));
        wcscpy_s(buf, len, path);
        SetPropW((HWND)hwnd, L"VB6_FilePath", buf);
    }
    vb6_FileListBoxRefresh(hwnd);
}

BSTR vb6_FileListBoxPattern(void* hwnd) {
    if (!hwnd) return SysAllocString(L"*.*");
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_FilePattern");
    if (h) return SysAllocString((wchar_t*)h);
    return SysAllocString(L"*.*");
}

void vb6_FileListBoxSetPattern(void* hwnd, BSTR pattern) {
    if (!hwnd) return;
    HANDLE h = GetPropW((HWND)hwnd, L"VB6_FilePattern");
    if (h) { HeapFree(GetProcessHeap(), 0, h); RemovePropW((HWND)hwnd, L"VB6_FilePattern"); }
    if (pattern && SysStringLen(pattern) > 0) {
        int len = SysStringLen(pattern) + 1;
        wchar_t* buf = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, len * sizeof(wchar_t));
        wcscpy_s(buf, len, pattern);
        SetPropW((HWND)hwnd, L"VB6_FilePattern", buf);
    }
    vb6_FileListBoxRefresh(hwnd);
}

BSTR vb6_FileListBoxFileName(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    int sel = (int)SendMessageW((HWND)hwnd, LB_GETCURSEL, 0, 0);
    if (sel == LB_ERR) return SysAllocString(L"");
    wchar_t buf[MAX_PATH] = {0};
    SendMessageW((HWND)hwnd, LB_GETTEXT, sel, (LPARAM)buf);
    return SysAllocString(buf);
}

void vb6_FileListBoxSetFileName(void* hwnd, BSTR fileName) {
    if (!hwnd || !fileName) return;
    int count = (int)SendMessageW((HWND)hwnd, LB_GETCOUNT, 0, 0);
    for (int i = 0; i < count; i++) {
        wchar_t buf[MAX_PATH] = {0};
        SendMessageW((HWND)hwnd, LB_GETTEXT, i, (LPARAM)buf);
        if (_wcsicmp(buf, fileName) == 0) {
            SendMessageW((HWND)hwnd, LB_SETCURSEL, i, 0);
            return;
        }
    }
}

void vb6_FileListBoxRefresh(void* hwnd) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, LB_RESETCONTENT, 0, 0);
    HANDLE hPath = GetPropW((HWND)hwnd, L"VB6_FilePath");
    if (!hPath) return;
    HANDLE hPat = GetPropW((HWND)hwnd, L"VB6_FilePattern");
    wchar_t* dirPath = (wchar_t*)hPath;
    wchar_t* pattern = hPat ? (wchar_t*)hPat : L"*.*";
    wchar_t searchPath[MAX_PATH];
    _snwprintf_s(searchPath, MAX_PATH, _TRUNCATE, L"%s\\%s", dirPath, pattern);
    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW(searchPath, &fd);
    if (hFind == INVALID_HANDLE_VALUE) return;
    do {
        if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)) {
            SendMessageW((HWND)hwnd, LB_ADDSTRING, 0, (LPARAM)fd.cFileName);
        }
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);
}
