// vb6rtl_system.c - VB6 运行时库: 系统对象家族：Dir/CurDir/Shell/Environ/Command/FormatDateTime + App/Clipboard/Screen/Printer/Forms
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 1790~1950 行
//   原第 1951~1993 行
//   原第 1994~2065 行
//   原第 2066~2133 行
//   原第 2134~2208 行
//   原第 2209~2236 行

#include "vb6rtl.h"
#include "vb6forms.h"   /* Fix 112: 宿主对象模型 (窗体/控件/集合/字体) */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <stdarg.h>
#include <wchar.h>
#include <wctype.h>
#include <time.h>
#ifdef _WIN32
#include <direct.h>
#include <io.h>
#include <oleauto.h>
#include <olectl.h>
#include <windows.h>
#endif

// P24-08: MessageBoxW (user32) + GetConsoleWindow (kernel32)
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "kernel32.lib")

// ============================================================

// ============================================================
// P14.2.2: 系统函数 (Dir/CurDir/Shell/Environ/Command)
// ============================================================

// Dir() - 静态状态，支持多次调用遍历
static HANDLE vb6_dir_handle = INVALID_HANDLE_VALUE;
static WIN32_FIND_DATAW vb6_dir_data;
static int vb6_dir_first = 0;


// P21-10: FormatDateTime — named date/time formatting
// namedFormat: 0=vbGeneralDate, 1=vbLongDate, 2=vbShortDate, 3=vbLongTime, 4=vbShortTime
BSTR vb6_FormatDateTime(double dateSerial, int32_t namedFormat) {
    SYSTEMTIME st;
    double intPart, fracPart;
    fracPart = modf(dateSerial, &intPart);
    if (dateSerial == 0.0) {
        GetLocalTime(&st);
    } else {
        if (!VariantTimeToSystemTime(dateSerial, &st)) {
            GetLocalTime(&st);
        }
    }
    wchar_t buf[256] = {0};
    int len = 0;
    switch (namedFormat) {
        case 1: // vbLongDate
            len = GetDateFormatW(LOCALE_USER_DEFAULT, DATE_LONGDATE, &st, NULL, buf, 256);
            break;
        case 2: // vbShortDate
            len = GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, buf, 256);
            break;
        case 3: // vbLongTime
            len = GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, NULL, buf, 256);
            break;
        case 4: // vbShortTime
            len = GetTimeFormatW(LOCALE_USER_DEFAULT, TIME_NOSECONDS, &st, NULL, buf, 256);
            break;
        case 0: // vbGeneralDate
        default: {
            // General date: date + time if time part nonzero
            len = GetDateFormatW(LOCALE_USER_DEFAULT, DATE_SHORTDATE, &st, NULL, buf, 256);
            if (len > 0) buf[len-1] = L' ';
            if (fracPart != 0.0) {
                GetTimeFormatW(LOCALE_USER_DEFAULT, 0, &st, NULL, buf + len, 256 - len);
            }
            len = (int)wcslen(buf);
            break;
        }
    }
    if (len <= 0) buf[0] = L'\0';
    return SysAllocString(buf);
}
BSTR vb6_Dir(BSTR pathname, int32_t attributes) {
    if (pathname && vb6_BSTR_Len(pathname) > 0) {
        // 新搜索: 关闭之前的句柄
        if (vb6_dir_handle != INVALID_HANDLE_VALUE) {
            FindClose(vb6_dir_handle);
            vb6_dir_handle = INVALID_HANDLE_VALUE;
        }
        vb6_dir_handle = FindFirstFileW(pathname, &vb6_dir_data);
        if (vb6_dir_handle == INVALID_HANDLE_VALUE) {
            return vb6_BSTR_Empty();  // 未找到
        }
        vb6_dir_first = 1;
        // 跳过 . 和 ..
        while (wcscmp(vb6_dir_data.cFileName, L".") == 0 ||
               wcscmp(vb6_dir_data.cFileName, L"..") == 0) {
            if (!FindNextFileW(vb6_dir_handle, &vb6_dir_data)) {
                FindClose(vb6_dir_handle);
                vb6_dir_handle = INVALID_HANDLE_VALUE;
                return vb6_BSTR_Empty();
            }
        }
        return vb6_BSTR_FromStr(vb6_dir_data.cFileName);
    } else {
        // 继续搜索
        if (vb6_dir_handle == INVALID_HANDLE_VALUE) {
            return vb6_BSTR_Empty();
        }
        while (FindNextFileW(vb6_dir_handle, &vb6_dir_data)) {
            if (wcscmp(vb6_dir_data.cFileName, L".") == 0 ||
                wcscmp(vb6_dir_data.cFileName, L"..") == 0) {
                continue;
            }
            return vb6_BSTR_FromStr(vb6_dir_data.cFileName);
        }
        FindClose(vb6_dir_handle);
        vb6_dir_handle = INVALID_HANDLE_VALUE;
        return vb6_BSTR_Empty();
    }
}

BSTR vb6_CurDir(BSTR drive) {
    wchar_t buf[MAX_PATH];
    if (drive && vb6_BSTR_Len(drive) > 0) {
        wchar_t driveLetter[4];
        driveLetter[0] = drive[0];
        driveLetter[1] = L':';
        driveLetter[2] = L'\0';
        if (GetDriveTypeW(driveLetter) == DRIVE_NO_ROOT_DIR) {
            return vb6_BSTR_Empty();
        }
        // 切换到指定驱动器获取当前目录
        wchar_t oldDir[MAX_PATH];
        GetCurrentDirectoryW(MAX_PATH, oldDir);
        SetCurrentDirectoryW(driveLetter);
        GetCurrentDirectoryW(MAX_PATH, buf);
        SetCurrentDirectoryW(oldDir);
    } else {
        GetCurrentDirectoryW(MAX_PATH, buf);
    }
    return vb6_BSTR_FromStr(buf);
}

int32_t vb6_Shell(BSTR pathname, int32_t windowstyle) {
    if (!pathname) return 0;
    // 使用WinExec简化实现 (返回值>31表示成功)
    // VB6 Shell返回进程ID，WinExec返回实例句柄
    UINT ret = WinExec(NULL, 0);  // avoid unused warning
    (void)ret;
    STARTUPINFOW si = {0};
    PROCESS_INFORMATION pi = {0};
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = (WORD)((windowstyle > 0) ? windowstyle : SW_SHOWNORMAL);
    if (CreateProcessW(NULL, pathname, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi)) {
        DWORD pid = pi.dwProcessId;
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        return (int32_t)pid;
    }
    return 0;
}

BSTR vb6_Environ(BSTR envstring) {
    if (!envstring) return vb6_BSTR_Empty();
    wchar_t buf[32768];  // Windows最大环境变量长度
    DWORD len = GetEnvironmentVariableW(envstring, buf, 32768);
    if (len == 0) return vb6_BSTR_Empty();
    return vb6_BSTR_FromStr(buf);
}

BSTR vb6_Command(void) {
    LPWSTR cmdLine = GetCommandLineW();
    if (!cmdLine) return vb6_BSTR_Empty();
    // 跳过可执行文件名 (可能在引号内)
    while (*cmdLine == L' ') cmdLine++;
    if (*cmdLine == L'"') {
        cmdLine++;
        while (*cmdLine && *cmdLine != L'"') cmdLine++;
        if (*cmdLine == L'"') cmdLine++;
    } else {
        while (*cmdLine && *cmdLine != L' ') cmdLine++;
    }
    while (*cmdLine == L' ') cmdLine++;
    if (*cmdLine == L'\0') return vb6_BSTR_Empty();
    return vb6_BSTR_FromStr(cmdLine);
}
// ============================================================
// P14.3.4: App全局对象属性
// ============================================================

/* App.Path: 返回EXE所在目录 (去掉文件名部分) */
BSTR vb6_App_Path(void) {
    wchar_t buf[1024];
    DWORD len = GetModuleFileNameW(NULL, buf, 1024);
    if (len == 0) return vb6_BSTR_FromStr(L".");
    for (DWORD i = len; i > 0; i--) {
        if (buf[i-1] == L'\\' || buf[i-1] == L'/') {
            buf[i-1] = L'\0';
            return vb6_BSTR_FromStr(buf);
        }
    }
    return vb6_BSTR_FromStr(L".");
}

/* App.EXEName: 返回EXE文件名(不含路径和扩展名) */
BSTR vb6_App_EXEName(void) {
    wchar_t buf[1024];
    DWORD len = GetModuleFileNameW(NULL, buf, 1024);
    if (len == 0) return vb6_BSTR_Empty();
    wchar_t* fname = buf;
    for (DWORD i = 0; i < len; i++) {
        if (buf[i] == L'\\' || buf[i] == L'/') fname = &buf[i+1];
    }
    wchar_t* dot = wcsrchr(fname, L'.');
    if (dot) *dot = L'\0';
    return vb6_BSTR_FromStr(fname);
}

/* App.hInstance: 返回模块实例句柄
 * Fix 179: 返回类型必须是**指针宽度**。此前返回 int32_t, 把 64 位镜像基址
 * (如 0x7FF7A6580000) 截成 0xA6580000, 写进 WNDCLASSEX.hInstance / CreateWindowEx
 * 的 HINSTANCE 形参时又被符号扩展成 0xFFFFFFFFA6580000 —— user32 在校验
 * hInstance (RtlImageNtHeader) 时裸读该地址 → 启动期 0xC0000005。
 * VB6 里 App.hInstance 声明为 Long 只是因为 VB6 只有 32 位。 */
intptr_t vb6_App_hInstance(void) {
    return (intptr_t)(uintptr_t)GetModuleHandleW(NULL);
}

/* Fix 158k: App.PrevInstance - 是否已有另一实例运行.
 * 经典单实例检测: 以 EXEName 命名互斥体, CreateMutexW 首次成功,
 * 已有实例则 GetLastError() == ERROR_ALREADY_EXISTS → 返回 True.
 * (VB6 的默认单实例机制兼容语义; 互斥体句柄随进程退出自动释放.)
 */
int32_t vb6_App_PrevInstance(void) {
    wchar_t buf[1024];
    DWORD len = GetModuleFileNameW(NULL, buf, 1024);
    if (len == 0) return 0;
    wchar_t* fname = buf;
    for (DWORD i = 0; i < len; i++) {
        if (buf[i] == L'\\' || buf[i] == L'/') fname = &buf[i+1];
    }
    wchar_t mutexName[1100];
    wsprintfW(mutexName, L"VB6_C3_SingleInstance_%s", fname);
    HANDLE hMutex = CreateMutexW(NULL, TRUE, mutexName);
    if (!hMutex) return 0;
    DWORD err = GetLastError();
    if (err == ERROR_ALREADY_EXISTS) return -1;  /* VB6 True */
    return 0;
}

/* App.HelpFile: 编译产物无 App COM 对象 → 返回空帮助文件名 (VB6默认同EXE名.hlp,
 * 语义上仅作错误/事件参数传递, 空串可编译且运行等价于无帮助文件) */
BSTR vb6_App_HelpFile(void) {
    return vb6_BSTR_FromStr(L"");
}

// ============================================================
// P18-C: Clipboard 对象
// ============================================================
void vb6_Clipboard_SetText(BSTR text) {
    if (!OpenClipboard(NULL)) return;
    EmptyClipboard();
    int len = SysStringLen(text);
    HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, (len + 1) * sizeof(wchar_t));
    if (hMem) {
        wchar_t* p = (wchar_t*)GlobalLock(hMem);
        if (p) {
            memcpy(p, text, len * sizeof(wchar_t));
            p[len] = 0;
            GlobalUnlock(hMem);
            SetClipboardData(CF_UNICODETEXT, hMem);
        } else {
            GlobalFree(hMem);
        }
    }
    CloseClipboard();
}

BSTR vb6_Clipboard_GetText(void) {
    if (!IsClipboardFormatAvailable(CF_UNICODETEXT)) return SysAllocString(L"");
    if (!OpenClipboard(NULL)) return SysAllocString(L"");
    HANDLE hData = GetClipboardData(CF_UNICODETEXT);
    BSTR result = SysAllocString(L"");
    if (hData) {
        wchar_t* p = (wchar_t*)GlobalLock(hData);
        if (p) {
            result = SysAllocString(p);
            GlobalUnlock(hData);
        }
    }
    CloseClipboard();
    return result;
}

void vb6_Clipboard_Clear(void) {
    if (OpenClipboard(NULL)) { EmptyClipboard(); CloseClipboard(); }
}

// Fix 056: Clipboard.SetData — copy IPicture bitmap to clipboard
// <vbeclipse> 补 format: VB6 有 `Clipboard.SetData data, [format]` 二形；
// format 走 vbCF* 常量 (2=vbCFBitmap, 3=vbCFMetafile, 14=vbCFEMetafile, 8=vbCFDIB)。
// 0/未指定按 vbCFBitmap 处理 (历史行为)。
void vb6_Clipboard_SetData(void* pPicture, int32_t format) {
    if (!pPicture) return;
    if (!OpenClipboard(NULL)) return;
    EmptyClipboard();
    // Try to get bitmap handle from IPicture
    HBITMAP hBmp = NULL;
    IPicture* pPic = (IPicture*)pPicture;
    HANDLE hPal = NULL;
    OLE_HANDLE hOleHandle = 0;
    pPic->lpVtbl->get_Handle(pPic, &hOleHandle);
    hBmp = (HBITMAP)(uintptr_t)hOleHandle;
    if (hBmp) {
        UINT cf = CF_BITMAP;
        if (format == 14) cf = CF_ENHMETAFILE;
        else if (format == 3) cf = CF_METAFILEPICT;
        else if (format == 8) cf = CF_DIB;
        SetClipboardData(cf, hBmp);
    }
    CloseClipboard();
}

int32_t vb6_Clipboard_GetFormat(int32_t format) {
    /* format: 1=vbCFText, 2=vbCFBitmap, 3=vbCFMetafile, 8=vbCFDIB, 9=vbCFPalette, &HBF00+=vbCFRtf */
    UINT cf = CF_UNICODETEXT;
    if (format == 1 || format == 2) cf = CF_UNICODETEXT;
    else if (format == 2) cf = CF_BITMAP;
    else if (format == 3) cf = CF_METAFILEPICT;
    else if (format == 8) cf = CF_DIB;
    else if (format == 9) cf = CF_PALETTE;
    else if (format == 0xBF00) cf = RegisterClipboardFormatW(L"Rich Text Format");
    return IsClipboardFormatAvailable(cf) ? -1 : 0;
}

// ============================================================
// P18-C: Screen 对象
// ============================================================
#define VB6_TWIPS_PER_INCH 1440

int32_t vb6_Screen_Width(void) {
    HDC dc = GetDC(NULL);
    int w = GetDeviceCaps(dc, HORZRES);
    int dpi = GetDeviceCaps(dc, LOGPIXELSX);
    ReleaseDC(NULL, dc);
    return dpi > 0 ? (int32_t)(w * (double)VB6_TWIPS_PER_INCH / dpi) : w * 15;
}

int32_t vb6_Screen_Height(void) {
    HDC dc = GetDC(NULL);
    int h = GetDeviceCaps(dc, VERTRES);
    int dpi = GetDeviceCaps(dc, LOGPIXELSY);
    ReleaseDC(NULL, dc);
    return dpi > 0 ? (int32_t)(h * (double)VB6_TWIPS_PER_INCH / dpi) : h * 15;
}

int32_t vb6_Screen_MouseX(void) {
    POINT pt; GetCursorPos(&pt);
    HDC dc = GetDC(NULL);
    int dpi = GetDeviceCaps(dc, LOGPIXELSX);
    ReleaseDC(NULL, dc);
    return dpi > 0 ? (int32_t)(pt.x * (double)VB6_TWIPS_PER_INCH / dpi) : pt.x * 15;
}

int32_t vb6_Screen_MouseY(void) {
    POINT pt; GetCursorPos(&pt);
    HDC dc = GetDC(NULL);
    int dpi = GetDeviceCaps(dc, LOGPIXELSY);
    ReleaseDC(NULL, dc);
    return dpi > 0 ? (int32_t)(pt.y * (double)VB6_TWIPS_PER_INCH / dpi) : pt.y * 15;
}

void* vb6_Screen_ActiveControl(void) {
    return (void*)GetFocus();
}

void* vb6_Screen_ActiveForm(void) {
    HWND hwnd = GetFocus();
    while (hwnd && !IsWindowVisible(hwnd)) hwnd = GetParent(hwnd);
    while (hwnd) {
        HWND parent = GetParent(hwnd);
        if (!parent || parent == GetDesktopWindow()) break;
        if (GetWindowLongPtrA(hwnd, GWLP_HWNDPARENT) == 0) break;
        hwnd = (HWND)GetWindowLongPtrA(hwnd, GWLP_HWNDPARENT);
        if (!hwnd) break;
    }
    return (void*)hwnd;
}

int32_t vb6_Screen_TwipsPerPixelX(void) {
    HDC dc = GetDC(NULL);
    int dpi = GetDeviceCaps(dc, LOGPIXELSX);
    ReleaseDC(NULL, dc);
    return dpi > 0 ? (int32_t)((double)VB6_TWIPS_PER_INCH / dpi + 0.5) : 15;
}

int32_t vb6_Screen_TwipsPerPixelY(void) {
    HDC dc = GetDC(NULL);
    int dpi = GetDeviceCaps(dc, LOGPIXELSY);
    ReleaseDC(NULL, dc);
    return dpi > 0 ? (int32_t)((double)VB6_TWIPS_PER_INCH / dpi + 0.5) : 15;
}

// Fix 161f: Screen.MousePointer 全局读/写。与控件级 vb6_Get/SetMousePointer
// (窗口 prop, vb6forms_style.c) 不同, Screen 的是 **VB6 全局**鼠标指针:
// CWaitCursor 这类类的用法 = Class_Initialize 存旧值 → ShowCursor(vbHourglass)
// → Class_Terminate 恢复。这里保存全局值并立即 SetCursor 生效。
static int32_t g_screenMousePointer = 0;

int32_t vb6_Screen_MousePointer(void) {
    return g_screenMousePointer;
}

void vb6_Screen_SetMousePointer(int32_t pointer) {
    g_screenMousePointer = pointer;
    // 99 = Custom ( vbCustom) 不动; 其余按 VB6 值映射系统光标。
    if (pointer == 99) return;
    LPCWSTR name = IDC_ARROW;
    switch (pointer) {
        case 2:  name = IDC_CROSS;   break;
        case 3:  name = IDC_IBEAM;   break;
        case 5:  case 6: case 7: case 8: case 9:
                 name = IDC_SIZEALL; break;
        case 10: name = IDC_UPARROW; break;
        case 11: name = IDC_WAIT;    break;
        case 0:  case 1: default:    break;   // Default/Arrow
    }
    SetCursor(LoadCursorW(NULL, name));
}

// ============================================================
// P18-C: Printer 对象
// ============================================================
static HDC g_printerDC = NULL;
static int g_printerCurrentX = 0;
static int g_printerCurrentY = 0;
static DOCINFOW g_docInfo = {0};
static int g_printerStarted = 0;

static void vb6_Printer_EnsureDC(void) {
    if (!g_printerDC) {
        g_printerDC = CreateDCW(L"WINSPOOL", NULL, NULL, NULL);
    }
    if (g_printerDC && !g_printerStarted) {
        memset(&g_docInfo, 0, sizeof(g_docInfo));
        g_docInfo.cbSize = sizeof(g_docInfo);
        g_docInfo.lpszDocName = L"VB6 Print Job";
        StartDocW(g_printerDC, &g_docInfo);
        StartPage(g_printerDC);
        g_printerStarted = 1;
        g_printerCurrentX = 0;
        g_printerCurrentY = 0;
    }
}

void vb6_Printer_Print(BSTR text) {
    vb6_Printer_EnsureDC();
    if (!g_printerDC) return;
    TextOutW(g_printerDC, g_printerCurrentX, g_printerCurrentY, text, SysStringLen(text));
    SIZE sz;
    GetTextExtentPoint32W(g_printerDC, text, SysStringLen(text), &sz);
    g_printerCurrentY += sz.cy;
}

void vb6_Printer_EndDoc(void) {
    if (g_printerDC && g_printerStarted) {
        EndPage(g_printerDC);
        EndDoc(g_printerDC);
        g_printerStarted = 0;
    }
    if (g_printerDC) { DeleteDC(g_printerDC); g_printerDC = NULL; }
}

void vb6_Printer_NewPage(void) {
    if (g_printerDC && g_printerStarted) {
        EndPage(g_printerDC);
        StartPage(g_printerDC);
        g_printerCurrentX = 0;
        g_printerCurrentY = 0;
    }
}

int32_t vb6_Printer_Width(void) {
    vb6_Printer_EnsureDC();
    if (!g_printerDC) return 0;
    // P20-44: VB6 Printer.Width返回twips(1440/inch), 不是pixels
    int px = GetDeviceCaps(g_printerDC, PHYSICALWIDTH);
    int dpi = GetDeviceCaps(g_printerDC, LOGPIXELSX);
    return dpi > 0 ? (int32_t)((int64_t)px * 1440 / dpi) : px;
}

int32_t vb6_Printer_Height(void) {
    vb6_Printer_EnsureDC();
    if (!g_printerDC) return 0;
    // P20-44: VB6 Printer.Height返回twips(1440/inch), 不是pixels
    int px = GetDeviceCaps(g_printerDC, PHYSICALHEIGHT);
    int dpi = GetDeviceCaps(g_printerDC, LOGPIXELSY);
    return dpi > 0 ? (int32_t)((int64_t)px * 1440 / dpi) : px;
}

int32_t vb6_Printer_CurrentX(void) { return g_printerCurrentX; }
int32_t vb6_Printer_CurrentY(void) { return g_printerCurrentY; }
void vb6_Printer_SetCurrentX(int32_t x) { g_printerCurrentX = x; }
void vb6_Printer_SetCurrentY(int32_t y) { g_printerCurrentY = y; }

// Task #39: Printer / Printers 内置全局对象 (声明见 vb6rtl_builtin.h)。
// cDlg.cls 的 `If Not Printer Is Nothing Then ... hDC = Printer.hDC` 与
// `For Each iPrn In Printers` 此前分别生成 me->Printer (C2039) / Printers
// (C2065) — 裸名拦截 (cgen_expr_ident_builtin.inc) 之后由此三个 RTL 入口承接。
void* vb6_Printer_Object(void) {
    vb6_Printer_EnsureDC();
    return (void*)g_printerDC;
}

void* vb6_Screen_Object(void) {
    return (void*)0;  /* Screen 对象哨兵: 未建模成员落 NULL 槽, 无害 */
}

void* vb6_Printer_hDC(void) {
    vb6_Printer_EnsureDC();
    return (void*)g_printerDC;
}

void* vb6_Printers_Collection(void) {
    // 空 RTL Collection — vb6_ForEach_Init 对 vb6_Collection_* 原生支持
    // (vb6_Collection_IsCollection 分支), 集合空 → For Each 循环零次。
    return vb6_Collection_New();
}

// ============================================================
// P18-C: Forms 集合
// ============================================================
#define VB6_MAX_FORMS 64
static HWND g_formList[VB6_MAX_FORMS] = {0};
static int g_formCount = 0;
static int g_msgLoopDepth = 0;   // Fix 188: 正在运行的消息循环层数 (含模态嵌套)

void vb6_Forms_LoopDepth(int delta) { g_msgLoopDepth += delta; }  // Fix 188

void vb6_Forms_Register(void* hwnd_) {
    if (g_formCount < VB6_MAX_FORMS) { g_formList[g_formCount++] = (HWND)hwnd_; }
    /* Fix 112: 窗体 HWND 也是宿主对象 (Me.ScaleWidth / Me.Controls / Me.hwnd) */
    vb6_HostObj_Register(hwnd_, NULL, "Form", 1, -1);
}
void vb6_Forms_Unregister(void* hwnd_) {
    HWND hwnd = (HWND)hwnd_;
    for (int i = 0; i < g_formCount; i++) {
        if (g_formList[i] == hwnd) {
            g_formList[i] = g_formList[--g_formCount];
            g_formList[g_formCount] = NULL;
            // Fix 188: VB6 语义是"最后一个窗体卸载 ⇒ 程序结束"。后端 Fix 144 只在
            // **启动窗体**的 WM_DESTROY 里发 PostQuitMessage, 而 `Startup = Sub Main`
            // 的工程没有启动窗体 ⇒ 没人投递 WM_QUIT, 关窗后 GetMessage 永久阻塞,
            // 进程带着存活线程驻留 (VBFlexGridDemo 实测: 窗口已销毁但 20s 后仍 6 线程)。
            // 必须限定"消息循环真在跑": 否则 Sub Main 里先 Load/Unload 再 Show 的工程
            // 会被这条提前投递的 WM_QUIT 判死。
            if (g_formCount == 0 && g_msgLoopDepth > 0) PostQuitMessage(0);
            return;
        }
    }
}

int32_t vb6_Forms_Count(void) { return g_formCount; }
void* vb6_Forms_Item(int32_t index) {
    if (index >= 0 && index < g_formCount) return (void*)g_formList[index];
    return NULL;
}

/* Fix 146: 当前活动窗体 (VB6 Screen.ActiveForm 语义).
 * 用途: 给 Show 出来的子窗体指定 Owner. Owner 为 NULL 时模态窗体无法禁用
 * 调用者 — 主窗体仍可点击, 于是能在模态循环中重入打开第二个模态窗体
 * (实测: 在 preset themes 模态窗打开时再点子窗体按钮 → 嵌套创建 → 崩溃). */
void* vb6_Forms_GetActive(void) {
    for (int i = g_formCount - 1; i >= 0; i--) {
        if (g_formList[i] && IsWindow(g_formList[i]) && IsWindowVisible(g_formList[i]))
            return (void*)g_formList[i];
    }
    if (g_formCount > 0) return (void*)g_formList[g_formCount - 1];
    return NULL;
}


