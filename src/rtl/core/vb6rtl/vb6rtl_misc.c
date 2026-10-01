// vb6rtl_misc.c - VB6 运行时库: 杂项家族：P18 补充函数（CCur/RGB/QBColor/FileDateTime/SendKeys/AppActivate/CDec/MidSet） + IIf/InputBox
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 954~1271 行
//   原第 2512~2567 行

#include "vb6rtl.h"
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
// P18: Missing RTL functions (CCur/RGB/QBColor/FileDateTime/FileLen/SendKeys/AppActivate)
// ============================================================

// Fix 126: CCur 返回 Currency 的**值** (生成代码里 Currency 按 double 承载, 与 Date
// 一致), 保留 4 位小数精度。放大整数 (值×10000) 只出现在 VT_CY 变体的 cyVal 里。
#ifdef vb6_CCur
#undef vb6_CCur   // vb6rtl_builtin.h 的 _Generic 宏在定义处必须关闭
#endif
double vb6_CCur(double v) {
    return round(v * 10000.0) / 10000.0;
}

long vb6_RGB(int32_t r, int32_t g, int32_t b) {
    /* RGB: combine red/green/blue into OLE color */
    if (r < 0) r = 0; if (r > 255) r = 255;
    if (g < 0) g = 0; if (g > 255) g = 255;
    if (b < 0) b = 0; if (b > 255) b = 255;
    return (long)(r | (g << 8) | (b << 16));
}

long vb6_QBColor(int32_t n) {
    /* QBColor: return RGB for 16 QB colors (0-15) */
    static const long qbc[16] = {
        0x000000, 0x800000, 0x008000, 0x808000,
        0x000080, 0x800080, 0x008080, 0xC0C0C0,
        0x808080, 0xFF0000, 0x00FF00, 0xFFFF00,
        0x0000FF, 0xFF00FF, 0x00FFFF, 0xFFFFFF
    };
    if (n >= 0 && n < 16) return qbc[n];
    return 0;
}

double vb6_FileDateTime(BSTR path) {
    /* FileDateTime: return VB6 date serial for file modification time */
    if (!path) return 0.0;
    WIN32_FILE_ATTRIBUTE_DATA fad;
    if (!GetFileAttributesExW(path, GetFileExInfoStandard, &fad)) return 0.0;
    /* Convert FILETIME to VB6 date serial */
    SYSTEMTIME st;
    FileTimeToSystemTime(&fad.ftLastWriteTime, &st);
    /* Build date: days since 1899-12-30 */
    struct tm t = {0};
    t.tm_year = st.wYear - 1900;
    t.tm_mon = st.wMonth - 1;
    t.tm_mday = st.wDay;
    t.tm_hour = st.wHour;
    t.tm_min = st.wMinute;
    t.tm_sec = st.wSecond;
    t.tm_isdst = -1;
    /* VB6 serial = days from 1899-12-30 + time fraction */
    /* 1899-12-30 is day -1, 1900-01-01 is day 2 in VB6 serial system */
    /* Use difftime from 1899-12-30 00:00:00 */
    struct tm epoch = {0};
    epoch.tm_year = -1;  /* 1899 */
    epoch.tm_mon = 11;   /* December */
    epoch.tm_mday = 30;
    double secs = difftime(mktime(&t), mktime(&epoch));
    return secs / 86400.0;
}

int32_t vb6_FileLen(BSTR path) {
    /* FileLen: return file size in bytes */
    if (!path) return 0;
    WIN32_FILE_ATTRIBUTE_DATA fad;
    if (!GetFileAttributesExW(path, GetFileExInfoStandard, &fad)) return 0;
    ULARGE_INTEGER size;
    size.LowPart = fad.nFileSizeLow;
    size.HighPart = fad.nFileSizeHigh;
    return (int32_t)size.QuadPart;  /* Truncate to Long for VB6 compatibility */
}
// P21-11: GetAttr — return file attributes
int32_t vb6_GetAttr(BSTR pathname) {
    if (!pathname) return 0;
    DWORD attrs = GetFileAttributesW(pathname);
    if (attrs == INVALID_FILE_ATTRIBUTES) return 0;
    return (int32_t)attrs;
}

// P21-12: SetAttr — set file attributes
void vb6_SetAttr(BSTR pathname, int32_t attributes) {
    if (!pathname) return;
    SetFileAttributesW(pathname, (DWORD)attributes);
}


/* Helper: send a single virtual key press+release */
static void vb6_SendKeyVk(WORD vk) {
    INPUT inputs[2] = {0};
    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = vk;
    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = vk;
    inputs[1].ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(2, inputs, sizeof(INPUT));
}

/* Helper: lookup VB6 SendKeys special key name to VK code */
static WORD vb6_SendKeysLookup(const wchar_t* name, size_t len) {
    /* F1-F24 */
    if ((len == 2 || len == 3) && (name[0]|0x20) == L'f') {
        int fnum = 0;
        for (size_t i = 1; i < len; i++) {
            if (name[i] >= L'0' && name[i] <= L'9') fnum = fnum * 10 + (name[i] - L'0');
            else { fnum = 0; break; }
        }
        if (fnum >= 1 && fnum <= 24) return (WORD)(VK_F1 + fnum - 1);
    }
    if (len == 3 && _wcsnicmp(name, L"TAB", 3) == 0) return VK_TAB;
    if (len == 5 && _wcsnicmp(name, L"ENTER", 5) == 0) return VK_RETURN;
    if (len == 3 && _wcsnicmp(name, L"ESC", 3) == 0) return VK_ESCAPE;
    if (len == 4 && _wcsnicmp(name, L"BACK", 4) == 0) return VK_BACK;
    if (len == 3 && _wcsnicmp(name, L"DEL", 3) == 0) return VK_DELETE;
    if (len == 6 && _wcsnicmp(name, L"DELETE", 6) == 0) return VK_DELETE;
    if (len == 3 && _wcsnicmp(name, L"INS", 3) == 0) return VK_INSERT;
    if (len == 4 && _wcsnicmp(name, L"HOME", 4) == 0) return VK_HOME;
    if (len == 3 && _wcsnicmp(name, L"END", 3) == 0) return VK_END;
    if (len == 4 && _wcsnicmp(name, L"PGUP", 4) == 0) return VK_PRIOR;
    if (len == 4 && _wcsnicmp(name, L"PGDN", 4) == 0) return VK_NEXT;
    if (len == 2 && _wcsnicmp(name, L"UP", 2) == 0) return VK_UP;
    if (len == 4 && _wcsnicmp(name, L"DOWN", 4) == 0) return VK_DOWN;
    if (len == 4 && _wcsnicmp(name, L"LEFT", 4) == 0) return VK_LEFT;
    if (len == 5 && _wcsnicmp(name, L"RIGHT", 5) == 0) return VK_RIGHT;
    if (len == 8 && _wcsnicmp(name, L"CAPSLOCK", 8) == 0) return VK_CAPITAL;
    if (len == 7 && _wcsnicmp(name, L"NUMLOCK", 7) == 0) return VK_NUMLOCK;
    if (len == 10 && _wcsnicmp(name, L"SCROLLLOCK", 10) == 0) return VK_SCROLL;
    if (len == 5 && _wcsnicmp(name, L"BREAK", 5) == 0) return VK_CANCEL;
    if (len == 4 && _wcsnicmp(name, L"HELP", 4) == 0) return VK_HELP;
    if (len == 6 && _wcsnicmp(name, L"PRTSC", 5) == 0) return VK_SNAPSHOT;
    if (len == 11 && _wcsnicmp(name, L"PRINTSCREEN", 11) == 0) return VK_SNAPSHOT;
    if (len == 5 && _wcsnicmp(name, L"SPACE", 5) == 0) return VK_SPACE;
    if (len == 8 && _wcsnicmp(name, L"BACKSPACE", 9) == 0) return VK_BACK;
    if (len == 8 && _wcsnicmp(name, L"CLEAR", 5) == 0) return VK_CLEAR;
    return 0;
}

void vb6_SendKeys(BSTR keys, int32_t wait) {
    /* SendKeys: supports single chars, {ENTER}/{TAB}/{ESC}/{F1}-{F24} etc., +^% modifiers */
    (void)wait;
    if (!keys) return;
    size_t len = vb6_BSTR_Len(keys);
    size_t i = 0;
    while (i < len) {
        int modShift = 0, modCtrl = 0, modAlt = 0;
        while (i < len && (keys[i] == L'+' || keys[i] == L'^' || keys[i] == L'%')) {
            if (keys[i] == L'+') modShift = 1;
            else if (keys[i] == L'^') modCtrl = 1;
            else if (keys[i] == L'%') modAlt = 1;
            i++;
        }
        if (i >= len) break;
        WORD vk = 0;
        int isUnicode = 0;
        wchar_t ch = 0;
        if (keys[i] == L'{') {
            size_t start = i + 1;
            size_t end = start;
            while (end < len && keys[end] != L'}') end++;
            if (end < len) {
                size_t klen = end - start;
                if (klen == 1 && keys[start] == L'~') {
                    vk = VK_RETURN;
                } else {
                    vk = vb6_SendKeysLookup(&keys[start], klen);
                }
            }
            i = end + 1;
        } else {
            isUnicode = 1;
            ch = keys[i];
            i++;
        }
        /* Press modifier keys */
        if (modShift) { INPUT mi = {0}; mi.type = INPUT_KEYBOARD; mi.ki.wVk = VK_SHIFT; SendInput(1, &mi, sizeof(INPUT)); }
        if (modCtrl) { INPUT mi = {0}; mi.type = INPUT_KEYBOARD; mi.ki.wVk = VK_CONTROL; SendInput(1, &mi, sizeof(INPUT)); }
        if (modAlt) { INPUT mi = {0}; mi.type = INPUT_KEYBOARD; mi.ki.wVk = VK_MENU; SendInput(1, &mi, sizeof(INPUT)); }
        if (isUnicode) {
            INPUT inputs[2] = {0};
            inputs[0].type = INPUT_KEYBOARD;
            inputs[0].ki.wScan = ch;
            inputs[0].ki.dwFlags = KEYEVENTF_UNICODE;
            inputs[1].type = INPUT_KEYBOARD;
            inputs[1].ki.wScan = ch;
            inputs[1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
            SendInput(2, inputs, sizeof(INPUT));
        } else if (vk) {
            vb6_SendKeyVk(vk);
        }
        /* Release modifier keys */
        if (modAlt) { INPUT mi = {0}; mi.type = INPUT_KEYBOARD; mi.ki.wVk = VK_MENU; mi.ki.dwFlags = KEYEVENTF_KEYUP; SendInput(1, &mi, sizeof(INPUT)); }
        if (modCtrl) { INPUT mi = {0}; mi.type = INPUT_KEYBOARD; mi.ki.wVk = VK_CONTROL; mi.ki.dwFlags = KEYEVENTF_KEYUP; SendInput(1, &mi, sizeof(INPUT)); }
        if (modShift) { INPUT mi = {0}; mi.type = INPUT_KEYBOARD; mi.ki.wVk = VK_SHIFT; mi.ki.dwFlags = KEYEVENTF_KEYUP; SendInput(1, &mi, sizeof(INPUT)); }
    }
}

/* AppActivate EnumWindows callback context */
typedef struct { DWORD pid; HWND result; } vb6_AppActivateCtx;
static BOOL CALLBACK vb6_AppActivateEnumProc(HWND hwnd, LPARAM lParam) {
    vb6_AppActivateCtx* ctx = (vb6_AppActivateCtx*)lParam;
    DWORD winPid = 0;
    GetWindowThreadProcessId(hwnd, &winPid);
    if (winPid == ctx->pid && IsWindowVisible(hwnd) && GetParent(hwnd) == NULL) {
        ctx->result = hwnd;
        return FALSE;
    }
    return TRUE;
}
void vb6_AppActivate(BSTR title, int32_t wait) {
    /* AppActivate: activate window by title or processID */
    (void)wait;
    if (!title) return;
    
    /* Check if title is a numeric processID (pure digits) */
    int32_t isPid = 1;
    int32_t pidVal = 0;
    size_t tLen = vb6_BSTR_Len(title);
    if (tLen == 0) return;
    for (size_t i = 0; i < tLen; i++) {
        if (title[i] >= L'0' && title[i] <= L'9') {
            pidVal = pidVal * 10 + (title[i] - L'0');
        } else {
            isPid = 0;
            break;
        }
    }
    
    if (isPid && pidVal > 0) {
        vb6_AppActivateCtx ctx = { (DWORD)pidVal, NULL };
        EnumWindows(vb6_AppActivateEnumProc, (LPARAM)&ctx);
        if (ctx.result) SetForegroundWindow(ctx.result);
    } else {
        HWND hwnd = FindWindowW(NULL, title);
        if (hwnd) SetForegroundWindow(hwnd);
    }
}

void vb6_AppActivateByPid(int32_t pid, int32_t wait) {
    /* AppActivate by processID */
    (void)wait;
    if (pid <= 0) return;
    vb6_AppActivateCtx ctx = { (DWORD)pid, NULL };
    EnumWindows(vb6_AppActivateEnumProc, (LPARAM)&ctx);
    if (ctx.result) SetForegroundWindow(ctx.result);
}
vb6_VARIANT vb6_CDec(vb6_VARIANT v) {
    /* P20-07: CDec - convert to Decimal (VT_DECIMAL=14) */
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    result.vt = vb6_vtDecimal;
#ifndef _WIN64
    /* Fix <vbeclipse> rev11: x86 的 union 只 8 字节 (对齐 OLE VARIANT), DECIMAL
     * 以指针形式存放; 释放由 vb6_VariantClear 负责。 */
    result.pdecVal = (DECIMAL*)CoTaskMemAlloc(sizeof(DECIMAL));
    if (!result.pdecVal) return vb6_VariantEmpty();
    memset(result.pdecVal, 0, sizeof(DECIMAL));
#endif
    switch (v.vt) {
        case vb6_vtInteger: case vb6_vtLong: case vb6_vtByte: {
            int32_t ival = (v.vt == vb6_vtInteger) ? v.iVal : (v.vt == vb6_vtByte) ? (int32_t)v.bVal : v.lVal;
            vb6_VARIANT_DECVAL(result).Lo32 = (uint32_t)(ival < 0 ? -ival : ival);
            vb6_VARIANT_DECVAL(result).Mid32 = 0;
            vb6_VARIANT_DECVAL(result).Hi32 = 0;
            vb6_VARIANT_DECVAL(result).scale = 0;
            vb6_VARIANT_DECVAL(result).sign = (ival < 0) ? 0x80 : 0;
            break;
        }
        case vb6_vtSingle: case vb6_vtDouble: case vb6_vtCurrency: {
            double dval = (v.vt == vb6_vtSingle) ? (double)v.fltVal :
                          (v.vt == vb6_vtCurrency) ? (double)v.cyVal / 10000.0 : v.dblVal;
            DECIMAL winDec;
            if (VarDecFromR8(dval, &winDec) == S_OK) {
                memcpy(&vb6_VARIANT_DECVAL(result), &winDec, sizeof(winDec));
            } else {
                int64_t i64 = (int64_t)dval;
                vb6_VARIANT_DECVAL(result).Lo32 = (uint32_t)(i64 & 0xFFFFFFFF);
                vb6_VARIANT_DECVAL(result).Mid32 = (uint32_t)((i64 >> 32) & 0xFFFFFFFF);
                vb6_VARIANT_DECVAL(result).Hi32 = 0;
                vb6_VARIANT_DECVAL(result).scale = 0;
                vb6_VARIANT_DECVAL(result).sign = (dval < 0) ? 0x80 : 0;
            }
            break;
        }
        case vb6_vtBSTR: {
            DECIMAL winDec;
            if (v.bstrVal && VarDecFromStr(v.bstrVal, LOCALE_USER_DEFAULT, 0, &winDec) == S_OK) {
                memcpy(&vb6_VARIANT_DECVAL(result), &winDec, sizeof(winDec));
            }
            break;
        }
        case vb6_vtDecimal: {
#ifndef _WIN64
            if (result.pdecVal) { CoTaskMemFree(result.pdecVal); result.pdecVal = NULL; }
#endif
            result = v;
            break;
        }
        default: {
#ifndef _WIN64
            if (result.pdecVal) { CoTaskMemFree(result.pdecVal); result.pdecVal = NULL; }
#endif
            break;
        }
    }
    return result;
}

void vb6_MidSet(BSTR* target, int32_t start, int32_t len, BSTR replacement) {
    /* P18-A: Mid$ statement assignment — Mid$(target, start, len) = replacement */
    /* Replaces len characters of target starting at position start (1-based) */
    if (!target || !*target || start < 1) return;
    int32_t tLen = vb6_BSTR_Len(*target);
    if (start > tLen) return;
    int32_t rLen = replacement ? vb6_BSTR_Len(replacement) : 0;
    /* If len <= 0, use replacement length (VB6 behavior when length omitted) */
    if (len <= 0) len = rLen;
    /* Clamp to available characters */
    int32_t avail = tLen - start + 1;
    if (len > avail) len = avail;
    if (len > rLen) len = rLen;
    if (len <= 0) return;
    /* Build new string: prefix + replacement[0..len-1] + suffix */
    int32_t newLen = tLen;  /* Mid$ doesn't change length, only replaces in-place */
    BSTR result = SysAllocStringLen(NULL, newLen);
    if (!result) return;
    /* Copy prefix (before start) */
    if (start > 1) memcpy(result, *target, (start - 1) * sizeof(wchar_t));
    /* Copy replacement */
    if (replacement && len > 0) memcpy(result + start - 1, replacement, len * sizeof(wchar_t));
    /* Copy suffix (after start+len-1) */
    int32_t afterStart = start - 1 + len;
    if (afterStart < tLen) memcpy(result + afterStart, *target + afterStart, (tLen - afterStart) * sizeof(wchar_t));
    result[newLen] = L'\0';
    /* Replace target */
    SysFreeString(*target);
    *target = result;
}

// ============================================================
// P14.2.4: IIf / InputBox
// ============================================================

/* IIf: VB6内联条件函数 - 通过函数调用确保两个分支都求值(C函数参数求值顺序不影响"都求值")
 * VB6 IIf 不短路，两个分支都必须计算，然后按条件选一个
 * 通过函数调用而非C三元运算符实现，保证VB6语义正确 */
BSTR vb6_IIfBSTR(int32_t cond, BSTR truepart, BSTR falsepart) {
    BSTR result = cond ? truepart : falsepart;
    /* 释放未被选中的分支(防止BSTR泄漏) */
    if (cond) { vb6_BSTR_Free(falsepart); }
    else      { vb6_BSTR_Free(truepart); }
    return result;
}
int32_t vb6_IIfLong(int32_t cond, int32_t truepart, int32_t falsepart) {
    (void)falsepart;
    return cond ? truepart : falsepart;
}
double vb6_IIfDouble(int32_t cond, double truepart, double falsepart) {
    (void)falsepart;
    return cond ? truepart : falsepart;
}
vb6_VARIANT vb6_IIfVariant(int32_t cond, vb6_VARIANT truepart, vb6_VARIANT falsepart) {
    if (cond) { vb6_VariantClear(&falsepart); return truepart; }
    else      { vb6_VariantClear(&truepart);  return falsepart; }
}

/* ============================================================
 * Fix 161: 控制台输入原语 (stdin 侧的多语言/多来源正确处理)
 *
 * 背景 —— 输出侧 vb6_ConWriteHandle 早已按"控制台 vs 管道"分流:
 *   控制台 → WriteConsoleW (宽字符直通, 与 chcp 无关)
 *   管道   → 按 GetConsoleOutputCP() 转字节
 * 输入侧此前只有 vb6_InputBox 的裸 fgetws(stdin): CRT 一律按 UTF-16 解析,
 * 于是**管道/重定向输入的字节被当 UTF-16 码元拼接** ⇒ ASCII 'AB'(0x41,0x42)
 * 读成 U+4241, GBK/UTF-8 中文同样全错。这不是"显示乱码", 是真读错。
 *
 * 本组函数把输出侧同一判据 (GetConsoleMode) 用到 stdin:
 *   真控制台 → ReadConsoleW   (键盘本就是 UTF-16)
 *   管道/重定向 → ReadFile 拿字节 → MultiByteToWideChar(GetConsoleCP() ?: ACP)
 *                装不下 (如 437 要读中文) 退回 UTF-8
 * 与输出侧口径对称。
 * ============================================================ */

/* 判别 stdin 是否为真控制台 (键盘)。返回 1=控制台, 0=管道/重定向/无效。 */
int32_t vb6_ConInputIsConsole(void) {
#ifdef _WIN32
    DWORD mode = 0;
    HANDLE h = GetStdHandle(STD_INPUT_HANDLE);
    if (h == NULL || h == INVALID_HANDLE_VALUE) return 0;
    return GetConsoleMode(h, &mode) ? 1 : 0;
#else
    return isatty(0) ? 1 : 0;
#endif
}

/* 从 stdin 读一行: 返回不含行尾符的 BSTR; EOF 且无内容时返回空 BSTR。
 * 控制台与管道两条路径都收敛到"宽字符结果"。 */
BSTR vb6_ConReadLine(void) {
#ifdef _WIN32
    HANDLE hIn = GetStdHandle(STD_INPUT_HANDLE);
    DWORD mode = 0;
    if (hIn == NULL || hIn == INVALID_HANDLE_VALUE) return vb6_BSTR_Empty();

    /* 可增长宽字符缓冲 (动态拼接, 覆盖长行) */
    size_t cap = 128, len = 0;
    wchar_t* out = (wchar_t*)malloc(cap * sizeof(wchar_t));
    if (!out) return vb6_BSTR_Empty();

    if (GetConsoleMode(hIn, &mode)) {
        /* ---- 真控制台: 键盘输入本就是 UTF-16 码元 ---- */
        for (;;) {
            wchar_t wc = 0;
            DWORD got = 0;
            if (!ReadConsoleW(hIn, &wc, 1, &got, NULL) || got == 0) break; /* EOF */
            if (wc == L'\n' || wc == L'\r') break;                        /* 行尾 */
            if (len + 1 >= cap) {
                cap *= 2;
                wchar_t* nb = (wchar_t*)realloc(out, cap * sizeof(wchar_t));
                if (!nb) break;
                out = nb;
            }
            out[len++] = wc;
        }
    } else {
        /* ---- 管道/重定向: 字节流, 按控制台输入代码页解码 ----
         * 逐块读入原始字节, 行尾 LF 结束; 最后整体转宽字符 (避免多字节
         * 序列被块边界切断)。 */
        size_t bcap = 256, blen = 0;
        char* bytes = (char*)malloc(bcap);
        if (!bytes) { free(out); return vb6_BSTR_Empty(); }
        for (;;) {
            char c = 0;
            DWORD got = 0;
            if (!ReadFile(hIn, &c, 1, &got, NULL) || got == 0) break; /* EOF */
            if (c == '\n') break;
            if (c == '\r') continue;                                   /* 兼容 CRLF */
            if (blen + 1 >= bcap) {
                bcap *= 2;
                char* nb = (char*)realloc(bytes, bcap);
                if (!nb) break;
                bytes = nb;
            }
            bytes[blen++] = c;
        }
        if (blen > 0) {
            UINT cp = GetConsoleCP();
            int need;
            if (!cp) cp = GetACP();
            need = MultiByteToWideChar(cp, 0, bytes, (int)blen, NULL, 0);
            if (need <= 0) {
                /* 代码页装不下 → 退 UTF-8, 不丢字 */
                cp = CP_UTF8;
                need = MultiByteToWideChar(cp, 0, bytes, (int)blen, NULL, 0);
            }
            if (need > 0) {
                if ((size_t)need + 1 > cap) {
                    cap = (size_t)need + 1;
                    wchar_t* nb = (wchar_t*)realloc(out, cap * sizeof(wchar_t));
                    if (nb) out = nb; else need = 0;
                }
                if (need > 0) {
                    MultiByteToWideChar(cp, 0, bytes, (int)blen, out, need);
                    len = (size_t)need;
                }
            }
        }
        free(bytes);
    }

    out[len] = L'\0';
    BSTR result = vb6_BSTR_FromStr(out);
    free(out);
    return result;
#else
    /* 非 Windows: CRT 窄字节 + 按 locale, 走 fgetws 亦可; 保持简单 */
    wchar_t buf[1024];
    if (fgetws(buf, 1024, stdin)) {
        int32_t l = (int32_t)wcslen(buf);
        while (l > 0 && (buf[l-1] == L'\n' || buf[l-1] == L'\r')) buf[--l] = L'\0';
        return vb6_BSTR_FromStr(buf);
    }
    return vb6_BSTR_Empty();
#endif
}

/* 从 stdin 读一个字符 (键盘/管道同一路径, 取首字符)。EOF 返回空 BSTR。 */
BSTR vb6_ConReadKey(void) {
    BSTR line = vb6_ConReadLine();
    BSTR result;
    if (!line || vb6_BSTR_Len(line) == 0) {
        if (line) vb6_BSTR_Free(line);
        return vb6_BSTR_Empty();
    }
    /* 取首个字符 (宽字符安全的左截断) */
    result = vb6_Left(line, 1);
    vb6_BSTR_Free(line);
    return result;
}

/* ============================================================
 * Fix 161: Console 对象 (对齐 twinBASIC)
 *   Console.WriteLine / Console.Write  → 输出 (复用 vb6_ConWriteOutW)
 *   Console.ReadLine / .ReadKey / .Read → 输入 (复用 vb6_ConReadLine/ConReadKey)
 * 输出走控制台/管道分流 (WriteConsoleW vs 代码页字节), 输入同理 (ReadConsoleW
 * vs ReadFile+代码页), 与 DEB.PRINT / InputBox 同一口径, 中文不乱。
 * ============================================================ */

void vb6_Console_Write(BSTR s) {
    if (s) vb6_ConWriteOutW((const wchar_t*)s, (int)vb6_BSTR_Len(s));
}

void vb6_Console_WriteLine(BSTR s) {
    if (s) vb6_ConWriteOutW((const wchar_t*)s, (int)vb6_BSTR_Len(s));
    vb6_ConWriteOutW(L"\r\n", 2);
}

BSTR vb6_Console_ReadLine(void) {
    return vb6_ConReadLine();
}

BSTR vb6_Console_ReadKey(void) {
    return vb6_ConReadKey();
}

/* InputBox: 简化实现 - 使用控制台输入 (非GUI环境)
 * Fix 161: 改用 vb6_ConReadLine —— 原 fgetws(stdin) 在管道输入下必乱码。 */
BSTR vb6_InputBox(BSTR prompt, BSTR title, BSTR defaultstr, int32_t xpos, int32_t ypos, BSTR helpfile, int32_t context) {
    (void)title; (void)xpos; (void)ypos; (void)helpfile; (void)context;
    /* 输出提示 (走 Debug.Print 同一条控制台/管道分流路径, 中文不乱) */
    if (prompt) {
        vb6_ConWriteOutW((const wchar_t*)prompt, (int)vb6_BSTR_Len(prompt));
        vb6_ConWriteOutW(L"\r\n", 2);
    }
    /* 显示默认值提示 */
    if (defaultstr && vb6_BSTR_Len(defaultstr) > 0) {
        vb6_ConWriteOutW(L"[", 1);
        vb6_ConWriteOutW((const wchar_t*)defaultstr, (int)vb6_BSTR_Len(defaultstr));
        vb6_ConWriteOutW(L"] ", 2);
    }
    vb6_ConWriteOutW(L"> ", 2);

    /* 读取一行 (控制台/管道双路径正确解码) */
    BSTR line = vb6_ConReadLine();
    if (line && vb6_BSTR_Len(line) > 0) {
        return line;
    }
    if (line) vb6_BSTR_Free(line);
    /* 空行/EOF 则返回默认值 */
    if (defaultstr) return vb6_BSTR_FromStr(defaultstr);
    return vb6_BSTR_Empty();
}

// Fix 133: 把 int32 写回 Variant 变量 (Declare 标量 ByRef 出参回写用)
void vb6_VariantSetI4(vb6_VARIANT* v, int32_t x) {
    if (!v) return;
    v->vt = (vb6_vartype)VT_I4;
    v->lVal = x;
}

// Fix 133x: 把 intptr_t (LongPtr 64位指针/句柄) 写回 Variant 变量,
// 存为 VT_I8 保持 64 位完整 (Charts 2020 x64: GdipCreateFont ByRef mFont
// As LongPtr 出参, 变体 hFont = GDI+ Font 对象指针不能截断为 32 位).
void vb6_VariantSetI8(vb6_VARIANT* v, intptr_t x) {
    if (!v) return;
    v->vt = (vb6_vartype)VT_I8;
    v->llVal = (int64_t)x;
}
