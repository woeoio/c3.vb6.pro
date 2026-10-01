// vb6rtl_string.c - VB6 运行时库: 字符串家族：BSTR 操作 + 内置字符串函数 + 字符串补充 + Like + Partition
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 40~64 行
//   原第 65~283 行
//   原第 1606~1728 行
//   原第 1729~1789 行
//   原第 4667~4698 行

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
// BSTR 操作
// ============================================================

BSTR vb6_BSTR_Concat(BSTR a, BSTR b) {
    int32_t lenA = vb6_BSTR_Len(a);
    int32_t lenB = vb6_BSTR_Len(b);
    int32_t total = lenA + lenB;

    BSTR result = SysAllocStringLen(NULL, total);
    if (!result) return NULL;

    if (a) memcpy(result, a, lenA * sizeof(wchar_t));
    if (b) memcpy(result + lenA, b, lenB * sizeof(wchar_t));
    result[total] = L'\0';

    return result;
}

BSTR vb6_BSTR_ConcatFree(BSTR a, BSTR b) {
    BSTR result = vb6_BSTR_Concat(a, b);
    vb6_BSTR_Free(a);  /* 释放中间临时BSTR */
    return result;
}

// ============================================================
// 内置函数实现
// ============================================================

int32_t vb6_Len(BSTR s) {
    return vb6_BSTR_Len(s);
}

int32_t vb6_LenB_BSTR(BSTR s) {  // Fix 048: LenB for BSTR — byte length (wchar count * 2)
    return vb6_BSTR_Len(s) * (int32_t)sizeof(wchar_t);
}

BSTR vb6_Left(BSTR s, int32_t n) {
    if (!s || n <= 0) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    if (n > len) n = len;
    wchar_t* buf = (wchar_t*)malloc((n + 1) * sizeof(wchar_t));
    memcpy(buf, s, n * sizeof(wchar_t));
    buf[n] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

BSTR vb6_Right(BSTR s, int32_t n) {
    if (!s || n <= 0) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    if (n > len) n = len;
    return vb6_BSTR_FromStr(s + len - n);
}

BSTR vb6_Mid(BSTR s, int32_t start, int32_t len) {
    if (!s || start < 1) return vb6_BSTR_Empty();
    int32_t slen = vb6_BSTR_Len(s);
    int32_t offset = start - 1;  // VB6是1-based
    if (offset >= slen) return vb6_BSTR_Empty();
    // len < 0 是 cgen 对省略 length 参数 (Mid$(s, n)) 的哨兵 — 取到结尾
    if (len < 0 || offset + len > slen) len = slen - offset;
    wchar_t* buf = (wchar_t*)malloc((len + 1) * sizeof(wchar_t));
    memcpy(buf, s + offset, len * sizeof(wchar_t));
    buf[len] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

// ============================================================
// <vbeclipse>: vbTextCompare 共用核 —— InStr/InStrRev/Replace/Split/Filter/StrComp 共用
// ============================================================
// VB6 的 compare 形参此前在 6 个入口被静默丢掉: RTL 里 5 处 (void)compare, 加上
// InStr 的第 4 参被 codegen 截成 3 参 (Fix 041)。查找类按"大小写折叠后逐字符相等"
// 实现 (与 RTL 自身的 LCase/UCase 同为 towlower); StrComp 走 CompareStringW 的区域
// 语言序 (那才是 VB6 的 vbTextCompare)。compare: 0 = vbBinaryCompare, 非 0 = 文本。
// 判据载荷保持纯 ASCII 时结果与码页/区域无关 (ai/029 那条 ASCII 载荷纪律)。

// 在 haystack 的 [from0, len-nlen] 区间找 needle, 命中返回 0 基下标, 否则 -1。
int32_t vb6_TextFind(BSTR haystack, BSTR needle, int32_t from0, int32_t compare) {
    if (!haystack || !needle) return -1;
    int32_t hlen = vb6_BSTR_Len(haystack);
    int32_t nlen = vb6_BSTR_Len(needle);
    if (nlen == 0) return (from0 >= 0 && from0 <= hlen) ? from0 : -1;
    if (from0 < 0) from0 = 0;
    if (compare == 0) {
        for (int32_t i = from0; i <= hlen - nlen; i++) {
            if (memcmp(haystack + i, needle, (size_t)nlen * sizeof(wchar_t)) == 0) return i;
        }
        return -1;
    }
    for (int32_t i = from0; i <= hlen - nlen; i++) {
        int32_t k = 0;
        while (k < nlen && towlower(haystack[i + k]) == towlower(needle[k])) k++;
        if (k == nlen) return i;
    }
    return -1;
}

// 判定 needle 是否**正好**出现在 haystack 的 at0 处 (Split 的分隔符扫描那一形)。
int32_t vb6_TextMatchAt(BSTR haystack, int32_t at0, BSTR needle, int32_t compare) {
    if (!haystack || !needle) return 0;
    int32_t hlen = vb6_BSTR_Len(haystack);
    int32_t nlen = vb6_BSTR_Len(needle);
    if (nlen == 0 || at0 < 0 || at0 + nlen > hlen) return 0;
    if (compare == 0)
        return memcmp(haystack + at0, needle, (size_t)nlen * sizeof(wchar_t)) == 0;
    for (int32_t k = 0; k < nlen; k++)
        if (towlower(haystack[at0 + k]) != towlower(needle[k])) return 0;
    return 1;
}

// 同 vb6_TextFind, 但返回 <= from0 的**最大**命中下标 (InStrRev 那一形)。
int32_t vb6_TextFindRev(BSTR haystack, BSTR needle, int32_t from0, int32_t compare) {
    if (!haystack || !needle) return -1;
    int32_t hlen = vb6_BSTR_Len(haystack);
    int32_t nlen = vb6_BSTR_Len(needle);
    if (from0 > hlen - nlen) from0 = hlen - nlen;
    if (nlen == 0) return (from0 >= 0) ? from0 : -1;
    if (from0 < 0) return -1;
    for (int32_t i = from0; i >= 0; i--) {
        if (compare == 0) {
            if (memcmp(haystack + i, needle, (size_t)nlen * sizeof(wchar_t)) == 0) return i;
        } else {
            int32_t k = 0;
            while (k < nlen && towlower(haystack[i + k]) == towlower(needle[k])) k++;
            if (k == nlen) return i;
        }
    }
    return -1;
}

#ifdef vb6_InStr
#undef vb6_InStr   // vb6rtl_builtin.h 的 _Generic 宏在定义处必须关闭, 否则本函数定义被改写
#endif
#ifdef vb6_InStrC
#undef vb6_InStrC  // 同上 (<vbeclipse>: 4 参形的 _Generic 分派宏)
#endif

// compare 版实现 (vbBinaryCompare = 0 走原 memcmp 路, 行为逐字不变)。
int32_t vb6_InStrC(int32_t start, BSTR haystack, BSTR needle, int32_t compare) {
    if (!haystack || !needle) return 0;
    int32_t nlen = vb6_BSTR_Len(needle);
    if (nlen == 0) return start;
    if (start < 1) start = 1;
    int32_t i = vb6_TextFind(haystack, needle, start - 1, compare);
    return (i < 0) ? 0 : i + 1;   // VB6 的 InStr 返回 1 基
}

int32_t vb6_InStr(int32_t start, BSTR haystack, BSTR needle) {
    return vb6_InStrC(start, haystack, needle, 0);
}

// Fix 158s: InStr 的 needle 实参是 vb6_VARIANT 时 (For 循环里拿 Variant 判定包含
// 关系, 如 VBFlexGrid DecStr/ExportString 一系) → 生成代码裸发
// vb6_InStr(1, Buffer, Value), Value 是 vb6_VARIANT 结构体 → MSVC C2440
// "无法从 vb6_VARIANT 转换到 BSTR"。vb6rtl_builtin.h 用 _Generic 在调用点分派到
// 这里, 先 vb6_VariantToString 解包再走标准 InStr。返回类型同为 int32_t, 与
// vb6_InStr 一致 (这一点与 158t 的 vb6_Mid 提案不同 —— 那里分派目标返回
// int32_t 却被用作 BSTR, 语义不成立, 已弃用并在 codegen 侧另修)。
int32_t vb6_InStrVar(int32_t start, BSTR haystack, vb6_VARIANT needle) {
    BSTR s = vb6_VariantToString(needle);
    int32_t r = vb6_InStrC(start, haystack, s, 0);
    vb6_BSTR_Free(s);
    return r;
}

// <vbeclipse>: 4 参形的同款 Variant 解包 (InStr 的字符串优先三参形与四参形都走这里)。
int32_t vb6_InStrVarC(int32_t start, BSTR haystack, vb6_VARIANT needle, int32_t compare) {
    BSTR s = vb6_VariantToString(needle);
    int32_t r = vb6_InStrC(start, haystack, s, compare);
    vb6_BSTR_Free(s);
    return r;
}


// Fix 093a: InStrB — VB6 字节版 InStr, 支持两种实参形态:
//  · Byte() 一维数组 (vb6_SafeArray1D*) — VB6 允许 InStrB(ByteArray1, ByteArray2)
//    做字节流查找 (类里常写成 `InStrB(a, b) = 1` 判两个字节数组相等);
//  · BSTR — 返回字节偏移 (VB6 InStrB 语义: 宽字符下标 * 2 + 1).
// 用 SafeArray1D 魔数 (0x5A1D) 区分两种载体, 避免在 RTL 里区分运行时指针类型.
int32_t vb6_InStrB(int32_t start, void* haystack, void* needle) {
    if (!haystack || !needle) return 0;
    if (start < 1) start = 1;
    {
        const vb6_SafeArray1D* ha = (const vb6_SafeArray1D*)haystack;
        const vb6_SafeArray1D* ne = (const vb6_SafeArray1D*)needle;
        if (ha->signature == 0x5A1D && ne->signature == 0x5A1D) {
            int32_t hlen = ha->count * ha->elemSize;
            int32_t nlen = ne->count * ne->elemSize;
            if (nlen == 0) return start;
            if (!ha->data || !ne->data) return 0;
            for (int32_t i = start - 1; i <= hlen - nlen; i++) {
                if (memcmp((const uint8_t*)ha->data + i, ne->data, (size_t)nlen) == 0) {
                    return i + 1;  // 1-based
                }
            }
            return 0;
        }
    }
    {
        BSTR h = (BSTR)haystack;
        BSTR n = (BSTR)needle;
        int32_t hlen = vb6_BSTR_Len(h);
        int32_t nlen = vb6_BSTR_Len(n);
        if (nlen == 0) return start;
        for (int32_t i = start - 1; i <= hlen - nlen; i++) {
            if (memcmp(h + i, n, (size_t)nlen * sizeof(wchar_t)) == 0) {
                return i * 2 + 1;  // 字节偏移
            }
        }
        return 0;
    }
}

BSTR vb6_UCase(BSTR s) {
    if (!s) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    wchar_t* buf = (wchar_t*)malloc((len + 1) * sizeof(wchar_t));
    for (int32_t i = 0; i < len; i++) {
        buf[i] = towupper(s[i]);
    }
    buf[len] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

BSTR vb6_LCase(BSTR s) {
    if (!s) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    wchar_t* buf = (wchar_t*)malloc((len + 1) * sizeof(wchar_t));
    for (int32_t i = 0; i < len; i++) {
        buf[i] = towlower(s[i]);
    }
    buf[len] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

BSTR vb6_Trim(BSTR s) {
    if (!s) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    int32_t start = 0, end = len;
    while (start < len && iswspace(s[start])) start++;
    while (end > start && iswspace(s[end - 1])) end--;
    int32_t trimmed = end - start;
    wchar_t* buf = (wchar_t*)malloc((trimmed + 1) * sizeof(wchar_t));
    memcpy(buf, s + start, trimmed * sizeof(wchar_t));
    buf[trimmed] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

// Fix 160w: 定义真函数, 宏展开时 default: 分支经 _Generic 后又展开成同名宏的
// 保护方式同 InStr: 宏在自身展开期间被禁用, 这里显式 #undef 关闭分派 (同下方
// vb6_LTrimVar 定义). 改真函数与其子串拷贝语义.
#undef vb6_LTrim
BSTR vb6_LTrim(BSTR s) {
    if (!s) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    int32_t start = 0;
    while (start < len && iswspace(s[start])) start++;
    return vb6_BSTR_FromStr(s + start);
}

// Fix 160w: LTrim(<Variant>) — 先 vb6_VariantToString 解包再走真函数.
BSTR vb6_LTrimVar(vb6_VARIANT s) {
    return vb6_LTrim(vb6_VariantToString(s));
}

BSTR vb6_RTrim(BSTR s) {
    if (!s) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    while (len > 0 && iswspace(s[len - 1])) len--;
    wchar_t* buf = (wchar_t*)malloc((len + 1) * sizeof(wchar_t));
    memcpy(buf, s, len * sizeof(wchar_t));
    buf[len] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

BSTR vb6_Chr(int32_t code) {
    wchar_t buf[2] = { (wchar_t)code, L'\0' };
    return vb6_BSTR_FromStr(buf);
}

int32_t vb6_Asc(BSTR s) {
    if (!s || vb6_BSTR_Len(s) == 0) return 0;
    return (int32_t)s[0];
}

double vb6_Val(BSTR s) {
    if (!s) return 0.0;
    // 简化: 转为窄字符串用strtod
    char narrow[256];
    size_t len = vb6_BSTR_Len(s);
    if (len > 255) len = 255;
    for (size_t i = 0; i < len; i++) narrow[i] = (char)s[i];
    narrow[len] = '\0';
    // czUI fix: VB6 的 Val 同样支持 &H十六进制 / &O八进制 字面量
    // (HexToColor 用 Val("&H" & hex) 解析 "#RRGGBB"; strtod 对 "&H" 返回 0,
    //  曾导致所有字符串颜色变纯黑 — 标题栏/面板/文本框全黑)
    const char* pv = narrow;
    while (*pv == ' ' || *pv == '\t') pv++;
    if (pv[0] == '&' && (pv[1] == 'H' || pv[1] == 'h')) {
        unsigned long long hv = 0;
        const char* q = pv + 2;
        while (isxdigit((unsigned char)*q)) {
            char c = *q;
            int d = (c <= '9') ? (c - '0') : ((c | 0x20) - 'a' + 10);
            hv = hv * 16 + (unsigned)d;
            q++;
        }
        return (double)hv;
    }
    if (pv[0] == '&' && (pv[1] == 'O' || pv[1] == 'o')) {
        unsigned long long ov = 0;
        const char* q = pv + 2;
        while (*q >= '0' && *q <= '7') { ov = ov * 8 + (unsigned)(*q - '0'); q++; }
        return (double)ov;
    }
    return strtod(narrow, NULL);
}

// Fix 090d: VB6 类型转换函数 (CByte/CInt/CLng/CDbl/CSng/CBool) 的字符串解析语义.
// 与 Val 不同, 类型转换识别 &H(hex)/&O(octal) 前缀, 十进制按前缀数字解析,
// 尾部无效字符忽略 (如 CByte("&HFF")=255, CByte("123abc")=123).
double vb6_NumVal(BSTR s) {
    if (!s) return 0.0;
    int32_t len = vb6_BSTR_Len(s);
    const wchar_t* p = s;
    int32_t i = 0;
    while (i < len && iswspace(p[i])) i++;
    if (i + 2 < len && p[i] == L'&') {
        wchar_t pre = p[i + 1];
        if (pre == L'H' || pre == L'h') {
            wchar_t tmp[64];
            int32_t n = len - (i + 2);
            if (n > 63) n = 63;
            for (int32_t k = 0; k < n; k++) tmp[k] = p[i + 2 + k];
            tmp[n] = L'\0';
            return (double)(int32_t)wcstoul(tmp, NULL, 16);
        }
        if (pre == L'O' || pre == L'o') {
            wchar_t tmp[64];
            int32_t n = len - (i + 2);
            if (n > 63) n = 63;
            for (int32_t k = 0; k < n; k++) tmp[k] = p[i + 2 + k];
            tmp[n] = L'\0';
            return (double)(int32_t)wcstoul(tmp, NULL, 8);
        }
    }
    wchar_t* end = NULL;
    double d = wcstod(p + i, &end);
    return d;
}

BSTR vb6_Str(int32_t n) {
    wchar_t buf[32];
    swprintf(buf, 32, L"%d", n);
    return vb6_BSTR_FromStr(buf);
}

// ============================================================
// 字符串函数 (补充)
// ============================================================

BSTR vb6_Replace(BSTR expr, BSTR find, BSTR rep, int32_t start, int32_t count, int32_t compare) {
    if (!expr || !find) return expr ? vb6_BSTR_FromStr(expr) : vb6_BSTR_Empty();
    int32_t exprLen = vb6_BSTR_Len(expr);
    int32_t findLen = vb6_BSTR_Len(find);
    int32_t repLen = rep ? vb6_BSTR_Len(rep) : 0;
    if (findLen == 0 || exprLen == 0) return vb6_BSTR_FromStr(expr);

    if (start < 1) start = 1;
    if (start > exprLen) return vb6_BSTR_Empty();   // VB6: start 超出长度 -> 空串
    int32_t maxCount = (count == -1) ? INT32_MAX : count;

    // 计算结果长度
    // <vbeclipse>: 原用 wcsstr ⇒ compare 形参被丢掉 (vbTextCompare 静默按二进制比)。
    // 改走 vb6_TextFind, 它按 compare 分流且二进制形与 wcsstr 同结果。
    int32_t matches = 0;
    int32_t pos = start - 1;
    while (matches < maxCount) {
        int32_t found = vb6_TextFind(expr, find, pos, compare);
        if (found < 0) break;
        matches++;
        pos = found + findLen;
    }
    if (matches == 0) return vb6_BSTR_FromStr(expr);

    // Fix 195: 结果必须按 BSTR 的既有约定构造 —— 长度前缀是**字节数**且由 OLE
    // 分配器管理。原实现无条件 malloc + 写字符数, 在 Windows 上撞两个坑:
    //   ① vb6_BSTR_Len 走 SysStringLen (字节/2), 于是结果被读成一半长度
    //      (Replace("abc","b","") 得 "a" 而非 "ac");
    //   ② 这块内存之后会被 vb6_BSTR_Free -> SysFreeString 释放, 不是 OLE 分配的
    //      指针 -> 堆损坏。
    int32_t resultLen = exprLen + matches * (repLen - findLen);
#ifdef _WIN32
    BSTR result = SysAllocStringLen(NULL, (UINT)resultLen);
    if (!result) return NULL;
#else
    uint32_t* p = (uint32_t*)malloc(sizeof(uint32_t) + (resultLen + 1) * sizeof(wchar_t));
    if (!p) return NULL;
    *p = (uint32_t)resultLen;
    BSTR result = (BSTR)(p + 1);
    result[resultLen] = L'\0';
#endif

    // 执行替换
    // Fix 195: start 之前的字符要原样保留 (VB6: 只有 start 起的位置参与替换)。
    // 原实现从 pos=start-1 直接开写, 前缀从不复制, 于是
    // Replace("aaa","a","b",2) 得 "bb"+"未初始化堆字节" 而不是 "abb"。
    int32_t outPos = 0;
    int32_t preLen = start - 1;
    if (preLen > 0) {
        memcpy(result, expr, preLen * sizeof(wchar_t));
        outPos = preLen;
    }
    pos = start - 1;
    int32_t done = 0;
    while (done < matches) {
        // <vbeclipse>: 构造段必须与上面的计数段用**同一个**查找核, 否则
        // vbTextCompare 下两处命中位置不一致 → 结果长度与实际写入错位。
        int32_t foundIdx = vb6_TextFind(expr, find, pos, compare);
        if (foundIdx < 0) break;
        int32_t beforeLen = foundIdx - pos;
        if (beforeLen > 0) {
            memcpy(result + outPos, expr + pos, beforeLen * sizeof(wchar_t));
            outPos += beforeLen;
        }
        if (repLen > 0) {
            memcpy(result + outPos, rep, repLen * sizeof(wchar_t));
            outPos += repLen;
        }
        pos = foundIdx + findLen;
        done++;
    }
    // 剩余部分
    int32_t remain = exprLen - pos;
    if (remain > 0) {
        memcpy(result + outPos, expr + pos, remain * sizeof(wchar_t));
        outPos += remain;
    }
    // 防御: 任何长度不一致都不许把未初始化的堆内存当字符串吐出去
    if (outPos < resultLen) {
        memset(result + outPos, 0, (size_t)(resultLen - outPos) * sizeof(wchar_t));
    }
    result[resultLen] = L'\0';
    return result;
}

BSTR vb6_Space(int32_t n) {
    if (n <= 0) return vb6_BSTR_Empty();
    wchar_t* buf = (wchar_t*)malloc((n + 1) * sizeof(wchar_t));
    for (int32_t i = 0; i < n; i++) buf[i] = L' ';
    buf[n] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

BSTR vb6_String(int32_t n, int32_t charCode) {
    if (n <= 0) return vb6_BSTR_Empty();
#ifdef _WIN32
    // Fix 161b-decl-out: 必须用 SysAllocStringLen 造**定长** BSTR, 不能走
    // vb6_BSTR_FromStr (= SysAllocString)。后者按 NUL 定长 —— `String$(260, 0)`
    // (VB6 里声明 Declare 出参缓冲的标准写法) 首字符就是 L'\0', SysAllocString
    // 会返回**长度 0** 的 BSTR。此后 Declare A 版 API (GetUserNameA/
    // GetModuleFileNameA) 按 nSize 往缓冲里写十几~几十字节 ⇒ 堆越界写 ⇒
    // STATUS_HEAP_CORRUPTION (0xC0000374), 且崩溃点漂移 (实测缓冲越小越易命中)。
    // 注意 charCode 可以是 0 —— 定长语义必须保住这 n 个 L'\0' (VB6 里 String$(n,0)
    // 是"n 个 NUL 的字符串", Len() 也是 n)。
    {
        BSTR result = SysAllocStringLen(NULL, (UINT)n);
        if (!result) return vb6_BSTR_Empty();
        for (int32_t i = 0; i < n; i++) result[i] = (wchar_t)charCode;
        result[n] = L'\0';
        return result;
    }
#else
    wchar_t* buf = (wchar_t*)malloc((n + 1) * sizeof(wchar_t));
    for (int32_t i = 0; i < n; i++) buf[i] = (wchar_t)charCode;
    buf[n] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
#endif
}

// P21-29: Spc — 打印定位函数 (VB6: Print #1, Spc(5); "x"), 返回 n 个空格
// 声明见 vb6rtl_class_com.h. 此前只有声明没有定义 → 用到即 LNK2019.
BSTR vb6_Spc(int32_t count) {
    return vb6_Space(count);
}

// P21-28: Tab — 打印定位函数 (VB6: Print #1, Tab(10); "x"), 定位到第 column 列 (1-based)
// 当前无打印列位置跟踪, 按"从第 1 列起算"补 column-1 个空格; VB6 中 column 不大于
// 当前列时会先换行, 该行为未建模。
BSTR vb6_Tab(int32_t column) {
    if (column <= 1) return vb6_BSTR_Empty();
    return vb6_Space(column - 1);
}

// <vbeclipse>: 文本比较的**唯一**实现口径 —— VB6 的 vbTextCompare 与
// `Option Compare Text` 是同一种比较, 必须走同一条路 (否则 `a = b` 与
// `StrComp(a,b,vbTextCompare)=0` 会给出不同答案)。CompareStringW 才是 VB6 用的
// 区域性语言序; 非 Windows / API 失败退到逐字符大小写折叠 (不引 _wcsicmp,
// 那是不带 C 标准的 MSVC 扩展, 会让非 Win 目标编不过)。
int vb6_TextCmp(const wchar_t* a, const wchar_t* b) {
    if (!a) a = L"";
    if (!b) b = L"";
#ifdef _WIN32
    int r = CompareStringW(LOCALE_USER_DEFAULT, NORM_IGNORECASE,
                           (LPCWCH)a, -1, (LPCWCH)b, -1);
    if (r == 1) return -1;
    if (r >= 2) return (int)r - 2;   // 2=相等 3=大于 → 0/1
#endif
    for (;;) {
        wchar_t ca = (wchar_t)towlower(*a), cb = (wchar_t)towlower(*b);
        if (ca != cb) return (ca < cb) ? -1 : 1;
        if (ca == L'\0') return 0;
        a++; b++;
    }
}

int32_t vb6_StrComp(BSTR s1, BSTR s2, int32_t compare) {
    // Fix 173: NULL BSTR (vbNullString / 未赋值的 String) 与 L"" 等价 —— 原来对
    // 单边 NULL 直接返回 ±1, 于是 `StrComp("", vbNullString)` 报"不等"。
    if (getenv("C3_STRCMP_TRACE")) {
        int32_t tl1 = s1 ? vb6_BSTR_Len(s1) : -1;
        int32_t tl2 = s2 ? vb6_BSTR_Len(s2) : -1;
        fprintf(stderr, "[SC] s1=%p len=%d <%ls> | s2=%p len=%d <%ls> cmp=%d\n",
                (void*)s1, tl1, s1 ? s1 : L"(null)",
                (void*)s2, tl2, s2 ? s2 : L"(null)", (int)compare);
    }
    if (!s1) s1 = L"";
    if (!s2) s2 = L"";
    int32_t len1 = vb6_BSTR_Len(s1);
    int32_t len2 = vb6_BSTR_Len(s2);
    // <vbeclipse>: vbTextCompare 此前被 (void)compare 丢掉 ⇒
    //   StrComp("a","B",vbTextCompare) 答"大于"(ASCII 97>66), 而 VB6 答"小于"。
    if (compare != 0) return vb6_TextCmp(s1, s2);
    int32_t minLen = (len1 < len2) ? len1 : len2;
    int cmp = memcmp(s1, s2, minLen * sizeof(wchar_t));
    if (cmp != 0) return (cmp < 0) ? -1 : 1;
    if (len1 < len2) return -1;
    if (len1 > len2) return 1;
    return 0;
}

BSTR vb6_StrReverse(BSTR s) {
    if (!s) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(s);
    wchar_t* buf = (wchar_t*)malloc((len + 1) * sizeof(wchar_t));
    for (int32_t i = 0; i < len; i++) buf[i] = s[len - 1 - i];
    buf[len] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

int32_t vb6_InStrRev(BSTR haystack, BSTR needle, int32_t start, int32_t compare) {
    if (!haystack || !needle) return 0;
    int32_t hLen = vb6_BSTR_Len(haystack);
    int32_t nLen = vb6_BSTR_Len(needle);
    if (nLen == 0) return hLen;
    if (nLen > hLen) return 0;
    if (start <= 0 || start > hLen) start = hLen;
    // <vbeclipse>: compare 形参此前被 (void) 丢掉; 改走共用查找核 (二进制形逐字不变)。
    int32_t i = vb6_TextFindRev(haystack, needle, start - nLen, compare);
    return (i < 0) ? 0 : i + 1;   // VB6 的 1 基
}

BSTR vb6_LCase_str(BSTR s) { return vb6_LCase(s); }  // 别名
BSTR vb6_UCase_str(BSTR s) { return vb6_UCase(s); }

// ============================================================
// P14.2.1: Like运算符 - 通配符模式匹配
// 支持: ? (单字符), * (零或多个字符), # (单个数字),
//       [charlist] (字符列表), [!charlist] (排除字符列表)
// ============================================================

// <vbeclipse>: mode != 0 ⇒ 文本形 (Option Compare Text / vbTextCompare 同口径):
// 逐字符比较与字符区间都按大小写折叠取。旧实现只有二进制一形, 于是
// "abc" Like "[A-Z]*" 在 Option Compare Text 模块里 VB6 答 True 而 C3 答 False。
static int likeMatch(const wchar_t* src, const wchar_t* pat, int mode) {
    wchar_t sc;
    while (*pat) {
        if (*pat == L'*') {
            while (*pat == L'*') pat++;
            if (*pat == L'\0') return 1;
            while (*src) {
                if (likeMatch(src, pat, mode)) return 1;
                src++;
            }
            return likeMatch(src, pat, mode);
        }
        else if (*pat == L'?') {
            if (*src == L'\0') return 0;
            src++; pat++;
        }
        else if (*pat == L'#') {
            if (*src == L'\0') return 0;
            if (*src < L'0' || *src > L'9') return 0;
            src++; pat++;
        }
        else if (*pat == L'[') {
            pat++;
            int negate = 0;
            if (*pat == L'!') { negate = 1; pat++; }
            int match = 0;
            if (*src == L'\0') return 0;
            sc = mode ? (wchar_t)towlower(*src) : *src;
            while (*pat && *pat != L']') {
                if (pat[1] == L'-' && pat[2] && pat[2] != L']') {
                    wchar_t lo = *pat, hi = pat[2];
                    if (mode) { lo = (wchar_t)towlower(lo); hi = (wchar_t)towlower(hi); }
                    if (sc >= lo && sc <= hi) match = 1;
                    pat += 3;
                } else {
                    wchar_t pc = mode ? (wchar_t)towlower(*pat) : *pat;
                    if (sc == pc) match = 1;
                    pat++;
                }
            }
            if (*pat == L']') pat++;
            if (negate) match = !match;
            if (!match) return 0;
            src++;
        }
        else {
            wchar_t pc = mode ? (wchar_t)towlower(*pat) : *pat;
            wchar_t cc = mode ? (wchar_t)towlower(*src) : *src;
            if (cc != pc) return 0;
            src++; pat++;
        }
    }
    return (*src == L'\0');
}

int16_t vb6_LikeC(BSTR source, BSTR pattern, int32_t mode) {
    const wchar_t* s = source ? source : L"";
    const wchar_t* p = pattern ? pattern : L"";
    return likeMatch(s, p, mode ? 1 : 0) ? -1 : 0;
}

int16_t vb6_Like(BSTR source, BSTR pattern) {
    return vb6_LikeC(source, pattern, 0);
}

// ============================================================
// P18-E: Partition Function
// ============================================================

BSTR vb6_Partition(int64_t number, int64_t start, int64_t stop, int64_t interval) {
    int64_t rangeStart = 0, rangeEnd = 0;
    int leftBlank = 0, rightBlank = 0;

    if (number < start) {
        leftBlank = 1;
        rangeEnd = start - 1;
    } else if (number > stop) {
        rightBlank = 1;
        rangeStart = stop + 1;
    } else {
        int64_t idx = (number - start) / interval;
        rangeStart = start + idx * interval;
        rangeEnd = rangeStart + interval - 1;
        if (rangeEnd > stop) rangeEnd = stop;
    }

    wchar_t buf[64];
    if (leftBlank) {
        swprintf(buf, 64, L"%10ls: %10lld", L"", (long long)rangeEnd);
    } else if (rightBlank) {
        swprintf(buf, 64, L"%10lld: %10ls", (long long)rangeStart, L"");
    } else {
        swprintf(buf, 64, L"%10lld: %10lld", (long long)rangeStart, (long long)rangeEnd);
    }
    return vb6_BSTR_FromStr(buf);
}

