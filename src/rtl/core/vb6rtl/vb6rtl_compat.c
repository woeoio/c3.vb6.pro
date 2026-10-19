// vb6rtl_compat.c - VB6 运行时库: 兼容填平家族：P18-D 字符串/指针/格式化函数 + Split/Join 数组-字符串互转
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 2237~2511 行
//   原第 2568~2679 行

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
// P18-D: 兼容性填平 — 字符串/指针/格式化函数
// ============================================================

int32_t vb6_AscW(BSTR s) {
    if (!s || vb6_BSTR_Len(s) == 0) return 0;
    return (int32_t)s[0];  // Unicode code point (same as Asc for BMP)
}

BSTR vb6_ChrW(int32_t code) {
    wchar_t buf[2] = { (wchar_t)code, L'\0' };
    return vb6_BSTR_FromStr(buf);
}

int32_t vb6_AscB(BSTR s) {
    if (!s || vb6_BSTR_Len(s) == 0) return 0;
    return (int32_t)(s[0] & 0xFF);  // Low byte of first character
}

BSTR vb6_ChrB(int32_t code) {
    wchar_t buf[2] = { (wchar_t)(code & 0xFF), L'\0' };
    return vb6_BSTR_FromStr(buf);
}

double vb6_Timer(void) {
    SYSTEMTIME st;
    GetLocalTime(&st);
    return (double)st.wHour * 3600.0 + (double)st.wMinute * 60.0 +
           (double)st.wSecond + (double)st.wMilliseconds / 1000.0;
}

BSTR vb6_StrConv(BSTR text, int32_t conversion, int32_t localeID) {
    if (!text) return vb6_BSTR_Empty();
    int32_t len = vb6_BSTR_Len(text);
    if (len == 0) return vb6_BSTR_Empty();

    /* vbUpperCase=1, vbLowerCase=2, vbProperCase=3 - simple wchar transforms */
    if (conversion == 1 || conversion == 2 || conversion == 3) {
        wchar_t* buf = (wchar_t*)malloc((len + 1) * sizeof(wchar_t));
        if (!buf) return vb6_BSTR_Empty();
        memcpy(buf, text, len * sizeof(wchar_t));
        buf[len] = L'\0';
        if (conversion == 1) {
            for (int i = 0; i < len; i++) buf[i] = towupper(buf[i]);
        } else if (conversion == 2) {
            for (int i = 0; i < len; i++) buf[i] = towlower(buf[i]);
        } else { /* vbProperCase */
            int capNext = 1;
            for (int i = 0; i < len; i++) {
                if (iswspace(buf[i])) { capNext = 1; }
                else if (capNext) { buf[i] = towupper(buf[i]); capNext = 0; }
                else { buf[i] = towlower(buf[i]); }
            }
        }
        BSTR result = vb6_BSTR_FromStr(buf);
        free(buf);
        return result;
    }

    /* vbWide=4, vbNarrow=8, vbKatakana=16, vbHiragana=32
       Use LCMapStringEx for CJK locale-aware conversions.
       These can be combined (e.g. vbWide+vbKatakana = 4+16 = 20).
       LCMapStringEx flag mapping:
         vbWide     -> LCMAP_FULLWIDTH     (0x00800000)
         vbNarrow   -> LCMAP_HALFWIDTH     (0x00400000)
         vbKatakana -> LCMAP_KATAKANA      (0x00200000)
         vbHiragana -> LCMAP_HIRAGANA      (0x00100000)
    */
    if (conversion & 0x3C) { /* bits 2-5: wide/narrow/katakana/hiragana */
        DWORD mapFlags = 0;
        if (conversion & 4)  mapFlags |= 0x00800000; /* LCMAP_FULLWIDTH */
        if (conversion & 8)  mapFlags |= 0x00400000; /* LCMAP_HALFWIDTH */
        if (conversion & 16) mapFlags |= 0x00200000; /* LCMAP_KATAKANA */
        if (conversion & 32) mapFlags |= 0x00100000; /* LCMAP_HIRAGANA */
        if (mapFlags == 0) goto strconv_unicode;

        /* Determine locale name from localeID */
        wchar_t localeName[85];
        if (localeID == 0) localeID = 0x0411; /* default to Japanese */
        if (!LCIDToLocaleName((DWORD)localeID, localeName, 84, 0)) {
            wcscpy(localeName, L"ja-JP"); /* fallback */
        }

        /* First call: get required buffer size */
        int outLen = LCMapStringEx(localeName, mapFlags, text, len, NULL, 0, NULL, NULL, 0);
        if (outLen <= 0) {
            /* LCMapStringEx failed - return original string unchanged */
            return SysAllocStringLen(text, len);
        }

        wchar_t* outBuf = (wchar_t*)malloc((outLen + 1) * sizeof(wchar_t));
        if (!outBuf) return vb6_BSTR_Empty();

        /* Second call: perform the mapping */
        outLen = LCMapStringEx(localeName, mapFlags, text, len, outBuf, outLen, NULL, NULL, 0);
        if (outLen <= 0) {
            free(outBuf);
            return SysAllocStringLen(text, len);
        }
        outBuf[outLen] = L'\0';
        BSTR result = vb6_BSTR_FromStr(outBuf);
        free(outBuf);
        return result;
    }

strconv_unicode:
    /* vbUnicode=64, vbFromUnicode=128: no-op (internal strings are already Unicode) */
    (void)localeID;
    return SysAllocStringLen(text, len);
}

/* ============================================================
 * StrConv -> Byte() 语义 (twinbasic 兼容)
 * ============================================================ */

/* String -> Byte(): 复制 BSTR 原始 UTF-16LE 内存字节。
 * 例: "一二" (U+4E00, U+4E8C) -> 00 4E 8C 4E */
struct vb6_SafeArray1D* vb6_StringToByteArray(BSTR text) {
    if (!text) return NULL;
    int32_t wlen = (int32_t)SysStringLen(text);
    if (wlen <= 0) return NULL;
    int32_t byteCount = wlen * 2;
    struct vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_byte, 0, byteCount - 1);
    if (!arr || !arr->data) return arr;
    /* Fix 140: 逻辑长度不变 (count = byteCount), 但数据缓冲多分配 2 字节并清零,
     * 使 StrPtr(arr) 可当 null 结尾宽字符串使用 —— LabelPlus.ctl 用
     * GdipMeasureString/GdipAddPathString(StrPtr(m_Caption), -1, ...). */
    void* grown = realloc(arr->data, (size_t)byteCount + 2);
    if (grown) arr->data = grown;
    memcpy(arr->data, text, (size_t)byteCount);
    if (grown) memset((char*)arr->data + byteCount, 0, 2);
    return arr;
}

/* Fix 140: Byte() -> BSTR — 把数组数据当作 UTF-16LE 宽字符串还原.
 * 与 vb6_StringToByteArray 互逆 (原始内存字节, 不做 ANSI 转码). */
BSTR vb6_ByteArrayToString(struct vb6_SafeArray1D* arr) {
    if (!arr || !arr->data || arr->count <= 0) return vb6_BSTR_Empty();
    int32_t wcCount = arr->count / 2;
    const wchar_t* p = (const wchar_t*)arr->data;
    while (wcCount > 0 && p[wcCount - 1] == 0) wcCount--;   /* 去掉 0 结尾 */
    if (wcCount <= 0) return vb6_BSTR_Empty();
    BSTR result = SysAllocStringLen(NULL, (UINT)wcCount);
    if (!result) return vb6_BSTR_Empty();
    memcpy(result, arr->data, (size_t)wcCount * sizeof(wchar_t));
    return result;
}

/* Fix 140: StrPtr(Byte()) — 返回数组数据指针 (而非 SafeArray1D* 结构体指针). */
void* vb6_SafeArrayDataPtr(struct vb6_SafeArray1D* arr) {
    return arr ? arr->data : NULL;
}

/* Fix 140: Variant -> Byte() 数组.
 * Variant 持数组       -> 直接返回其 parray (与 vb6_VariantToSafeArray1D 同);
 * Variant 持字符串 BSTR -> 按 VB6 语义复制为字节数组 (原始 UTF-16LE 字节);
 * 其余 (Empty/Null/标量) -> NULL.
 * 场景: LabelPlus.ctl `m_Caption = .ReadProperty("Caption", Ambient.DisplayName)`
 * — PropertyBag 中 Caption 存的是字符串, 必须转成字节数组, 否则 UBound/StrPtr
 * 读不到内容, 右侧 "Venta diaria"/"USD $532.00" 卡片空白. */
struct vb6_SafeArray1D* vb6_VariantToByteArray(vb6_VARIANT v) {
    if ((v.vt & vb6_vtArray) && v.parray) {
        return v.parray;
    }
    if (v.vt == vb6_vtBSTR && v.bstrVal) {
        return vb6_StringToByteArray(v.bstrVal);
    }
    return NULL;
}

/* Fix 140: COM调用结果 -> Byte() 数组.
 * vb6_ComCall 返回 Windows VARIANT*; 用 vb6_VariantFromComResult 深转换并把
 * 内容所有权转移走 (BSTR/SAFEARRAY), 然后按 Byte() 语义解封. */
struct vb6_SafeArray1D* vb6_ComCallByteArray(void* disp, const wchar_t* methodName,
                                             void* args, int32_t argc) {
    void* pv = vb6_ComCall(disp, methodName, args, argc);
    if (!pv) return NULL;
    /* 注意: vb6_VariantFromComResult 内部已 VariantClear(pv) + free(pv)
     * (它"消耗" VARIANT* 所有权 — 见 vb6rtl_com.c)，此处绝不能再 free(pv),
     * 否则 double-free → STATUS_HEAP_CORRUPTION。 */
    vb6_VARIANT v = vb6_VariantFromComResult(pv);
    return vb6_VariantToByteArray(v);
}

/* StrConv(s, conversion) -> Byte() */
struct vb6_SafeArray1D* vb6_StrConvToByteArray(BSTR text, int32_t conversion, int32_t localeID) {
    if (!text) return NULL;
    int32_t wlen = (int32_t)SysStringLen(text);
    if (wlen <= 0) return NULL;

    if (conversion == 128) {
        /* vbFromUnicode: Unicode -> 系统 ANSI/GBK 字节。
         * 例: "一二" -> D2 BB B6 FE */
        int ansiLen = WideCharToMultiByte(CP_ACP, 0, text, wlen, NULL, 0, NULL, NULL);
        if (ansiLen <= 0) return NULL;
        struct vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_byte, 0, ansiLen - 1);
        if (!arr || !arr->data) return arr;
        WideCharToMultiByte(CP_ACP, 0, text, wlen, (char*)arr->data, ansiLen, NULL, NULL);
        return arr;
    }

    if (conversion == 64) {
        /* vbUnicode: 怪癖 -- 把 BSTR 的原始内存字节当成 ANSI/GBK 文本解码,
         * 再重建 UTF-16LE 字节。例: 原始 00 4E 8C 4E -> 00 00 4E 00 68 5B */
        int rawBytes = wlen * 2;
        char* raw = (char*)malloc((size_t)rawBytes);
        if (!raw) return NULL;
        memcpy(raw, text, (size_t)rawBytes);
        int wcCount = MultiByteToWideChar(CP_ACP, 0, raw, rawBytes, NULL, 0);
        if (wcCount <= 0) { free(raw); return NULL; }
        struct vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_byte, 0, wcCount * 2 - 1);
        if (!arr || !arr->data) { free(raw); return arr; }
        MultiByteToWideChar(CP_ACP, 0, raw, rawBytes, (wchar_t*)arr->data, wcCount);
        free(raw);
        return arr;
    }

    (void)localeID;
    /* 其他转换模式退化为原始内存复制 */
    return vb6_StringToByteArray(text);
}

/* Byte() -> BSTR: 反向转换, 供 s = StrConv(ba, vbUnicode) */
BSTR vb6_StrConvFromByteArray(struct vb6_SafeArray1D* arr, int32_t conversion, int32_t localeID) {
    if (!arr || !arr->data || arr->count <= 0) return vb6_BSTR_Empty();
    (void)localeID;

    if (conversion == 64) {
        /* vbUnicode: ANSI/GBK 字节 -> Unicode BSTR */
        const char* bytes = (const char*)arr->data;
        int byteCount = arr->count;
        int wcCount = MultiByteToWideChar(CP_ACP, 0, bytes, byteCount, NULL, 0);
        if (wcCount <= 0) return vb6_BSTR_Empty();
        BSTR result = SysAllocStringLen(NULL, wcCount);
        if (!result) return vb6_BSTR_Empty();
        MultiByteToWideChar(CP_ACP, 0, bytes, byteCount, result, wcCount);
        return result;
    }

    if (conversion == 128) {
        /* vbFromUnicode: UTF-16LE 字节 -> ANSI/GBK 字节串 (以 BSTR 承载单字节) */
        int wcCount = arr->count / 2;
        if (wcCount <= 0) return vb6_BSTR_Empty();
        int ansiLen = WideCharToMultiByte(CP_ACP, 0, (const wchar_t*)arr->data, wcCount,
                                          NULL, 0, NULL, NULL);
        if (ansiLen <= 0) return vb6_BSTR_Empty();
        char* ansi = (char*)malloc((size_t)ansiLen);
        if (!ansi) return vb6_BSTR_Empty();
        WideCharToMultiByte(CP_ACP, 0, (const wchar_t*)arr->data, wcCount,
                            ansi, ansiLen, NULL, NULL);
        BSTR result = SysAllocStringLen(NULL, ansiLen);
        if (!result) { free(ansi); return vb6_BSTR_Empty(); }
        for (int i = 0; i < ansiLen; i++) result[i] = (wchar_t)(unsigned char)ansi[i];
        free(ansi);
        return result;
    }

    return vb6_BSTR_Empty();
}


// <vbeclipse>: 按 elemType 读"字符串数组"的一个槽位。VB6 里 Join/Filter 的源数组必须是
// 字符串数组 —— 可以是 `Dim s() As String` (槽 = BSTR), 也可以是 `Array("a","b")` 这种
// Variant 数组 (槽 = 16 字节 vb6_VARIANT)。此前两处都无条件 `VB6_SA_AT(BSTR, …)` 取槽,
// 对 Variant 数组就把 VARIANT 的头 8 字节当 BSTR 指针用 → 解引用垃圾 → 0xC0000005
// (实测 `Join(Array("abc","xyz"), "|")` 直接段错误)。
// 元素不是字符串 ⇒ 与 VB6 一致抛 13 (Type mismatch): 有 On Error 走它的处理器,
// 没有则 vb6_ErrRaise 的未处理路径报错退出 (同 vb6_UBound 的 rev2 约定)。
static BSTR vb6_SA_ReadStrAt(struct vb6_SafeArray1D* arr, int32_t index) {
    void* slot = (char*)arr->data + (size_t)(index - arr->lBound) * (size_t)arr->elemSize;
    if (arr->elemType == vb6_sa_bstr) return *(BSTR*)slot;
    if (arr->elemType == vb6_sa_variant) {
        vb6_VARIANT* v = (vb6_VARIANT*)slot;
        if (v->vt == vb6_vtEmpty) return vb6_BSTR_Empty();
        if (v->vt == vb6_vtBSTR) return v->bstrVal;
    }
    vb6_ErrRaise(13, vb6_BSTR_FromStr(L"VBA.Information"),
                 vb6_BSTR_FromStr(L"Type mismatch"));
    return NULL;  /* 不可达: vb6_ErrRaise 必 longjmp 或 ExitProcess */
}

struct vb6_SafeArray1D* vb6_Filter(struct vb6_SafeArray1D* source, BSTR match, int32_t include, int32_t compare) {
    if (!source) return NULL;
    // <vbeclipse>: VB6 的 Filter 只吃**字符串数组** —— String() 数组 (vb6_sa_bstr) 或
    // Array("a","b") 这种 Variant 数组 (槽是 16 字节 vb6_VARIANT)。旧代码无条件按
    // BSTR* 取槽, 对 Variant 数组就把 VARIANT 的前 8 字节当指针用 → 垃圾 → 崩溃。
    // 命中判定走共用查找核 vb6_TextFind ⇒ vbTextCompare 时大小写不敏感。
    int32_t wantInclude = (include != 0) ? 1 : 0;  // VB6 的 True = -1, 按非零归一
    /* First pass: count matching elements */
    int32_t matchCount = 0;
    for (int32_t i = source->lBound; i <= source->uBound; i++) {
        BSTR elem = vb6_SA_ReadStrAt(source, i);
        int32_t found = 0;
        if (elem && match) {
            int32_t matchLen = vb6_BSTR_Len(match);
            if (matchLen == 0) { found = 1; }
            else if (vb6_TextFind(elem, match, 0, compare) >= 0) { found = 1; }
        }
        if (found == wantInclude) matchCount++;
    }
    /* Create result array */
    // 零命中 ⇒ VB6 的空数组: uBound = -1 (与 vb6_ArrayCreate(0)/Split 同形),
    // 旧写法夹到 0 ⇒ 凭空多出一个幻影元素, `UBound(x)+1` 读出 1。
    struct vb6_SafeArray1D* result = vb6_SafeArrayCreate1D(vb6_sa_bstr, 0, matchCount - 1);
    if (!result) return NULL;
    /* Second pass: copy matching elements — 判定必须与计数段同一个谓词 */
    int32_t idx = 0;
    for (int32_t i = source->lBound; i <= source->uBound; i++) {
        BSTR elem = vb6_SA_ReadStrAt(source, i);
        int32_t found = 0;
        if (elem && match) {
            int32_t matchLen = vb6_BSTR_Len(match);
            if (matchLen == 0) { found = 1; }
            else if (vb6_TextFind(elem, match, 0, compare) >= 0) { found = 1; }
        }
        if (found == wantInclude) {
            VB6_SA_AT(BSTR, result, idx + result->lBound) = elem ? SysAllocString(elem) : vb6_BSTR_Empty();
            idx++;
        }
    }
    return result;
}

// 账 #257 第十三层: StrPtr 的空串口径。
// VB6 里"空串"的 StrPtr 答 **0** —— 未赋值的 String、`vbNullString`、`""` 一律如此。
// 这不是实现细节而是 VB6 明文语义, 生态里两处标准写法都靠它:
//   · `If StrPtr(s) = 0 Then` 判"可选字符串省略了 / InputBox 被取消";
//   · Declare 侧 `If StrPtr(FileName) = NULL_PTR Then <用默认> Else <LoadLibrary(StrPtr)>`
//     (ComCtlsDemo Animation.LoadRes 实测就是这一处: 单参调用 `LoadRes ResID` 走了
//      LoadLibrary 分支 ⇒ NULL ⇒ Err.Raise 53)。
// ⚠ 判空**不能**看首字符: `String$(n, vbNullChar)` 首字符就是 NUL 但长度非 0, StrPtr 必须非 0。
// 一律按 BSTR 的**长度**判 (SysStringLen), 它同时也把 NULL 与"显式零长 BSTR"统一成同一档。
void* vb6_StrPtr(BSTR s) {
    if (vb6_BSTR_Len(s) == 0) return NULL;
    return (void*)s;  // BSTR points directly to string data
}

// 账 #305 §B180 (第六十九刀): StrPtr 收 **Variant** 时借槽里那枚 BSTR。
// 发码侧原来走 `vb6_StrPtr(vb6_VariantToString(v))`，而 `vb6_VariantToString` 的定义就一句
// `return vb6_CStr(v)` (vb6rtl.c:484) = `vb6_Format(x, NULL, 1, 1)` (vb6rtl_conv.c:253)
// ⇒ 每次调用**现造一枚新 BSTR** 把地址交出去，那枚串没人回收。实测 (`.build/k69sp`，
// 18000 次减去同形状普通 String 的零分配基线窗) = x64 1089536 / x86 909312 字节，
// 而真语料里这一形有 18 份文件 41 处，全坐在 Win32 的指针交接位上
// (`lpszIcon` / `pszText` / `CreateWindowExW` / `FindFirstFileW` / `VarI8FromStr` / `dwTypeData`)。
// 这一枚只在「`vt` 已是串」那一档改借 —— 空串/NULL 仍由上面的 `vb6_StrPtr` 统一答 0，
// 与账 #257 第十三层那条口径一字不差。⚠ **其余各档照旧 mint**：非串的 Variant 交给 StrPtr
// 本来就是「按串算」的强制转换 (VB6 那三枚指针函数是 undocumented，运行时里 `StrPtr` 那格的
// 形参声明就是 `[in] BSTR`)，那一档漏的那枚临时串归 §B164 形① 的带权临时量名单，不在本刀撤。
void* vb6_StrPtrVar(vb6_VARIANT v) {
    if (v.vt == (uint16_t)vb6_vtBSTR) return vb6_StrPtr(v.bstrVal);
    return vb6_StrPtr(vb6_CStr(v));
}

uintptr_t vb6_ObjPtr(void* obj) {
    return (uintptr_t)obj;
}

BSTR vb6_LSet(BSTR str, int32_t length) {
    if (length <= 0) return vb6_BSTR_Empty();
    wchar_t* buf = (wchar_t*)calloc(length + 1, sizeof(wchar_t));
    if (!buf) return vb6_BSTR_Empty();
    /* Fill with spaces */
    for (int i = 0; i < length; i++) buf[i] = L' ';
    /* Copy string (left-justified, truncated if too long) */
    if (str) {
        int32_t slen = vb6_BSTR_Len(str);
        if (slen > length) slen = length;
        memcpy(buf, str, slen * sizeof(wchar_t));
    }
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}

BSTR vb6_LSetFree(BSTR str, int32_t length) {
    /* 账 #118: 定长串赋值收口 —— LSet 同款结果, 顺带释放入参。
     * 发码侧只在"RHS 是自有临时串"那一侧发这个 (与 vb6_BSTR_AssignMove 的选择同一把尺),
     * 借用侧 (变量/形参/字段) 发裸 vb6_LSet —— 那时不能释放别人的串。 */
    BSTR r = vb6_LSet(str, length);
    vb6_BSTR_Free(str);
    return r;
}

BSTR vb6_RSet(BSTR str, int32_t length) {
    if (length <= 0) return vb6_BSTR_Empty();
    wchar_t* buf = (wchar_t*)calloc(length + 1, sizeof(wchar_t));
    if (!buf) return vb6_BSTR_Empty();
    /* Fill with spaces */
    for (int i = 0; i < length; i++) buf[i] = L' ';
    /* Copy string (right-justified, truncated if too long) */
    if (str) {
        int32_t slen = vb6_BSTR_Len(str);
        if (slen > length) slen = length;
        /* Right-justify: copy to end of buffer */
        memcpy(buf + (length - slen), str + (vb6_BSTR_Len(str) - slen), slen * sizeof(wchar_t));
    }
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    return result;
}


BSTR vb6_WeekdayName(int32_t weekday, int32_t abbreviate, int32_t firstDayOfWeek) {
    /* firstDayOfWeek: 1=Sunday(default), 2=Monday, ..., 7=Saturday
       weekday is relative to firstDayOfWeek */
    static const wchar_t* fullNames[] = { L"Sunday", L"Monday", L"Tuesday", L"Wednesday", L"Thursday", L"Friday", L"Saturday" };
    static const wchar_t* abbrNames[] = { L"Sun", L"Mon", L"Tue", L"Wed", L"Thu", L"Fri", L"Sat" };
    if (firstDayOfWeek < 1 || firstDayOfWeek > 7) firstDayOfWeek = 1;
    /* Convert weekday (relative to firstDayOfWeek) to absolute (1=Sunday) */
    int32_t idx = (weekday - 1 + (firstDayOfWeek - 1)) % 7;
    if (idx < 0) idx += 7;
    const wchar_t* name = abbreviate ? abbrNames[idx] : fullNames[idx];
    return vb6_BSTR_FromStr(name);
}

BSTR vb6_MonthName(int32_t month, int32_t abbreviate) {
    static const wchar_t* fullNames[] = { L"January", L"February", L"March", L"April", L"May", L"June",
        L"July", L"August", L"September", L"October", L"November", L"December" };
    static const wchar_t* abbrNames[] = { L"Jan", L"Feb", L"Mar", L"Apr", L"May", L"Jun",
        L"Jul", L"Aug", L"Sep", L"Oct", L"Nov", L"Dec" };
    if (month < 1 || month > 12) return vb6_BSTR_Empty();
    const wchar_t* name = abbreviate ? abbrNames[month - 1] : fullNames[month - 1];
    return vb6_BSTR_FromStr(name);
}

// Fix 161: 原先三个 FormatXxx 都用 `swprintf(buf, 64, L"%%.%df", numDigits)` 想拼出
// "%.2f" —— 但 `%%` 在 printf 家族里是**字面百分号**, 拼出的 buf 不是可再用的格式串。
// 实测 FormatPercent 把上次结果当格式反复解析 → "123456.7856.7856.78..." 无限重复。
// 正解 = 直接用宽度精度参数 `%.*f`, 一次成型, 不再二次套用。
//
// 账 #439 (§B212，第一百一十二刀) = 「UseParentheses 这一位要什么」这一问，**收成一位主人**。
// 病机制：三家（FormatNumber / FormatCurrency / FormatPercent）各自的出口第一行都写着
//   `(void)incLeading; (void)useParens; (void)groupDigits;` —— 这一位被**明确丢掉**，
//   于是负数永远交 `-1,234.57`，而运行期在 useParens 为真时交 `(1,234.57)`。
//   实测（cscript 同机代理，尺 .build/kpar2.vbs）三家口径**完全一致**:
//       省略(默认 -1) → minus=True  parens=False
//       useParens=True → minus=False parens=True
//       useParens=0   → minus=True  parens=False
//       正数 + True    → minus=False parens=False   (括号只裹负数)
// ⚠ 判据是 `!= 0`（不是 `> 0`）：发码那一侧把「**省略**」这一档填成 **0**
//   （`kOD_FormatNum` 位 3，§B212 改的），而显式写的 True 在 VB6 里恒为 **-1** ⇒
//   0 = 不要括号、-1/1 = 要括号，两边**分得开**。
//   ⚠⚠ 若某一格把省略档改回 -1，这里立刻回到「省略与显式 True 同码」的老坑：要么显式 True
//   被吃掉（实测 -1,234.57 vs 运行期 (1,234.57)），要么省略档全变括号把正常调用一起改坏。
static BSTR vb6_fmtApplyUseParens(const wchar_t* numBuf, double value, int32_t useParens) {
    if (useParens == 0) return vb6_BSTR_FromStr(numBuf);       // 省略档(0) 与显式 False(0) 都走负号
    if (!(value < 0.0)) return vb6_BSTR_FromStr(numBuf);       // 括号只裹负数
    if (numBuf[0] != L'-') return vb6_BSTR_FromStr(numBuf);
    wchar_t out[80];
    out[0] = L'(';
    wcscpy(out + 1, numBuf + 1);                              // 去掉负号那一枚
    wcscat(out, L")");
    return vb6_BSTR_FromStr(out);
}

// ===== §B213: 剩下那两枚可选位的语义（IncludeLeadingDigits / GroupThousands）=====
// 账 #437：§B212 收了 useParens，但同一条 5 位签名上还剩**两枚**被 `(void)` 丢掉的位。
//   量出来的运行期口径（cscript 同机代理，尺 `.build/klead.vbs` / `kedge.vbs`）三家**完全一致**。

// IncludeLeadingDigits：显式 False 时，整数部分**恰为 "0"** ⇒ 去掉那个 "0"，小数点/百分号留下。
//   运行期实测：FormatNumber(0.5,2,False)=`.50` / (0,2,False)=`.00` / (-0.5,2,False)=`-.50`
//     / (0,0,False)=**空串** / 省略 ⇒ `0.50`（默认带前导 0）/ FormatPercent(0,0,False)=**`%`**
//   ⚠ 判据「整数部分恰好是 0」必须连**百分号**一起看：percent 的 buf 可能是 `"0%"`，下一位是
//     `'%'` 而不是小数点或结束。漏掉这一位 ⇒ 把 `"0%"` 当成「不止一位」而保留 0，交出 `0%`
//     而不是运行期的 `%`。同理 `(0,0,False)` 是**空串**而不是 `0` —— 去掉 0 就是去掉 0。
static void vb6_fmtDropLeadingZero(wchar_t* buf, int32_t incLeading) {
    if (incLeading != 0) return;                 // 省略(-1) 与显式 True(-1) 同义 ⇒ 保留前导 0
    wchar_t* p = (buf[0] == L'-') ? buf + 1 : buf;
    if (p[0] != L'0') return;
    const wchar_t nxt = p[1];
    if (nxt != 0 && nxt != L'.' && nxt != L'%') return;   // 整数部分不止一位 ⇒ 这一位无关
    // ⚠⚠ 这里必须搬 **wcslen(p+1)+1** 个元素（**含那个 '\0'**）。只搬 wcslen(p+1) 个会**少搬
    // 收尾那一格**，被删掉的 '0' 原地留在末尾 ⇒ `"0.50"` 变成 `".500"`、`"0.50%"` 变成
    // `".50%%"`、`"0"` 变成 `"0"`（压根没删）。这三个错都长得像「位数不对」而不是像「多了一格」，
    // 本机第一版就是这么写错的，四个读数同时对不上才定位到。
    memmove(p, p + 1, (wcslen(p + 1) + 1) * sizeof(wchar_t));
}

// GroupThousands：显式 False ⇒ 不插千分位；省略/True ⇒ 每 3 位插 ','（运行期就是 ','）。
//   ⚠ 只数**开头那段数字**：FormatPercent 的 buf 可能是 `"1234.50%"`，若把 `%` 也算成一位，
//     整数段长度会错 1、逗号插错位置（`FormatPercent(12.345,2,True,False,True)` 应为
//     `1,234.50%`，错算会交成 `1,23,4.50%` 之类）。所以这里按 `isdigit` 数，不按 '.' 找。
static void vb6_fmtGroupThousands(wchar_t* buf, int32_t groupDigits) {
    if (groupDigits == 0) return;                // 省略(-1) 与显式 True(-1) 同义 ⇒ 分组
    wchar_t* p = (buf[0] == L'-') ? buf + 1 : buf;
    int intLen = 0;
    while (p[intLen] >= L'0' && p[intLen] <= L'9') intLen++;
    if (intLen <= 3) return;
    wchar_t out[160];                             // 96 加上最多 ~32 个逗号, 留足
    int o = 0;
    for (int i = 0; i < intLen; i++) {
        if (i > 0 && (intLen - i) % 3 == 0) out[o++] = L',';
        out[o++] = p[i];
    }
    wcscpy(out + o, p + intLen);                  // 小数点 + 后缀(可能是 '%') 原样接回
    // ⚠ 负号要**接在最前面**。p 已经跳过了符号位，而 out 是从数字段重新写的 —— 不把符号补回去
    //   就是「插完千分位负号没了」：`FormatNumber(-1234.5678,2)` 交 `1,234.57` 而不是
    //   `-1,234.57`。运行期 `FN neg def=[-1,234.57]` / `FP neg def=[-123,456.78%]` 钉住这一格。
    if (buf[0] == L'-') {
        memmove(out + 1, out, (wcslen(out) + 1) * sizeof(wchar_t));
        out[0] = L'-';
    }
    wcscpy(buf, out);
}

// 注: FormatCurrency 的 `¥`(U+00A5) 前缀**本来就没坏** —— 曾经看到的 "гд" 是终端
// 误解码: 936 下 vb6_ConWriteHandle 正确写出 GBK 双字节 A3A4(即全角"￥"),
// 65001 下写出 UTF-8 C2A5; 码点探针实测 BSTR 内首字符恒为 U+00A5。别当 bug 修。
BSTR vb6_FormatCurrency(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits) {
    if (numDigits < 0) numDigits = 2;
    wchar_t numBuf[96];
    swprintf(numBuf, 96, L"%.*f", (int)numDigits, value);
    // §B213: 这一位改前是 `(void)incLeading; (void)groupDigits` **两个一起丢掉**。
    //   两枚都在括号/符号**之前**定形 —— 千分位必须先插好，否则 `(`/`¥` 会被当成一位数字数错。
    vb6_fmtDropLeadingZero(numBuf, incLeading);      // §B213
    vb6_fmtGroupThousands(numBuf, groupDigits);      // §B213
    // §B212: useParens 为真且值为负 ⇒ 去掉负号、裹进括号（运行期 FC_T 同格）。
    //   ⚠ 货币符号那一枚（这里写死 ¥）**跟着区域设置走**，所以判据那一族只钉「括号/负号」
    //     的形状，不钉符号本身 —— 与 §B142「区域设置那一格照打不钉」同一条纪律。
    BSTR shaped = vb6_fmtApplyUseParens(numBuf, value, useParens);
    // ⚠⚠ **符号在括号里面**：运行期实测是 `(¥1,234.57)`，不是 `¥(1,234.57)`（尺
    //   `.build/kpg.vbs`：`FC parens+group=[(¥1,234.57)]`）。§B212 第一次就是写成
    //   「符号在括号外」的，而那一刀只钉了 `InStr(s,"(")>0` 的**形状** ⇒ 两边都有括号、
    //   形状判据照样绿，**位置错了整整一刀没人看见**。⇒ 这一族要么连符号位置一起钉，要么
    //   就承认自己只钉了形状。§B213 起连位置一起钉。
    const wchar_t* body = shaped;
    int skip = (body[0] == L'(') ? 1 : 0;
    wchar_t result[128];
    if (skip) {
        wcscpy(result, L"(\x00a5");           // '(' + 货币符号
        // ⚠ **不要再补一个 ')'**：body+skip 的**收尾那个 ')' 就在 body 里**（`vb6_fmtApplyUseParens`
        //   交回来的是完整的 `(…��)`），补一个就成 `(¥1,234.57))`。符号只该挪到 '(' **之后**，
        //   右括号原样接着。
        wcscat(result, body + skip);
    } else {
        wcscpy(result, L"\x00a5");
        wcscat(result, body);
    }
    vb6_BSTR_Free(shaped);
    return vb6_BSTR_FromStr(result);
}

BSTR vb6_FormatNumber(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits) {
    if (numDigits < 0) numDigits = 2;
    wchar_t numBuf[96];
    swprintf(numBuf, 96, L"%.*f", (int)numDigits, value);
    // §B213: 这一位改前是 `(void)incLeading` 直接丢掉 ⇒ `FormatNumber(0.5,2,False)` 交 `0.50`，
    //   运行期是 `.50`。千分位那一枚改前是**手写的一遍**（含负号回填那套 memmove），现在也一并
    //   走同一位主人：负号交给 `vb6_fmtDropLeadingZero` / `vb6_fmtGroupThousands` 各自处理。
    vb6_fmtDropLeadingZero(numBuf, incLeading);      // §B213
    vb6_fmtGroupThousands(numBuf, groupDigits);      // §B213
    return vb6_fmtApplyUseParens(numBuf, value, useParens);       // §B212
}

BSTR vb6_FormatPercent(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits) {
    if (numDigits < 0) numDigits = 2;
    wchar_t buf[96];
    /* 百分比: 值 ×100 再加 '%' —— 用 %.*f 直接成型, 不套用二次格式串 */
    swprintf(buf, 96, L"%.*f%%", (int)numDigits, value * 100.0);
    // §B213: 改前 `(void)incLeading` 且 groupDigits 压根没接（运行期 `1,234.50%` / `.50%` 两格
    //   都交不出来）。⚠ buf 里带着 '%'，所以两位主人必须都按「数字跑」而不是「到 '.' 为止」来看。
    vb6_fmtDropLeadingZero(buf, incLeading);          // §B213
    vb6_fmtGroupThousands(buf, groupDigits);          // §B213
    return vb6_fmtApplyUseParens(buf, value, useParens);           // §B212
}


// ============================================================
// P14.2.3: Split/Join 字符串数组函数
// ============================================================

vb6_SafeArray1D* vb6_Split(BSTR expr, BSTR delimiter, int32_t limit, int32_t compare) {
    // <vbeclipse>: compare 此前被 (void) 丢掉 ⇒ Split(s, "x", -1, vbTextCompare)
    // 只按二进制切。分隔符扫描改走 vb6_TextMatchAt (二进制形逐字不变)。
    if (!expr) expr = vb6_BSTR_Empty();
    // Default delimiter is space " " when NULL is passed
    BSTR defaultDelim = NULL;
    if (!delimiter) {
        defaultDelim = vb6_BSTR_FromStr(L" ");
        delimiter = defaultDelim;
    }
    if (limit == 0) limit = -1;
    
    int32_t exprLen = vb6_BSTR_Len(expr);
    int32_t delimLen = vb6_BSTR_Len(delimiter);
    
    // Empty string -> single empty element
    if (exprLen == 0) {
        vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_bstr, 0, 0);
        if (arr) VB6_SA_AT(BSTR, arr, 0) = vb6_BSTR_Empty();
        return arr;
    }
    
    // Count substrings
    int32_t count = 1;
    if (delimLen > 0) {
        for (int32_t i = 0; i <= exprLen - delimLen; ) {
            if (vb6_TextMatchAt(expr, i, delimiter, compare)) {
                count++;
                i += delimLen;
                if (limit > 0 && count >= limit) break;
            } else {
                i++;
            }
        }
    } else {
        // Empty delimiter: split each character
        count = exprLen;
    }
    if (limit > 0 && count > limit) count = limit;
    
    // Create array
    vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_bstr, 0, count - 1);
    if (!arr) return NULL;
    
    // Populate elements
    int32_t idx = 0, start = 0;
    if (delimLen > 0) {
        for (int32_t i = 0; i <= exprLen - delimLen && idx < count - 1; ) {
            if (vb6_TextMatchAt(expr, i, delimiter, compare)) {
                int32_t len = i - start;
                VB6_SA_AT(BSTR, arr, idx) = SysAllocStringLen(expr + start, len);
                idx++;
                start = i + delimLen;
                i += delimLen;
            } else {
                i++;
            }
        }
        // Last element
        VB6_SA_AT(BSTR, arr, idx) = SysAllocStringLen(expr + start, exprLen - start);
    } else {
        // Empty delimiter: each character as element
        for (int32_t i = 0; i < count; i++) {
            VB6_SA_AT(BSTR, arr, i) = SysAllocStringLen(expr + i, 1);
        }
    }
    
    if (defaultDelim) vb6_BSTR_Free(defaultDelim);
    return arr;
}

BSTR vb6_Join(vb6_SafeArray1D* arr, BSTR delimiter) {
    BSTR defaultDelim = NULL;
    if (!delimiter) { defaultDelim = vb6_BSTR_FromStr(L" "); delimiter = defaultDelim; }
    if (!arr || arr->count <= 0) { if (defaultDelim) vb6_BSTR_Free(defaultDelim); return vb6_BSTR_Empty(); }
    
    int32_t delimLen = vb6_BSTR_Len(delimiter);
    
    // Calculate total length
    // <vbeclipse>: 槽位改走 vb6_SA_ReadStrAt (按 elemType 读) —— 旧写法对 Array("a","b")
    // 的 Variant 槽按 BSTR* 取指针, 实测 `Join(Array("abc","xyz"), "|")` 直接段错误。
    int32_t totalLen = 0;
    for (int32_t i = 0; i < arr->count; i++) {
        BSTR elem = vb6_SA_ReadStrAt(arr, i + arr->lBound);
        totalLen += elem ? vb6_BSTR_Len(elem) : 0;
        if (i < arr->count - 1) totalLen += delimLen;
    }
    
    // Build result
    wchar_t* buf = (wchar_t*)malloc((totalLen + 1) * sizeof(wchar_t));
    if (!buf) { if (defaultDelim) vb6_BSTR_Free(defaultDelim); return vb6_BSTR_Empty(); }
    int32_t pos = 0;
    for (int32_t i = 0; i < arr->count; i++) {
        BSTR elem = vb6_SA_ReadStrAt(arr, i + arr->lBound);
        if (elem) {
            int32_t elemLen = vb6_BSTR_Len(elem);
            memcpy(buf + pos, elem, elemLen * sizeof(wchar_t));
            pos += elemLen;
        }
        if (i < arr->count - 1 && delimLen > 0) {
            memcpy(buf + pos, delimiter, delimLen * sizeof(wchar_t));
            pos += delimLen;
        }
    }
    buf[totalLen] = L'\0';
    BSTR result = vb6_BSTR_FromStr(buf);
    free(buf);
    if (defaultDelim) vb6_BSTR_Free(defaultDelim);
    return result;
}

