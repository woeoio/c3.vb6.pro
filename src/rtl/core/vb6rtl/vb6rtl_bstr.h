#pragma once
// vb6rtl_bstr.h - BSTR（VB6 字符串类型）与字符串内存管理
// 由 vb6rtl.h 伞头 include；生成代码不要直接 include 本文件
#include "vb6rtl_base.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// BSTR - VB6 字符串类型 (OLE BSTR)
// ============================================================

// BSTR 本质是 wchar_t* (指向OLE BSTR长度前缀之后的字符数据)
// Windows: 使用OLE API (SysAllocString/SysFreeString/SysStringLen)
// 非Windows: 自定义malloc实现
typedef wchar_t* BSTR;

// 空BSTR
static inline BSTR vb6_BSTR_Empty(void) {
#ifdef _WIN32
    return SysAllocStringLen(L"", 0);
#else
    return NULL;
#endif
}

// 定长字符串初始化: 返回 len 个 **NUL 字符** (Chr$(0) / vbNullChar) 的 BSTR
// (VB6 `Dim a As String * N`)。
//
// 账 #118 校准 (2026-10-07): VB6 的定长串有**两套**填充, 别混成一套 ——
//   · **未赋值时用 0 填充**。依据两处独立: ① 仓库自带手册 Dim/Private/Static 语句
//     ("定长的字符串则用 0 填充"); ② MSDN 同句英文本 "a fixed-length string is
//     filled with zeros"; ③ 真机口径可查 (用 CopyMemory 取 StrPtr 后 20 字节,
//     未赋值前全是 Chr$(0))。RTrim$/Trim$ **不裁 NUL**, 所以未赋值的定长串拿去做
//     宽字符串 (如 `FindWindowExW` 的类名) 只会得到空串, 这是 VB6 行为不是缺陷
//     (ai/029 里 `FindWindowExW` + `StrPtr("BUTTON")` 那枚探针当年读空就是这么来的)。
//   · **赋值时右侧补空格、超长截右** (见 vb6_LSet / 发码侧 cgen_assign_value_sem.inc)。
// 此前这里补的是空格, 依据是 ai/开发历程/53-M13-FIX4 那句"初始内容为10个空格" ——
// 那句话把"赋值补空格"错记成了"初始化", 本格纠正。
static inline BSTR vb6_BSTR_FixedSTR(int32_t len) {
    if (len <= 0) return vb6_BSTR_Empty();
#ifdef _WIN32
    // SysAllocStringLen(NULL, len) 本身就是零填充 (连 buf[len] 的终止符一起), 显式
    // 再写一遍只为把"0 填充"这条口径落在代码上, 不靠 API 的隐含行为。
    BSTR bstr = SysAllocStringLen(NULL, len);
    if (bstr) { for (int32_t i = 0; i < len; i++) bstr[i] = L'\0'; }
    return bstr;
#else
    return NULL;
#endif
}

// 从宽字符串创建BSTR
// Windows: 使用SysAllocString (分配副本, 可安全释放)
// 非Windows: malloc分配自定义长度前缀
static inline BSTR vb6_BSTR_FromStr(const wchar_t* s) {
    if (!s) return NULL;
#ifdef _WIN32
    return SysAllocString(s);
#else
    size_t len = wcslen(s);
    uint32_t* p = (uint32_t*)malloc(sizeof(uint32_t) + (len + 1) * sizeof(wchar_t));
    if (!p) return NULL;
    *p = (uint32_t)len;
    BSTR bstr = (BSTR)(p + 1);
    memcpy(bstr, s, (len + 1) * sizeof(wchar_t));
    return bstr;
#endif
}

// 从**定长**宽串创建 BSTR —— 走 SysAllocStringLen, 不按 NUL 截断。
// 专给含**内嵌 NUL** 的串常量用 (VB6 `vbNullChar` / Chr$(0) 拼出的串):
// 普通 vb6_BSTR_FromStr 走 wcslen, 会把内嵌 NUL 当串尾 ⇒ 长度丢失,
// 最直接的受害者是 `s & vbNullChar` 这类「末尾补 NUL」的写法 (ListView 的
// LVN_GETDISPINFO 回调正是靠它给 pszText 落终止符) —— 少了终止符, 控件读越界,
// 屏幕上是「正确文字 + 后接一串乱码」。与 vb6_BSTR_FromBSTR (Fix 161b-decl-out) 同口径。
static inline BSTR vb6_BSTR_FromStrN(const wchar_t* s, int32_t n) {
    if (!s || n <= 0) return vb6_BSTR_Empty();
#ifdef _WIN32
    BSTR r = SysAllocStringLen(NULL, (UINT)n);
    if (!r) return vb6_BSTR_Empty();
    memcpy(r, s, (size_t)n * sizeof(WCHAR));
    r[n] = L'\0';
    return r;
#else
    uint32_t* p = (uint32_t*)malloc(sizeof(uint32_t) + (size_t)(n + 1) * sizeof(wchar_t));
    if (!p) return NULL;
    *p = (uint32_t)n;
    BSTR bstr = (BSTR)(p + 1);
    memcpy(bstr, s, (size_t)n * sizeof(wchar_t));
    bstr[n] = 0;
    return bstr;
#endif
}

// 复制BSTR (创建独立副本, 调用方负责释放)
static inline BSTR vb6_BSTR_FromBSTR(BSTR bstr) {
    if (!bstr) return vb6_BSTR_Empty();
#ifdef _WIN32
    // Fix 161b-decl-out: 必须按 SysStringLen 全量复制, 不能用 SysAllocString ——
    // 后者按 NUL 定长, 会把含内嵌 NUL 的字符串 (VB6 里 `String$(260, 0)` 这种
    // Declare 出参缓冲的标准写法, Len() 应为 260) 截成 0 长度。此前
    // `vb6_BSTR_Assign(&buf, vb6_String(260, 0))` 正是经此把 260 丢成 0,
    // 后续 A 版 API 按 nSize=260 往只有 1 字节的缓冲写 ⇒ 堆越界 (0xC0000374)。
    {
        UINT len = SysStringLen(bstr);
        BSTR r = SysAllocStringLen(NULL, len);
        if (!r) return vb6_BSTR_Empty();
        if (len) memcpy(r, bstr, (size_t)len * sizeof(WCHAR));
        r[len] = L'\0';
        return r;
    }
#else
    return vb6_BSTR_FromStr(bstr);
#endif
}

// 释放BSTR
static inline void vb6_BSTR_Free(BSTR bstr) {
    if (bstr) {
#ifdef _WIN32
        SysFreeString(bstr);
#else
        uint32_t* p = ((uint32_t*)bstr) - 1;
        free(p);
#endif
    }
}

// M22: BSTR->ANSI conversion (for Declare A-version API ByVal String params)
// Converts BSTR(UTF-16) to GBK multibyte string. Caller must free with vb6_FreeANSI.
static inline char* vb6_BSTR_ToANSI(BSTR bstr) {
    if (!bstr) {
        char* r = (char*)malloc(1);
        if (r) r[0] = '\0';
        return r;
    }
#ifdef _WIN32
    int len = WideCharToMultiByte(CP_ACP, 0, bstr, -1, NULL, 0, NULL, NULL);
    if (len <= 0) { char* r = (char*)malloc(1); if (r) r[0] = '\0'; return r; }
    char* buf = (char*)malloc(len);
    if (!buf) return NULL;
    WideCharToMultiByte(CP_ACP, 0, bstr, -1, buf, len, NULL, NULL);
    return buf;
#else
    // Non-Windows: simplified to UTF-8
    size_t len = wcstombs(NULL, bstr, 0);
    if (len == (size_t)-1) { char* r = (char*)malloc(1); if (r) r[0] = '\0'; return r; }
    char* buf = (char*)malloc(len + 1);
    if (!buf) return NULL;
    wcstombs(buf, bstr, len + 1);
    return buf;
#endif
}

// M22: Free ANSI string buffer
static inline void vb6_FreeANSI(char* ansi) {
    free(ansi);
}

// Fix 161b-decl-out: Declare A 版 ByVal String 当**可写缓冲**时的 ANSI 转换。
// 与 vb6_BSTR_ToANSI 的区别: 后者按 NUL 结尾算长度 (-1), 而 VB6 的
// `buf = String$(256, 0)` 内容全是 NUL ⇒ 按 -1 只转出 1 字节, 远小于缓冲容量,
// 被调用方一写就越界 (t1 实测: 只回读到 NUL ⇒ 空串)。这里按 **BSTR 的字符数**
// (SysStringLen, 含内嵌 NUL) 全量转换, 并额外多留 1 字节给终止符, 缓冲容量足够。
static inline char* vb6_BSTR_ToANSI_cap(BSTR bstr) {
    if (!bstr) { char* r = (char*)malloc(1); if (r) r[0] = '\0'; return r; }
#ifdef _WIN32
    UINT clen = SysStringLen(bstr);
    if (clen == 0) { char* r = (char*)malloc(1); if (r) r[0] = '\0'; return r; }
    int len = WideCharToMultiByte(CP_ACP, 0, bstr, (int)clen, NULL, 0, NULL, NULL);
    if (len <= 0) { char* r = (char*)malloc(1); if (r) r[0] = '\0'; return r; }
    char* buf = (char*)malloc((size_t)len + 1);
    if (!buf) return NULL;
    WideCharToMultiByte(CP_ACP, 0, bstr, (int)clen, buf, len, NULL, NULL);
    buf[len] = '\0';
    return buf;
#else
    size_t clen = wcslen(bstr);
    if (clen == 0) { char* r = (char*)malloc(1); if (r) r[0] = '\0'; return r; }
    size_t len = wcstombs(NULL, bstr, clen);
    if (len == (size_t)-1) { char* r = (char*)malloc(1); if (r) r[0] = '\0'; return r; }
    char* buf = (char*)malloc(len + 1);
    if (!buf) return NULL;
    wcstombs(buf, bstr, clen);
    buf[len] = '\0';
    return buf;
#endif
}

// Fix 161b-decl-out: Declare A 版 ByVal String **出参回写**。
// VB6 里 `Declare Function GetUserName Lib ... (ByVal lpBuffer As String, ...)`
// 的 ByVal String 实参是"可写缓冲"——被调用方填字节后, VB6 按 NUL 截断并把内容
// 回写成新的 String。C3 侧实参是 vb6_BSTR_ToANSI() 的临时副本, 故调用后须把
// 副本按 NUL 截断转回 BSTR (cap 是缓冲容量上限, <=0 表示按 NUL 定长)。
static inline BSTR vb6_BSTR_FromANSIBuf(const char* buf, int32_t cap) {
    if (!buf) return vb6_BSTR_Empty();
#ifdef _WIN32
    // 只取 NUL 前的有效字节; cap 只是上限, 实际以 NUL 为准 (API 写成有终止)
    int n = 0;
    if (cap > 0) {
        while (n < cap && buf[n] != '\0') n++;
    } else {
        while (buf[n] != '\0') n++;
    }
    if (n <= 0) return vb6_BSTR_Empty();
    int wlen = MultiByteToWideChar(CP_ACP, 0, buf, n, NULL, 0);
    if (wlen <= 0) return vb6_BSTR_Empty();
    BSTR r = SysAllocStringLen(NULL, (UINT)wlen);
    if (!r) return vb6_BSTR_Empty();
    MultiByteToWideChar(CP_ACP, 0, buf, n, r, wlen);
    r[wlen] = L'\0';
    return r;
#else
    int n = 0;
    if (cap > 0) { while (n < cap && buf[n] != '\0') n++; } else { while (buf[n] != '\0') n++; }
    if (n <= 0) return vb6_BSTR_Empty();
    size_t wlen = mbstowcs(NULL, buf, (size_t)n);
    if (wlen == (size_t)-1) return vb6_BSTR_Empty();
    wchar_t* tmp = (wchar_t*)malloc((wlen + 1) * sizeof(wchar_t));
    if (!tmp) return vb6_BSTR_Empty();
    mbstowcs(tmp, buf, (size_t)n);
    tmp[wlen] = L'\0';
    BSTR r = vb6_BSTR_FromStr(tmp);
    free(tmp);
    return r;
#endif
}

// BSTR赋值 (释放旧值, 复制新值, 防止悬垂指针和双重释放)
static inline void vb6_BSTR_Assign(BSTR* target, BSTR source) {
    if (target) {
        BSTR copy = source ? vb6_BSTR_FromBSTR(source) : vb6_BSTR_Empty();
        vb6_BSTR_Free(*target);
        *target = copy;
    }
}

// BSTR连接
BSTR vb6_BSTR_Concat(BSTR a, BSTR b);

// BSTR连接并释放第一个参数 (用于链式Concat中间节点, 防止临时BSTR泄漏)
// a & b & c → vb6_BSTR_ConcatFree(vb6_BSTR_Concat(a, b), c) → 内层结果被ConcatFree释放
BSTR vb6_BSTR_ConcatFree(BSTR a, BSTR b);

// BSTR转移所有权赋值 (free旧值+直接持有source, 不做deep copy)
// 用于Concat等返回新BSTR的赋值场景, source是返回的新BSTR不存在共享引用
static inline void vb6_BSTR_AssignMove(BSTR* target, BSTR source) {
    if (target) {
        vb6_BSTR_Free(*target);
        *target = source;
    }
}

// BSTR长度 (返回字符数, 非字节数)
static inline int32_t vb6_BSTR_Len(BSTR bstr) {
    if (!bstr) return 0;
#ifdef _WIN32
    return (int32_t)SysStringLen(bstr);
#else
    uint32_t* p = ((uint32_t*)bstr) - 1;
    return (int32_t)*p;
#endif
}

#ifdef __cplusplus
}
#endif
