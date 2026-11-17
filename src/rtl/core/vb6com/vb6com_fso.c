// vb6com_fso.c - 内置 Scripting.FileSystemObject 的 C 实现（IDispatch）
//
// 来历：C30-11 第一刀。在此之前 CreateObject("Scripting.FileSystemObject") 与
// New Scripting.FileSystemObject 都落到真 scrrun.dll，目标机缺 Scripting Runtime 时 429。
// 与 vb6com_dict.c / vb6com_regexp.c 同一条路线：自带实现 + ProgID 拦截。
//
// **本刀只做纯路径面与存在性**（见 ai/040 §4 刀1）。对象面（GetFolder/Drive/集合）与
// TextStream 是刀2/刀3 —— 未落地的名字在 GetIDsOfNames 里答 DISP_E_UNKNOWNNAME。
// 实测（ai/040 §7.5）：当前 vb6_ComCall 拿到 UNKNOWNNAME 的做法是**打一行 stderr + 交 Nothing**，
// 而不是 VB6 的 438；要不要接 438 是全局行为改动（影响所有 COM 对象），不在本刀。
//
// 开关：默认仍走 scrrun（语料 12 份行为一格不变）。内建实现只在
//   C3_FSO_NATIVE=1
// 下抢先拦截；成员面覆盖语料后（刀5）再把默认翻过来，C3_FSO_PREFER_MS=1 留作对照组 ——
// 与 Dictionary 的 C3_DICT_PREFER_MS 同款。
//
// 口径来源：ai/040 §5。手册能定的按手册写并标「册」，真机读数没到的标「待读数」——
// 那些格子等 ProbeFso.bas 的读数回来再改，不在这里猜。

#include "vb6com.h"
#include "vb6com_internal.h"

#include <stdlib.h>
#include <string.h>
#include <wchar.h>

// ============================================================
// DISPID 约定（GetIDsOfNames 返回什么，Invoke 就按什么分发）
// 自有编号：不追 scrrun 的真实 DISPID —— 那要类型库，而这里只走名字入口。
// ============================================================
#define FSO_DISPID_BUILDPATH        1
#define FSO_DISPID_GETFILENAME      2
#define FSO_DISPID_GETBASENAME      3
#define FSO_DISPID_GETEXTENSIONNAME 4
#define FSO_DISPID_GETPARENTFOLDER  5
#define FSO_DISPID_GETABSOLUTEPATH  6
#define FSO_DISPID_FILEEXISTS       7
#define FSO_DISPID_FOLDEREXISTS     8
#define FSO_DISPID_DRIVEEXISTS      9
#define FSO_DISPID_GETTEMPNAME      10

#define FSO_ERR_INVALID_ARG  5     // 无效的过程调用或参数
#define FSO_ERR_NO_MEMBER  438     // 对象不支持此属性或方法

typedef struct vb6_BuiltinFso {
    IDispatchVtbl* lpVtbl;     // 必须首成员 (COM ABI)
    LONG           refCount;
} vb6_BuiltinFso;

static IDispatchVtbl g_fsoVtbl;

// ============================================================
// 错误上报（与 Collection/Dictionary 同族：填 EXCEPINFO.scode，由 vb6_ComCheckError 转 Err）
// ============================================================
static HRESULT fso_fail(EXCEPINFO* pexcep, int32_t errNum) {
    return vb6_ComFailExcep(pexcep, errNum, vb6_StdErrorDesc(errNum), L"FileSystemObject error");
}

// 形参取值（COM 约定：rgvarg 逆序，形参 formalIdx 在 rgvarg[cArgs-1-formalIdx]）
static VARIANT* fso_argRaw(DISPPARAMS* pdp, int32_t formalIdx) {
    if (!pdp || !pdp->rgvarg) return NULL;
    if ((int32_t)pdp->cArgs <= formalIdx) return NULL;
    return &pdp->rgvarg[pdp->cArgs - 1 - formalIdx];
}

// 取一枚字符串实参（返回 malloc 的副本，调用方 free）。
// Empty/Null 当空串 —— FSO 对空参数的答法待读数（ai/040 §5.2），本刀先按"空路径"处理。
static wchar_t* fso_argStr(DISPPARAMS* pdp, int32_t formalIdx) {
    VARIANT* v = fso_argRaw(pdp, formalIdx);
    wchar_t* out = (wchar_t*)calloc(1, sizeof(wchar_t) * 1024);
    if (!out) return NULL;
    if (!v || V_VT(v) == VT_EMPTY || V_VT(v) == VT_NULL) return out;
    if (V_VT(v) == VT_BSTR) {
        const wchar_t* s = V_BSTR(v) ? V_BSTR(v) : L"";
        wcsncpy(out, s, 1023);
        return out;
    }
    {
        VARIANT tmp;
        VariantInit(&tmp);
        if (SUCCEEDED(VariantChangeType(&tmp, v, 0, VT_BSTR)) && V_VT(&tmp) == VT_BSTR)
            wcsncpy(out, V_BSTR(&tmp) ? V_BSTR(&tmp) : L"", 1023);
        VariantClear(&tmp);
    }
    return out;
}

// ============================================================
// 路径小工具（全部在宽字符上就地算，不碰 locale）
// ============================================================

// VB6 的 FSO 接受带引号的路径（"C:\a b"）—— 去掉包裹的双引号。
static void fso_stripQuotes(wchar_t* s) {
    size_t n = s ? wcslen(s) : 0;
    if (n >= 2 && s[0] == L'"' && s[n - 1] == L'"') {
        memmove(s, s + 1, sizeof(wchar_t) * (n - 1));
        s[n - 2] = 0;
    }
}

// 去掉末尾的分隔符（保留 "C:\" 这种盘根本身）
static void fso_stripTrailSep(wchar_t* s) {
    size_t n = wcslen(s);
    while (n > 0 && s[n - 1] == L'\\') {
        if (n == 3 && s[1] == L':') break;      // "C:\" 保留
        if (n == 1) break;                       // "\" 保留
        s[--n] = 0;
    }
}

// 最后一个路径分隔符之后的部分（'\\' 或盘符冒号都当分隔）
static const wchar_t* fso_lastComponent(const wchar_t* p) {
    const wchar_t* last = p;
    for (const wchar_t* q = p; *q; ++q)
        if (*q == L'\\' || *q == L':') last = q + 1;
    return last;
}

// "C:" 结尾的盘符要补分隔符 —— BuildPath("C:", "x") = "C:\x"
static int fso_endsWithDriveColon(const wchar_t* s) {
    size_t n = wcslen(s);
    return n == 2 && s[1] == L':';
}

static void fso_setBstr(VARIANT* pResult, const wchar_t* s) {
    if (!pResult) return;
    V_VT(pResult) = VT_BSTR;
    // 不用 vb6_BSTR_FromStr：它是 vb6rtl_bstr.h 里的 static inline，
    // 在 COM 这一族编译单元里没有可链接的符号（本刀实测 LNK2019）。
    V_BSTR(pResult) = SysAllocString(s ? s : L"");
}

static void fso_setBool(VARIANT* pResult, int b) {
    if (!pResult) return;
    V_VT(pResult) = VT_BOOL;
    V_BOOL(pResult) = b ? VARIANT_TRUE : VARIANT_FALSE;
}

// ============================================================
// 成员实现（刀1：纯路径面 + 存在性）
// ============================================================

// BuildPath(path, name)  —— 册：path 尾部补一个分隔符，name 若已带前导分隔符则不补
static HRESULT fso_doBuildPath(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* path = fso_argStr(pdp, 0);
    wchar_t* name = fso_argStr(pdp, 1);
    if (!path || !name) { free(path); free(name); return fso_fail(pexcep, FSO_ERR_INVALID_ARG); }
    fso_stripQuotes(path);
    fso_stripQuotes(name);

    wchar_t out[2048];
    size_t n;
    wcsncpy(out, path, 2000);
    out[2000] = 0;
    n = wcslen(out);

    if (name[0] == L'\\') {
        // name 自带前导分隔符：去掉 path 的尾部分隔符再拼（不产生双反斜杠）
        fso_stripTrailSep(out);
        n = wcslen(out);
        if (n + wcslen(name) < 2047) wcscat(out, name);
    } else {
        if (n > 0 && out[n - 1] != L'\\') {
            if (n + 1 < 2047) { out[n++] = L'\\'; out[n] = 0; }
        } else if (n == 0) {
            out[0] = L'\\'; out[1] = 0; n = 1;
        }
        if (n + wcslen(name) < 2047) wcscat(out, name);
    }
    (void)fso_endsWithDriveColon;   // "C:" 已由上面"补分隔符"那支覆盖
    fso_setBstr(pResult, out);
    free(path); free(name);
    return S_OK;
}

// GetFileName(path) —— 最后一个分隔符之后；路径以分隔符结尾时答空串
static HRESULT fso_doGetFileName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    const wchar_t* base = fso_lastComponent(p);
    size_t n = wcslen(p);
    if (n > 0 && p[n - 1] == L'\\') fso_setBstr(pResult, L"");   // 待读数 §5.1
    else fso_setBstr(pResult, base);
    free(p);
    return S_OK;
}

// GetBaseName(path) —— GetFileName 去掉最后一个 '.' 及其后
static HRESULT fso_doGetBaseName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    wchar_t name[1024];
    wcsncpy(name, fso_lastComponent(p), 1023);
    name[1023] = 0;
    if (wcslen(p) > 0 && p[wcslen(p) - 1] == L'\\') name[0] = 0;
    wchar_t* dot = wcsrchr(name, L'.');
    if (dot) *dot = 0;
    fso_setBstr(pResult, name);
    free(p);
    return S_OK;
}

// GetExtensionName(path) —— 最后一个 '.' 之后；无点答空串
static HRESULT fso_doGetExtensionName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    const wchar_t* base = fso_lastComponent(p);
    const wchar_t* dot = wcsrchr(base, L'.');
    fso_setBstr(pResult, dot ? dot + 1 : L"");
    free(p);
    return S_OK;
}

// GetParentFolderName(path) —— 去掉最后一段；盘根答 ""（手册：根没有父文件夹）
static HRESULT fso_doGetParentFolderName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    size_t n = wcslen(p);
    while (n > 0 && p[n - 1] == L'\\') p[--n] = 0;      // 末尾分隔符不参与
    wchar_t* cut = wcsrchr(p, L'\\');
    if (!cut) { fso_setBstr(pResult, L""); free(p); return S_OK; }
    if (cut == p) { fso_setBstr(pResult, L""); free(p); return S_OK; }   // "\x"
    // "C:\a" → "C:" 要还原成 "C:\"
    if (cut - p == 2 && p[1] == L':') { cut[1] = 0; fso_setBstr(pResult, p); free(p); return S_OK; }
    *cut = 0;
    fso_setBstr(pResult, p);
    free(p);
    return S_OK;
}

// GetAbsolutePathName(path) —— 相对路径按**当前目录**展开，"."/".." 折叠（ai/040 §5.1）
static HRESULT fso_doGetAbsolutePathName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);

    wchar_t full[2048];
    int rooted = (p[0] && p[1] == L':') || (p[0] == L'\\' && p[1] == L'\\') || (p[0] == L'\\');
    if (rooted) {
        wcsncpy(full, p, 2000); full[2000] = 0;
    } else {
        wchar_t cwd[1024];
        DWORD got = GetCurrentDirectoryW(1024, cwd);
        if (got == 0 || got >= 1024) cwd[0] = 0;
        _snwprintf(full, 2000, L"%s\\%s", cwd, p);
        full[2000] = 0;
    }

    // 折叠：逐段压栈，"." 丢弃，".." 弹一段（不越过根）
    wchar_t out[2048];
    size_t on = 0;
    wchar_t* ctx = NULL;
    wchar_t* seg = wcstok_s(full, L"\\", &ctx);
    if (full[0] == L'\\') { out[on++] = L'\\'; out[on] = 0; }
    while (seg) {
        if (wcscmp(seg, L".") == 0) { /* 丢 */ }
        else if (wcscmp(seg, L"..") == 0) {
            if (on > 0) {
                size_t k = on;
                while (k > 0 && out[k - 1] != L'\\') --k;
                if (k > 0) { on = k - 1; out[on] = 0; }
            }
        } else {
            size_t sl = wcslen(seg);
            // 段之间补分隔符（盘符 "C:" 之后也要补 —— 实测第一版在此漏了：
            // CWD "C:\Users\x" 展开成 "C:Users\x"）。
            if (on > 0 && out[on - 1] != L'\\') {
                if (on + 1 < 2047) { out[on++] = L'\\'; }
            }
            if (on + sl < 2047) { wmemcpy(out + on, seg, sl); on += sl; out[on] = 0; }
        }
        seg = wcstok_s(NULL, L"\\", &ctx);
    }
    if (on == 0) { out[0] = L'\\'; out[1] = 0; }
    fso_setBstr(pResult, out);
    free(p);
    return S_OK;
}

// FileExists(path) —— 目录传进来答 False（待读数 §5.2；本条先按"只认文件"写）
static HRESULT fso_doFileExists(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    DWORD a = p[0] ? GetFileAttributesW(p) : INVALID_FILE_ATTRIBUTES;
    fso_setBool(pResult, a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY));
    free(p);
    return S_OK;
}

// FolderExists(path) —— 文件传进来答 False
static HRESULT fso_doFolderExists(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    DWORD a = p[0] ? GetFileAttributesW(p) : INVALID_FILE_ATTRIBUTES;
    fso_setBool(pResult, a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_DIRECTORY));
    free(p);
    return S_OK;
}

// DriveExists(spec) —— "C" / "C:" / "C:\" 三形等价（手册）
static HRESULT fso_doDriveExists(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    wchar_t root[8] = L"C:\\";
    int ok = 0;
    if (p[0]) {
        root[0] = p[0];
        UINT t = GetDriveTypeW(root);
        ok = (t != DRIVE_NO_ROOT_DIR) && (p[0] >= L'A' && p[0] <= L'Z' || p[0] >= L'a' && p[0] <= L'z');
    }
    fso_setBool(pResult, ok);
    free(p);
    return S_OK;
}

// GetTempName() —— 待读数：真 VB6 的形状（前缀/长度/是否建文件）要到 ProbeFso.bas A10 才钉得死。
// 本刀按"rad + 5 位 + .tmp、**不建文件**"落地；读数回来若不同，改这一处即可。
static HRESULT fso_doGetTempName(VARIANT* pResult) {
    static const wchar_t kChars[] = L"ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    static unsigned seed = 0;
    wchar_t name[32];
    if (seed == 0) seed = (unsigned)GetTickCount() ^ (unsigned)(ULONG_PTR)&name;
    wcscpy(name, L"rad");
    for (int i = 0; i < 5; i++) {
        seed = seed * 1103515245u + 12345u;
        name[3 + i] = kChars[(seed >> 16) % 36];
    }
    name[8] = 0;
    wcscat(name, L".tmp");
    fso_setBstr(pResult, name);
    return S_OK;
}

// ============================================================
// IDispatch
// ============================================================
static HRESULT STDMETHODCALLTYPE fso_QueryInterface(IDispatch* pDisp, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = pDisp;
        ((vb6_BuiltinFso*)pDisp)->refCount++;
        return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE fso_AddRef(IDispatch* pDisp) {
    return (ULONG)++((vb6_BuiltinFso*)pDisp)->refCount;
}

static ULONG STDMETHODCALLTYPE fso_Release(IDispatch* pDisp) {
    vb6_BuiltinFso* f = (vb6_BuiltinFso*)pDisp;
    LONG n = --f->refCount;
    if (n <= 0) free(f);
    return (ULONG)n;
}

static HRESULT STDMETHODCALLTYPE fso_GetTypeInfoCount(IDispatch* pDisp, UINT* pctinfo) {
    (void)pDisp;
    if (pctinfo) *pctinfo = 0;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE fso_GetTypeInfo(IDispatch* pDisp, UINT iTInfo,
                                                  LCID lcid, ITypeInfo** ppTInfo) {
    (void)pDisp; (void)iTInfo; (void)lcid;
    if (ppTInfo) *ppTInfo = NULL;
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE fso_GetIDsOfNames(IDispatch* pDisp, REFIID riid,
                                                   LPOLESTR* rgszNames, UINT cNames,
                                                   LCID lcid, DISPID* rgDispId) {
    (void)pDisp; (void)riid; (void)lcid;
    if (!rgszNames || !rgDispId) return E_POINTER;
    for (UINT i = 0; i < cNames; i++) {
        const wchar_t* n = rgszNames[i];
        if (!n) { rgDispId[i] = DISPID_UNKNOWN; continue; }
        if      (_wcsicmp(n, L"BuildPath")            == 0) rgDispId[i] = FSO_DISPID_BUILDPATH;
        else if (_wcsicmp(n, L"GetFileName")          == 0) rgDispId[i] = FSO_DISPID_GETFILENAME;
        else if (_wcsicmp(n, L"GetBaseName")          == 0) rgDispId[i] = FSO_DISPID_GETBASENAME;
        else if (_wcsicmp(n, L"GetExtensionName")     == 0) rgDispId[i] = FSO_DISPID_GETEXTENSIONNAME;
        else if (_wcsicmp(n, L"GetParentFolderName")  == 0) rgDispId[i] = FSO_DISPID_GETPARENTFOLDER;
        else if (_wcsicmp(n, L"GetAbsolutePathName")  == 0) rgDispId[i] = FSO_DISPID_GETABSOLUTEPATH;
        else if (_wcsicmp(n, L"FileExists")           == 0) rgDispId[i] = FSO_DISPID_FILEEXISTS;
        else if (_wcsicmp(n, L"FolderExists")         == 0) rgDispId[i] = FSO_DISPID_FOLDEREXISTS;
        else if (_wcsicmp(n, L"DriveExists")          == 0) rgDispId[i] = FSO_DISPID_DRIVEEXISTS;
        else if (_wcsicmp(n, L"GetTempName")          == 0) rgDispId[i] = FSO_DISPID_GETTEMPNAME;
        else {
            // 刀2/刀3 的成员（GetFolder/Drives/OpenTextFile/...）也落这里 ⇒ 438，
            // 而不是让调用方在"返回 Empty"里静默错下去。
            rgDispId[i] = DISPID_UNKNOWN;
            return DISP_E_UNKNOWNNAME;
        }
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE fso_Invoke(IDispatch* pDisp, DISPID dispIdMember,
                                            REFIID riid, LCID lcid, WORD wFlags,
                                            DISPPARAMS* pdp, VARIANT* pVarResult,
                                            EXCEPINFO* pexcep, UINT* puArgErr) {
    (void)pDisp; (void)riid; (void)lcid; (void)puArgErr;
    int isCall = (wFlags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET)) != 0;
    if (!isCall) return fso_fail(pexcep, FSO_ERR_NO_MEMBER);

    switch (dispIdMember) {
        case FSO_DISPID_BUILDPATH:       return fso_doBuildPath(pdp, pVarResult, pexcep);
        case FSO_DISPID_GETFILENAME:     return fso_doGetFileName(pdp, pVarResult, pexcep);
        case FSO_DISPID_GETBASENAME:     return fso_doGetBaseName(pdp, pVarResult, pexcep);
        case FSO_DISPID_GETEXTENSIONNAME:return fso_doGetExtensionName(pdp, pVarResult, pexcep);
        case FSO_DISPID_GETPARENTFOLDER: return fso_doGetParentFolderName(pdp, pVarResult, pexcep);
        case FSO_DISPID_GETABSOLUTEPATH: return fso_doGetAbsolutePathName(pdp, pVarResult, pexcep);
        case FSO_DISPID_FILEEXISTS:      return fso_doFileExists(pdp, pVarResult, pexcep);
        case FSO_DISPID_FOLDEREXISTS:    return fso_doFolderExists(pdp, pVarResult, pexcep);
        case FSO_DISPID_DRIVEEXISTS:     return fso_doDriveExists(pdp, pVarResult, pexcep);
        case FSO_DISPID_GETTEMPNAME:     return fso_doGetTempName(pVarResult);
        default:                         return fso_fail(pexcep, FSO_ERR_NO_MEMBER);
    }
}

static IDispatchVtbl g_fsoVtbl = {
    fso_QueryInterface,
    fso_AddRef,
    fso_Release,
    fso_GetTypeInfoCount,
    fso_GetTypeInfo,
    fso_GetIDsOfNames,
    fso_Invoke,
};

// ============================================================
// 对外入口
// ============================================================
// 内建实现是否启用：默认**关**（仍走 scrrun，语料 12 份行为一格不变）。
// 刀5 翻默认时把这一句改成 default-on，并把 C3_FSO_PREFER_MS 留成对照组。
int vb6_Fso_NativeEnabled(void) {
    static int on = -1;
    if (on < 0) on = (GetEnvironmentVariableW(L"C3_FSO_NATIVE", NULL, 0) > 0) ? 1 : 0;
    return on;
}

int32_t vb6_Fso_IsProgramId(const wchar_t* progId) {
    if (!progId || !*progId) return 0;
    if (_wcsicmp(progId, L"Scripting.FileSystemObject") == 0) return 1;
    if (_wcsicmp(progId, L"Scripting.FileSystemObject.1") == 0) return 1;
    return 0;
}

void* vb6_Fso_New(void) {
    vb6_BuiltinFso* f = (vb6_BuiltinFso*)calloc(1, sizeof(vb6_BuiltinFso));
    if (!f) return NULL;
    f->lpVtbl = &g_fsoVtbl;
    f->refCount = 1;
    vb6_ComRememberProgID(f, L"Scripting.FileSystemObject");
    return (void*)f;
}

int32_t vb6_Fso_IsClassName(const wchar_t* className) {
    if (!className || !*className) return 0;
    if (_wcsicmp(className, L"FileSystemObject") == 0) return 1;
    if (_wcsicmp(className, L"Scripting.FileSystemObject") == 0) return 1;
    return 0;
}

int vb6_Fso_IsInstance(void* obj) {
    if (!obj) return 0;
    return *(void***)obj == (void**)&g_fsoVtbl;
}
