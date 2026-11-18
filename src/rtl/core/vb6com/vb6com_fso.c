// vb6com_fso.c - 内置 Scripting.FileSystemObject 的 C 实现（IDispatch + IEnumVARIANT）
//
// 来历：C30-11。在此之前 CreateObject("Scripting.FileSystemObject") 与
// New Scripting.FileSystemObject 都落到真 scrrun.dll，目标机缺 Scripting Runtime 时 429。
// 与 vb6com_dict.c / vb6com_regexp.c 同一条路线：自带实现 + ProgID 拦截。
//
// 分刀（ai/040 §4）：
//   刀1 纯路径面 + 存在性        —— 已落（BuildPath/GetFileName/.../GetTempName）
//   刀2 对象面                   —— 本文件：GetDrive/GetFolder/GetFile/GetSpecialFolder +
//                                   Drives/Drive/Folder/File/Files/Folders 集合
//   刀3 TextStream               —— 本文件：CreateTextFile/OpenTextFile + TextStream 全成员
//   刀4 写面                     —— 本文件顺带：CreateFolder/DeleteFile/DeleteFolder/
//                                   CopyFile/MoveFile（CopyFolder/MoveFolder **未做**，
//                                   别照这行注释当已实现 —— 名字表里没有它们）
//
// 开关：**默认开启内建实现**（2026-11-18 用户要求翻面），scrrun 只作对照组：
//   C3_FSO_PREFER_MS=1  → 回到真 scrrun（A/B 对照、以及"我这边语义可疑"时的退路）
//
// 口径来源：ai/040 §5。手册能定的按手册写并标「册」，真机读数没到的标「待读数」——
// 那些格子等 ProbeFso.bas 的读数回来再改，不在这里猜。

#include "vb6com.h"
#include "vb6com_internal.h"

#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <stdio.h>
#include <shellapi.h>   // SHGetFileInfoW（Type 属性；shell32.lib 三个链接分支本就都带）

// ============================================================
// 对象种类
// ============================================================
enum {
    FSO_K_FSO = 1,
    FSO_K_DRIVE,
    FSO_K_FOLDER,
    FSO_K_FILE,
    FSO_K_DRIVES,      // 集合
    FSO_K_FOLDERS,     // 集合
    FSO_K_FILES,       // 集合
    FSO_K_TEXTSTREAM,
};

// ============================================================
// DISPID 表（按对象种类各自编号；GetIDsOfNames 先问 kind 再查名字）
// ============================================================
// FSO
#define FD_BUILDPATH        1
#define FD_GETFILENAME      2
#define FD_GETBASENAME      3
#define FD_GETEXTENSIONNAME 4
#define FD_GETPARENTFOLDER  5
#define FD_GETABSOLUTEPATH  6
#define FD_FILEEXISTS       7
#define FD_FOLDEREXISTS     8
#define FD_DRIVEEXISTS      9
#define FD_GETTEMPNAME      10
#define FD_GETFOLDER        11
#define FD_GETFILE          12
#define FD_GETDRIVE         13
#define FD_GETSPECIALFOLDER 14
#define FD_DRIVES           15
#define FD_CREATEFOLDER     16
#define FD_DELETEFILE       17
#define FD_DELETEFOLDER     18
#define FD_COPYFILE         19
#define FD_MOVEFILE         20
#define FD_CREATETEXTFILE   21
#define FD_OPENTEXTFILE     22
#define FD_GETDRIVENAME     23   // C30-11 第六刀：盘符抽取（读数 "C:\a\b"->"C:"、UNC->"\\srv\share"、相对->""）
// 集合
#define FD_COUNT            30
#define FD_ITEM             31
#define FD_NEWENUM          (-4)
// Folder / File 公共属性
#define FD_NAME             40
#define FD_PATH             41
#define FD_SHORTPATH        42
#define FD_TYPE             43
#define FD_SIZE             44
#define FD_DATECREATED      45
#define FD_DATELASTMODIFIED 46
#define FD_DATELASTACCESSED 47
#define FD_ATTRIBUTES       48
#define FD_PARENTFOLDER     49
#define FD_DRIVE            50
#define FD_SHORTNAME        51   // C30-11 第六刀：8.3 **名字**（不是路径），读数 Folder=rd1 / File=orig.txt
// Folder 专有
#define FD_ISROOTFOLDER     60
#define FD_FILES            61
#define FD_SUBFOLDERS       62
#define FD_CREATETEXTFILE2  63
// File 专有
#define FD_OPENASTEXTSTREAM 70
// 集合计算属性
#define FD_DELETE           80
#define FD_COPY             81
#define FD_MOVE             82
// Drive
#define FD_DRIVELETTER      90
#define FD_ROOTFOLDER       91
#define FD_DRIVETYPE        92
#define FD_FILESYSTEM       93
#define FD_ISREADY          94
#define FD_VOLUMENAME       95
#define FD_TOTALSIZE        96
#define FD_FREESPACE        97
#define FD_AVAILABLESPACE   98
#define FD_SERIALNUMBER     99
#define FD_SHARENAME       100   // C30-11 第六刀：本地盘读数 = ""（空串），UNC 盘给 share 名
// TextStream
#define FD_READ             110
#define FD_READLINE         111
#define FD_READALL          112
#define FD_SKIP             113
#define FD_SKIPLINE         114
#define FD_WRITE            115
#define FD_WRITELINE        116
#define FD_WRITEBLANKLINES  117
#define FD_CLOSE            118
#define FD_ATENDOFSTREAM    119
#define FD_ATENDOFLINE      120
#define FD_LINE             121
#define FD_COLUMN           122

// VB6 错误号（描述文本复用 vb6com.c 的 vb6_StdErrorDesc 表）
#define FSO_ERR_INVALID_ARG  5
#define FSO_ERR_PERMISSION  70
#define FSO_ERR_PATH_NOT_FOUND 76
#define FSO_ERR_FILE_NOT_FOUND 53
#define FSO_ERR_ALREADY_EXISTS 58
#define FSO_ERR_BAD_FILE_NAME 52
#define FSO_ERR_NO_MEMBER   438

// ============================================================
// 对象
// ============================================================
typedef struct vb6_FsoObj {
    IDispatchVtbl* lpVtbl;      // 必须首成员 (COM ABI)
    LONG   refCount;
    int    kind;
    wchar_t path[1024];         // Folder/File 的完整路径；Drive 的根（"C:\"）
    wchar_t name[512];          // 集合里的名字（Files/Folders 的 Item 用）
    // TextStream
    int    reading, writing, unicode;
    wchar_t* buf;               // 全文缓冲（读：解码后；写：待落盘）
    int    len, cap, pos;
    int    line, lineStart;
} vb6_FsoObj;

static IDispatchVtbl g_fsoVtbl;
static IEnumVARIANTVtbl g_fsoEnumVtbl;

// ============================================================
// 错误上报（与 Collection/Dictionary 同族）
// ============================================================
static HRESULT fso_fail(EXCEPINFO* pexcep, int32_t errNum) {
    return vb6_ComFailExcep(pexcep, errNum, vb6_StdErrorDesc(errNum), L"FileSystemObject error");
}

static VARIANT* fso_argRaw(DISPPARAMS* pdp, int32_t formalIdx) {
    if (!pdp || !pdp->rgvarg) return NULL;
    if ((int32_t)pdp->cArgs <= formalIdx) return NULL;
    return &pdp->rgvarg[pdp->cArgs - 1 - formalIdx];
}

// 取一枚字符串实参（malloc 的副本，调用方 free）。Empty/Null 当空串。
static wchar_t* fso_argStr(DISPPARAMS* pdp, int32_t formalIdx) {
    VARIANT* v = fso_argRaw(pdp, formalIdx);
    wchar_t* out = (wchar_t*)calloc(1, sizeof(wchar_t) * 1024);
    if (!out) return NULL;
    if (!v || V_VT(v) == VT_EMPTY || V_VT(v) == VT_NULL) return out;
    if (V_VT(v) == VT_BSTR) { wcsncpy(out, V_BSTR(v) ? V_BSTR(v) : L"", 1023); return out; }
    {
        VARIANT tmp; VariantInit(&tmp);
        // R1（账 #305 §B224）: 目的档是 BSTR 时**只许问 vb6_ComChangeType** ——
        // oleaut32 的 VariantChangeType 把 True/False 交成 "-1"/"0"，与 VB6 的
        // "True"/"False" 不同口径。静态门 check_variant_to_string_authority 就盯这一格。
        if (SUCCEEDED(vb6_ComChangeType(&tmp, v, VT_BSTR)) && V_VT(&tmp) == VT_BSTR)
            wcsncpy(out, V_BSTR(&tmp) ? V_BSTR(&tmp) : L"", 1023);
        VariantClear(&tmp);
    }
    return out;
}

// 取一枚整数实参（缺省用 def）
static long fso_argLong(DISPPARAMS* pdp, int32_t formalIdx, long def) {
    VARIANT* v = fso_argRaw(pdp, formalIdx);
    if (!v) return def;
    VARIANT tmp; VariantInit(&tmp);
    if (SUCCEEDED(VariantChangeType(&tmp, v, 0, VT_I4))) { long r = V_I4(&tmp); VariantClear(&tmp); return r; }
    VariantClear(&tmp);
    return def;
}

// ============================================================
// 路径小工具
// ============================================================
static void fso_stripQuotes(wchar_t* s) {
    size_t n = s ? wcslen(s) : 0;
    if (n >= 2 && s[0] == L'"' && s[n - 1] == L'"') {
        memmove(s, s + 1, sizeof(wchar_t) * (n - 1));
        s[n - 2] = 0;
    }
}

static void fso_stripTrailSep(wchar_t* s) {
    size_t n = wcslen(s);
    while (n > 0 && s[n - 1] == L'\\') {
        if (n == 3 && s[1] == L':') break;
        if (n == 1) break;
        s[--n] = 0;
    }
}

static const wchar_t* fso_lastComponent(const wchar_t* p) {
    const wchar_t* last = p;
    for (const wchar_t* q = p; *q; ++q)
        if (*q == L'\\' || *q == L':') last = q + 1;
    return last;
}

static void fso_setBstr(VARIANT* pResult, const wchar_t* s) {
    if (!pResult) return;
    V_VT(pResult) = VT_BSTR;
    // 不用 vb6_BSTR_FromStr：它是 vb6rtl_bstr.h 里的 static inline，COM 这一族编译单元里没有
    // 可链接的符号（刀1 实测 LNK2019）。
    V_BSTR(pResult) = SysAllocString(s ? s : L"");
}

static void fso_setBool(VARIANT* pResult, int b) {
    if (!pResult) return;
    V_VT(pResult) = VT_BOOL;
    V_BOOL(pResult) = b ? VARIANT_TRUE : VARIANT_FALSE;
}

static void fso_setLong(VARIANT* pResult, long v) {
    if (!pResult) return;
    V_VT(pResult) = VT_I4;
    V_I4(pResult) = v;
}

static void fso_setDate(VARIANT* pResult, double d) {
    if (!pResult) return;
    V_VT(pResult) = VT_DATE;
    V_DATE(pResult) = d;
}

// 相对路径按当前目录展开 + "."/".." 折叠
static void fso_absolutePath(const wchar_t* in, wchar_t* out, size_t cap) {
    wchar_t full[2048];
    int rooted = (in[0] && in[1] == L':') || in[0] == L'\\';
    if (rooted) { wcsncpy(full, in, 2000); full[2000] = 0; }
    else {
        wchar_t cwd[1024];
        DWORD got = GetCurrentDirectoryW(1024, cwd);
        if (got == 0 || got >= 1024) cwd[0] = 0;
        _snwprintf(full, 2000, L"%s\\%s", cwd, in);
        full[2000] = 0;
    }
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
            // 段之间补分隔符（盘符 "C:" 之后也要补 —— 第一版在此漏了，C:\Users 会展开成 C:Users）
            if (on > 0 && out[on - 1] != L'\\' && on + 1 < cap) out[on++] = L'\\';
            if (on + sl < cap) { wmemcpy(out + on, seg, sl); on += sl; out[on] = 0; }
        }
        seg = wcstok_s(NULL, L"\\", &ctx);
    }
    if (on == 0) { out[0] = L'\\'; out[1] = 0; }
}

// 拼目录 + 名字
static void fso_join(wchar_t* out, size_t cap, const wchar_t* dir, const wchar_t* name) {
    wcsncpy(out, dir, cap - 2);
    out[cap - 2] = 0;
    fso_stripTrailSep(out);
    size_t n = wcslen(out);
    if (name && *name) {
        if (n + 1 < cap) { out[n++] = L'\\'; out[n] = 0; }
        wcsncat(out, name, cap - n - 1);
    }
    if (n == 2 && out[1] == L':') { wcscat(out, L"\\"); }   // "C:" → "C:\"
}

// ============================================================
// 时间 / 类型 / 属性
// ============================================================
// FILETIME → VB6 Date（本地时间；待读数 §5.3：秒以下是否截断）
static double fso_timeToDate(const FILETIME* ft) {
    FILETIME local;
    ULARGE_INTEGER u;
    if (!FileTimeToLocalFileTime(ft, &local)) local = *ft;
    u.LowPart = local.dwLowDateTime; u.HighPart = local.dwHighDateTime;
    return (double)u.QuadPart / 864000000000.0 - 25569.0;
}

// 扩展名 → 注册表里的类型描述（"Text Document"）；取不到回 "<EXT> File"
// 走 RegOpenKeyExW/RegQueryValueExW（advapi32 三个链接分支本就都带；RegGetValueW 要
// _WIN32_WINNT>=0x0600，本工程没统一抬那一档，不在这里新开一条 DLL 依赖）。
// 注意这是**兜底**：主路径是 fso_typeName 的 shell 类型名（本地化）。
static void fso_fileTypeReg(const wchar_t* path, wchar_t* out, size_t cap) {
    const wchar_t* base = fso_lastComponent(path);
    const wchar_t* dot = wcsrchr(base, L'.');
    if (!dot || !dot[1]) { wcsncpy(out, L"File", cap - 1); out[cap - 1] = 0; return; }
    HKEY k = NULL;
    wchar_t val[512] = L"";
    DWORD sz = sizeof(val) - sizeof(wchar_t), typ = 0;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, dot, 0, KEY_QUERY_VALUE, &k) == ERROR_SUCCESS) {
        if (RegQueryValueExW(k, NULL, NULL, &typ, (LPBYTE)val, &sz) == ERROR_SUCCESS && val[0]) {
            RegCloseKey(k);
            // 第一级是 ProgID（".txt" → "txtfile"），scrrun 给的是**描述**，
            // 所以还要再跳一级：HKCR\txtfile 默认值 = "Text Document"。
            wchar_t desc[512] = L"";
            DWORD sz2 = sizeof(desc) - sizeof(wchar_t), typ2 = 0;
            HKEY k2 = NULL;
            if (RegOpenKeyExW(HKEY_CLASSES_ROOT, val, 0, KEY_QUERY_VALUE, &k2) == ERROR_SUCCESS) {
                RegQueryValueExW(k2, NULL, NULL, &typ2, (LPBYTE)desc, &sz2);
                RegCloseKey(k2);
            }
            wcsncpy(out, desc[0] ? desc : val, cap - 1);
            out[cap - 1] = 0;
            return;
        }
        RegCloseKey(k);
    }
    _snwprintf(out, cap - 1, L"%s File", dot + 1);
    out[cap - 1] = 0;
}

// Type 属性的主口径：**shell 的类型名**（本地化）。
// 实测（ai/040 §7.6）：scrrun 给文件夹答 "文件夹"、给 .txt 答 "文本文档"，
// 不是 HKCR\Directory 的英文默认值 "File Folder"、也不是 .txt 的 ProgID "txtfile"。
// 这就是 SHGetFileInfoW(SHGFI_TYPENAME) 的答案，`SHGFI_USEFILEATTRIBUTES` 让它
// 不摸盘（按扩展名/属性查映射），与 scrrun 对不存在路径也能答 Type 的行为一致。
static void fso_typeName(const wchar_t* path, int isDir, wchar_t* out, size_t cap) {
    SHFILEINFOW sfi;
    ZeroMemory(&sfi, sizeof(sfi));
    DWORD attr = isDir ? FILE_ATTRIBUTE_DIRECTORY : FILE_ATTRIBUTE_NORMAL;
    if (SHGetFileInfoW(path, attr, &sfi, sizeof(sfi), SHGFI_TYPENAME | SHGFI_USEFILEATTRIBUTES)
        && sfi.szTypeName[0]) {
        wcsncpy(out, sfi.szTypeName, cap - 1);
        out[cap - 1] = 0;
        return;
    }
    fso_fileTypeReg(path, out, cap);
}

// 文件夹的 Size 是**递归内容总和**（实测 scrrun：本目录 4846545，与 0 差得远）。
// VB6 手册写“Folder.Size 对文件夹无意义”，真机读数反驳了手册这一句 —— 按读数走。
static double fso_dirSize(const wchar_t* dir) {
    wchar_t pattern[1200];
    _snwprintf(pattern, 1100, L"%s\\*", dir);
    pattern[1100] = 0;
    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(pattern, &fd);
    if (h == INVALID_HANDLE_VALUE) return 0.0;
    double sum = 0.0;
    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        ULARGE_INTEGER u;
        u.LowPart = fd.nFileSizeLow; u.HighPart = fd.nFileSizeHigh;
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            wchar_t sub[1200];
            fso_join(sub, 1200, dir, fd.cFileName);
            sum += fso_dirSize(sub);
        } else {
            sum += (double)u.QuadPart;
        }
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return sum;
}

static long fso_attrs(DWORD a) {
    long r = 0;
    if (a & FILE_ATTRIBUTE_READONLY) r |= 1;
    if (a & FILE_ATTRIBUTE_HIDDEN)   r |= 2;
    if (a & FILE_ATTRIBUTE_SYSTEM)   r |= 4;
    if (a & FILE_ATTRIBUTE_DIRECTORY) r |= 16;
    if (a & FILE_ATTRIBUTE_ARCHIVE)  r |= 32;
    return r;
}

// ============================================================
// 新建对象
// ============================================================
static vb6_FsoObj* fso_objNew(int kind) {
    vb6_FsoObj* o = (vb6_FsoObj*)calloc(1, sizeof(vb6_FsoObj));
    if (!o) return NULL;
    o->lpVtbl = &g_fsoVtbl;
    o->refCount = 1;
    o->kind = kind;
    // TypeName 走 vb6_CreateObject 那一刻记下的类名（vb6_ComCreatedTypeNameBStr）。
    // 刀2 起 Folder/File/Drive/TextStream 也是「对象」，各记自己的类名 ——
    // VB6 的 TypeName(folder) 读作 "Folder"，不是 "FileSystemObject"（册）。
    switch (kind) {
        case FSO_K_FSO:        vb6_ComRememberProgID(o, L"Scripting.FileSystemObject"); break;
        case FSO_K_DRIVE:      vb6_ComRememberProgID(o, L"Scripting.Drive"); break;
        case FSO_K_FOLDER:     vb6_ComRememberProgID(o, L"Scripting.Folder"); break;
        case FSO_K_FILE:       vb6_ComRememberProgID(o, L"Scripting.File"); break;
        case FSO_K_DRIVES:     vb6_ComRememberProgID(o, L"Scripting.Drives"); break;
        case FSO_K_FOLDERS:    vb6_ComRememberProgID(o, L"Scripting.Folders"); break;
        case FSO_K_FILES:      vb6_ComRememberProgID(o, L"Scripting.Files"); break;
        case FSO_K_TEXTSTREAM: vb6_ComRememberProgID(o, L"Scripting.TextStream"); break;
        default: break;
    }
    return o;
}

static void fso_setObj(VARIANT* pResult, vb6_FsoObj* o) {
    if (!pResult || !o) return;
    o->refCount++;
    V_VT(pResult) = VT_DISPATCH;
    V_DISPATCH(pResult) = (IDispatch*)o;
}

// ============================================================
// 文件系统查询小工具
// ============================================================
// 目录/文件是否存在（区分种类）
static int fso_existsKind(const wchar_t* path, int wantDir) {
    DWORD a = path && path[0] ? GetFileAttributesW(path) : INVALID_FILE_ATTRIBUTES;
    if (a == INVALID_FILE_ATTRIBUTES) return 0;
    int isDir = (a & FILE_ATTRIBUTE_DIRECTORY) ? 1 : 0;
    return wantDir ? isDir : !isDir;
}

// 枚举目录下的名字（wantDir=1 子目录，0 文件）。返回 malloc 的数组，调用方 free 每项 + 数组。
static wchar_t** fso_listNames(const wchar_t* dir, int wantDir, int* pCount) {
    *pCount = 0;
    wchar_t pattern[1200];
    _snwprintf(pattern, 1100, L"%s\\*", dir);
    pattern[1100] = 0;
    WIN32_FIND_DATAW fd;
    HANDLE h = FindFirstFileW(pattern, &fd);
    if (h == INVALID_HANDLE_VALUE) return NULL;
    int cap = 16, n = 0;
    wchar_t** arr = (wchar_t**)calloc((size_t)cap, sizeof(wchar_t*));
    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        int isDir = (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ? 1 : 0;
        if (isDir != (wantDir ? 1 : 0)) continue;
        if (n == cap) {
            cap *= 2;
            wchar_t** grown = (wchar_t**)realloc(arr, (size_t)cap * sizeof(wchar_t*));
            if (!grown) break;
            arr = grown;
        }
        arr[n] = _wcsdup(fd.cFileName);
        if (arr[n]) n++;
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    *pCount = n;
    return arr;
}

static void fso_freeNames(wchar_t** arr, int n) {
    if (!arr) return;
    for (int i = 0; i < n; i++) free(arr[i]);
    free(arr);
}

// ============================================================
// 集合的 IEnumVARIANT（枚举的是 Folder / File 对象）
// ============================================================
typedef struct vb6_FsoEnum {
    IEnumVARIANTVtbl* lpVtbl;
    LONG  refCount;
    int   kind;        // 元素对象种类
    wchar_t** paths;   // 完整路径快照
    int   count, index;
} vb6_FsoEnum;

static void fso_enumMake(vb6_FsoObj* coll, VARIANT* pResult);

// ============================================================
// TextStream 的读写
// ============================================================
static int ts_reserve(vb6_FsoObj* o, int need) {
    if (need <= o->cap) return 1;
    int cap = o->cap ? o->cap : 256;
    while (cap < need) cap *= 2;
    wchar_t* grown = (wchar_t*)realloc(o->buf, (size_t)cap * sizeof(wchar_t));
    if (!grown) return 0;
    o->buf = grown; o->cap = cap;
    return 1;
}

static void ts_append(vb6_FsoObj* o, const wchar_t* s, int n) {
    if (n <= 0) return;
    if (!ts_reserve(o, o->len + n + 1)) return;
    wmemcpy(o->buf + o->len, s, (size_t)n);
    o->len += n;
    o->buf[o->len] = 0;
}

// 读文件 → 缓冲。**只按 Format 实参决定编码**，不嗅 BOM：
// 实测（ai/040 §7.6）：CreateTextFile(...,True) 写出的 UTF-16 文件，用
// OpenTextFile(p, 1)（Format 缺省 = TristateUseDefault = 系统 ANSI）读回来，scrrun 给的是
// 一行乱码而不是 "wide" —— 即 scrrun 不会自己认 BOM，TristateTrue(-1) 才按 UTF-16 解码。
// 本实现跟着这条读数走（要真 VB6 的乱码也一致，不能比它“聪明”）。
static int ts_load(vb6_FsoObj* o, const wchar_t* path, int useUnicode) {
    FILE* f = _wfopen(path, L"rb");
    if (!f) return 0;
    fseek(f, 0, SEEK_END);
    long sz = ftell(f);
    fseek(f, 0, SEEK_SET);
    if (sz < 0) sz = 0;
    unsigned char* raw = (unsigned char*)malloc((size_t)sz + 4);
    if (!raw) { fclose(f); return 0; }
    size_t got = fread(raw, 1, (size_t)sz, f);
    fclose(f);
    if (useUnicode) {
        o->unicode = 1;
        size_t off = (got >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) ? 2 : 0;
        int n = (int)((got - off) / 2);
        if (!ts_reserve(o, n + 1)) { free(raw); return 0; }
        memcpy(o->buf, raw + off, (size_t)n * 2);
        o->len = n; o->buf[n] = 0;
    } else {
        int n = MultiByteToWideChar(CP_ACP, 0, (const char*)raw, (int)got, NULL, 0);
        if (n > 0 && ts_reserve(o, n + 1)) {
            MultiByteToWideChar(CP_ACP, 0, (const char*)raw, (int)got, o->buf, n);
            o->len = n; o->buf[n] = 0;
        }
    }
    free(raw);
    return 1;
}

static int ts_flush(vb6_FsoObj* o) {
    if (!o->writing || !o->path[0]) return 1;
    FILE* f = _wfopen(o->path, L"wb");
    if (!f) return 0;
    if (o->unicode) {
        const unsigned char bom[2] = { 0xFF, 0xFE };
        fwrite(bom, 1, 2, f);
        if (o->len > 0) fwrite(o->buf, sizeof(wchar_t), (size_t)o->len, f);
    } else {
        if (o->len > 0) {
            int need = WideCharToMultiByte(CP_ACP, 0, o->buf, o->len, NULL, 0, NULL, NULL);
            if (need > 0) {
                char* mb = (char*)malloc((size_t)need);
                if (mb) {
                    WideCharToMultiByte(CP_ACP, 0, o->buf, o->len, mb, need, NULL, NULL);
                    fwrite(mb, 1, (size_t)need, f);
                    free(mb);
                }
            }
        }
    }
    fclose(f);
    return 1;
}

// 父目录是否存在（CreateTextFile/OpenTextFile 失败时区分 76 / 70）
static int fso_parentExists(const wchar_t* path) {
    wchar_t parent[1024];
    wcsncpy(parent, path, 1023); parent[1023] = 0;
    fso_stripTrailSep(parent);
    wchar_t* cut = wcsrchr(parent, L'\\');
    if (!cut) return fso_existsKind(L".", 1);
    if (cut == parent) return 1;
    *cut = 0;
    return fso_existsKind(parent, 1);
}

// 打开/新建 TextStream
static vb6_FsoObj* ts_open(const wchar_t* path, int iomode, int create, int format) {
    int exists = fso_existsKind(path, 0);
    if (iomode == 1) {                 // ForReading
        if (!exists) return NULL;
    } else if (iomode == 2) {          // ForWriting
        if (!exists && !create) return NULL;
    } else if (iomode == 8) {          // ForAppending
        if (!exists && !create) return NULL;
    }
    vb6_FsoObj* o = fso_objNew(FSO_K_TEXTSTREAM);
    if (!o) return NULL;
    wcsncpy(o->path, path, 1023);
    o->unicode = (format == -1) ? 1 : 0;
    o->reading = (iomode == 1);
    o->writing = (iomode != 1);
    o->line = 1; o->lineStart = 0; o->pos = 0;
    if (iomode == 1 || iomode == 8) {
        if (!ts_load(o, path, format == -1)) { free(o->buf); free(o); return NULL; }
        if (iomode == 8) o->pos = o->len;   // 追加：游标到文末
    } else {
        ts_reserve(o, 16);
        o->len = 0; if (o->buf) o->buf[0] = 0;
    }
    return o;
}

// ============================================================
// FSO 成员实现（刀1 纯路径面）
// ============================================================
static HRESULT fso_doBuildPath(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* path = fso_argStr(pdp, 0);
    wchar_t* name = fso_argStr(pdp, 1);
    if (!path || !name) { free(path); free(name); return fso_fail(pexcep, FSO_ERR_INVALID_ARG); }
    fso_stripQuotes(path); fso_stripQuotes(name);
    wchar_t out[2048];
    wcsncpy(out, path, 2000); out[2000] = 0;
    size_t n = wcslen(out);
    if (name[0] == L'\\') {
        fso_stripTrailSep(out); n = wcslen(out);
        if (n + wcslen(name) < 2047) wcscat(out, name);
    } else {
        if (n > 0 && out[n - 1] != L'\\') { if (n + 1 < 2047) { out[n++] = L'\\'; out[n] = 0; } }
        else if (n == 0) { out[0] = L'\\'; out[1] = 0; n = 1; }
        if (n + wcslen(name) < 2047) wcscat(out, name);
    }
    fso_setBstr(pResult, out);
    free(path); free(name);
    return S_OK;
}

static HRESULT fso_doGetFileName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    size_t n = wcslen(p);
    if (n > 0 && p[n - 1] == L'\\') fso_setBstr(pResult, L"");
    else fso_setBstr(pResult, fso_lastComponent(p));
    free(p);
    return S_OK;
}

static HRESULT fso_doGetBaseName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    wchar_t name[1024];
    wcsncpy(name, fso_lastComponent(p), 1023); name[1023] = 0;
    if (wcslen(p) > 0 && p[wcslen(p) - 1] == L'\\') name[0] = 0;
    wchar_t* dot = wcsrchr(name, L'.');
    if (dot) *dot = 0;
    fso_setBstr(pResult, name);
    free(p);
    return S_OK;
}

static HRESULT fso_doGetExtensionName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    const wchar_t* dot = wcsrchr(fso_lastComponent(p), L'.');
    fso_setBstr(pResult, dot ? dot + 1 : L"");
    free(p);
    return S_OK;
}

static HRESULT fso_doGetParentFolderName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    size_t n = wcslen(p);
    while (n > 0 && p[n - 1] == L'\\') p[--n] = 0;
    wchar_t* cut = wcsrchr(p, L'\\');
    if (!cut || cut == p) { fso_setBstr(pResult, L""); free(p); return S_OK; }
    if (cut - p == 2 && p[1] == L':') { cut[1] = 0; fso_setBstr(pResult, p); free(p); return S_OK; }
    *cut = 0;
    fso_setBstr(pResult, p);
    free(p);
    return S_OK;
}

static HRESULT fso_doGetAbsolutePathName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    wchar_t out[2048];
    fso_absolutePath(p, out, 2048);
    fso_setBstr(pResult, out);
    free(p);
    return S_OK;
}

static HRESULT fso_doFileExists(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    fso_setBool(pResult, fso_existsKind(p, 0));
    free(p);
    return S_OK;
}

static HRESULT fso_doFolderExists(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    fso_setBool(pResult, fso_existsKind(p, 1));
    free(p);
    return S_OK;
}

static HRESULT fso_doDriveExists(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    wchar_t root[8] = L"C:\\";
    int ok = 0;
    if (p[0]) {
        root[0] = p[0];
        UINT t = GetDriveTypeW(root);
        ok = (t != DRIVE_NO_ROOT_DIR) && ((p[0] >= L'A' && p[0] <= L'Z') || (p[0] >= L'a' && p[0] <= L'z'));
    }
    fso_setBool(pResult, ok);
    free(p);
    return S_OK;
}

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

// ---- 刀2：GetFolder / GetFile / GetDrive / GetSpecialFolder / Drives ----
static HRESULT fso_doGetFolder(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    if (!fso_existsKind(p, 1)) { free(p); return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND); }
    vb6_FsoObj* o = fso_objNew(FSO_K_FOLDER);
    if (!o) { free(p); return E_OUTOFMEMORY; }
    wcsncpy(o->path, p, 1023);
    fso_setObj(pResult, o);
    free(p);
    return S_OK;
}

static HRESULT fso_doGetFile(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    if (!fso_existsKind(p, 0)) { free(p); return fso_fail(pexcep, FSO_ERR_FILE_NOT_FOUND); }
    vb6_FsoObj* o = fso_objNew(FSO_K_FILE);
    if (!o) { free(p); return E_OUTOFMEMORY; }
    wcsncpy(o->path, p, 1023);
    fso_setObj(pResult, o);
    free(p);
    return S_OK;
}

static HRESULT fso_doGetDrive(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    wchar_t root[8] = L"C:\\";
    if (p[0]) root[0] = p[0];
    if (GetDriveTypeW(root) == DRIVE_NO_ROOT_DIR) { free(p); return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND); }
    vb6_FsoObj* o = fso_objNew(FSO_K_DRIVE);
    if (!o) { free(p); return E_OUTOFMEMORY; }
    wcsncpy(o->path, root, 1023);
    fso_setObj(pResult, o);
    free(p);
    return S_OK;
}

// 0=Windows 1=System 2=Temp（册）；真机读数待 ai/040 §5 的 D 段。
// 不用 SHGetSpecialFolderPathW：那要 shell32，本工程三种链接分支都不带；
// GetWindowsDirectoryW / GetSystemDirectoryW / GetTempPathW 同口径且零新依赖。
static HRESULT fso_doGetSpecialFolder(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    long which = fso_argLong(pdp, 0, 0);
    wchar_t path[1024] = L"";
    UINT got = 0;
    if (which == 1) got = GetSystemDirectoryW(path, 1024);
    else if (which == 2) got = GetTempPathW(1024, path);
    else got = GetWindowsDirectoryW(path, 1024);
    if (got == 0 || got >= 1024) return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND);
    size_t n = wcslen(path);
    while (n > 0 && path[n - 1] == L'\\') path[--n] = 0;
    vb6_FsoObj* o = fso_objNew(FSO_K_FOLDER);
    if (!o) return E_OUTOFMEMORY;
    wcsncpy(o->path, path, 1023);
    fso_setObj(pResult, o);
    return S_OK;
}

// Drives 集合：按盘符 A..Z 只列存在的
static HRESULT fso_doDrivesProp(VARIANT* pResult) {
    vb6_FsoObj* o = fso_objNew(FSO_K_DRIVES);
    if (!o) return E_OUTOFMEMORY;
    fso_setObj(pResult, o);
    return S_OK;
}

// ---- 刀4：写面 ----
static HRESULT fso_doCreateFolder(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    if (fso_existsKind(p, 1) || fso_existsKind(p, 0)) { free(p); return fso_fail(pexcep, FSO_ERR_ALREADY_EXISTS); }
    if (!CreateDirectoryW(p, NULL)) { free(p); return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND); }
    vb6_FsoObj* o = fso_objNew(FSO_K_FOLDER);
    if (!o) { free(p); return E_OUTOFMEMORY; }
    wcsncpy(o->path, p, 1023);
    fso_setObj(pResult, o);
    free(p);
    return S_OK;
}

static HRESULT fso_doCreateTextFile(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    long overwrite = fso_argLong(pdp, 1, -1);
    long unicode = fso_argLong(pdp, 2, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    if (fso_existsKind(p, 0) && overwrite == 0) { free(p); return fso_fail(pexcep, FSO_ERR_ALREADY_EXISTS); }
    if (!fso_parentExists(p)) { free(p); return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND); }
    vb6_FsoObj* ts = ts_open(p, 2, 1, unicode ? -1 : 0);
    free(p);
    if (!ts) return fso_fail(pexcep, FSO_ERR_PERMISSION);
    fso_setObj(pResult, ts);
    return S_OK;
}

static HRESULT fso_doOpenTextFile(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    long iomode = fso_argLong(pdp, 1, 1);
    long create = fso_argLong(pdp, 2, 0);
    long format = fso_argLong(pdp, 3, -2);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    if (!fso_existsKind(p, 0) && iomode == 1) { free(p); return fso_fail(pexcep, FSO_ERR_FILE_NOT_FOUND); }
    if (!fso_parentExists(p)) { free(p); return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND); }
    vb6_FsoObj* ts = ts_open(p, (int)iomode, (int)create, (int)format);
    free(p);
    if (!ts) return fso_fail(pexcep, FSO_ERR_FILE_NOT_FOUND);
    fso_setObj(pResult, ts);
    return S_OK;
}

static int fso_deleteOne(const wchar_t* p, int wantDir) {
    if (!fso_existsKind(p, wantDir)) return 0;
    DWORD a = GetFileAttributesW(p);
    if (a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_READONLY)) SetFileAttributesW(p, a & ~FILE_ATTRIBUTE_READONLY);
    return wantDir ? (RemoveDirectoryW(p) ? 1 : 0) : (DeleteFileW(p) ? 1 : 0);
}

// DeleteFolder 是**递归**的（册）：force 只决定要不要先清只读属性，不影响是否往下钻。
// 刀3 实测就栽在这：目录里还剩一个 inner.txt，RemoveDirectoryW 直接失败 → 76。
static int fso_deleteTree(const wchar_t* dir) {
    int cnt = 0;
    wchar_t** names = fso_listNames(dir, 0, &cnt);
    for (int i = 0; i < cnt; i++) {
        wchar_t full[1200];
        fso_join(full, 1200, dir, names[i]);
        fso_deleteOne(full, 0);
    }
    fso_freeNames(names, cnt);
    names = fso_listNames(dir, 1, &cnt);
    for (int i = 0; i < cnt; i++) {
        wchar_t full[1200];
        fso_join(full, 1200, dir, names[i]);
        fso_deleteTree(full);
    }
    fso_freeNames(names, cnt);
    return fso_deleteOne(dir, 1);
}

static HRESULT fso_doDeleteFile(DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    if (!fso_deleteOne(p, 0)) { free(p); return fso_fail(pexcep, FSO_ERR_FILE_NOT_FOUND); }
    free(p);
    return S_OK;
}

static HRESULT fso_doDeleteFolder(DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
    fso_stripQuotes(p);
    if (!fso_deleteTree(p)) { free(p); return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND); }
    free(p);
    return S_OK;
}

static HRESULT fso_doCopyFile(DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    wchar_t* src = fso_argStr(pdp, 0);
    wchar_t* dst = fso_argStr(pdp, 1);
    long overwrite = fso_argLong(pdp, 2, -1);
    if (!src || !dst) { free(src); free(dst); return fso_fail(pexcep, FSO_ERR_INVALID_ARG); }
    fso_stripQuotes(src); fso_stripQuotes(dst);
    if (!fso_existsKind(src, 0)) { free(src); free(dst); return fso_fail(pexcep, FSO_ERR_FILE_NOT_FOUND); }
    if (fso_existsKind(dst, 1)) {   // 目标是目录 ⇒ 拼文件名
        wchar_t joined[1200];
        fso_join(joined, 1200, dst, fso_lastComponent(src));
        free(dst); dst = _wcsdup(joined);
    }
    int ok = CopyFileW(src, dst, overwrite == 0);
    free(src); free(dst);
    if (!ok) return fso_fail(pexcep, FSO_ERR_PERMISSION);
    return S_OK;
}

static HRESULT fso_doMoveFile(DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    wchar_t* src = fso_argStr(pdp, 0);
    wchar_t* dst = fso_argStr(pdp, 1);
    if (!src || !dst) { free(src); free(dst); return fso_fail(pexcep, FSO_ERR_INVALID_ARG); }
    fso_stripQuotes(src); fso_stripQuotes(dst);
    if (!fso_existsKind(src, 0)) { free(src); free(dst); return fso_fail(pexcep, FSO_ERR_FILE_NOT_FOUND); }
    int ok = MoveFileW(src, dst);
    free(src); free(dst);
    if (!ok) return fso_fail(pexcep, FSO_ERR_PERMISSION);
    return S_OK;
}

// ============================================================
// GetDriveName（C30-11 第六刀）
// ============================================================
// 真引擎读数（.temp/leftover/oracle_members.vbs R05-R08）：
//   "C:\\a\\b"       -> "C:"
//   "C:"             -> "C:"
//   "\\\\srv\\share\\x" -> "\\\\srv\\share"
//   "a\\b"（相对）    -> ""（空串，不报错）
// 只答盘符（或 UNC 的 \\server\share），不带尾部反斜杠。
static HRESULT fso_doGetDriveName(DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    (void)pexcep;
    wchar_t* p = fso_argStr(pdp, 0);
    if (!p) return E_OUTOFMEMORY;
    wchar_t out[1200] = L"";
    if (p[0] == L'\\' && p[1] == L'\\') {
        const wchar_t* s = p + 2;
        const wchar_t* cut1 = wcschr(s, L'\\');
        if (cut1) {
            const wchar_t* s2 = cut1 + 1;
            const wchar_t* cut2 = wcschr(s2, L'\\');
            size_t n = cut2 ? (size_t)(cut2 - p) : wcslen(p);
            if (n > 1199) n = 1199;
            wcsncpy(out, p, n); out[n] = 0;
        }
    } else if (((p[0] >= L'A' && p[0] <= L'Z') || (p[0] >= L'a' && p[0] <= L'z')) && p[1] == L':') {
        out[0] = p[0]; out[1] = L':'; out[2] = 0;
    }
    fso_setBstr(pResult, out);
    free(p);
    return S_OK;
}

// ============================================================
// Folder / File / Drive 的属性
// ============================================================
static HRESULT fso_doPathProp(vb6_FsoObj* o, int dispId, VARIANT* pResult) {
    if (!pResult) return S_OK;
    switch (dispId) {
        case FD_NAME: {
            const wchar_t* base = fso_lastComponent(o->path);
            if (o->kind == FSO_K_DRIVE) { wchar_t d[4] = { o->path[0], L':', 0, 0 }; fso_setBstr(pResult, d); return S_OK; }
            fso_setBstr(pResult, base);
            return S_OK;
        }
        case FD_PATH:
            if (o->kind == FSO_K_DRIVE) { wchar_t d[4] = { o->path[0], L':', 0, 0 }; fso_setBstr(pResult, d); return S_OK; }
            fso_setBstr(pResult, o->path);
            return S_OK;
        case FD_SHORTPATH: {
            wchar_t shortPath[1200];
            DWORD n = GetShortPathNameW(o->path, shortPath, 1200);
            fso_setBstr(pResult, (n > 0 && n < 1200) ? shortPath : o->path);
            return S_OK;
        }
        case FD_SHORTNAME: {
            /* C30-11 第六刀（真引擎读数）: Folder.ShortName=rd1、File.ShortName=orig.txt ——
             * 交的是**名字**的 8.3 形，不是整条路径（那是 ShortPath）。
             * NTFS 关掉 8.3 生成时 GetShortPathNameW 原样返回 ⇒ 退化到与 ShortPath 同一档。 */
            wchar_t sp[1200];
            DWORD n = GetShortPathNameW(o->path, sp, 1200);
            const wchar_t* src = (n > 0 && n < 1200) ? sp : o->path;
            if (o->kind == FSO_K_DRIVE) { wchar_t d[4] = { o->path[0], L':', 0, 0 }; fso_setBstr(pResult, d); return S_OK; }
            fso_setBstr(pResult, fso_lastComponent(src));
            return S_OK;
        }
        case FD_TYPE: {
            if (o->kind == FSO_K_FOLDER) {
                wchar_t t[512];
                fso_typeName(o->path, 1, t, 512);
                fso_setBstr(pResult, t);
                return S_OK;
            }
            if (o->kind == FSO_K_DRIVE) {
                UINT t = GetDriveTypeW(o->path);
                const wchar_t* s = (t == DRIVE_REMOVABLE) ? L"Removable"
                                 : (t == DRIVE_FIXED) ? L"Fixed"
                                 : (t == DRIVE_REMOTE) ? L"Network"
                                 : (t == DRIVE_CDROM) ? L"CD-ROM"
                                 : (t == DRIVE_RAMDISK) ? L"RAM Disk" : L"Unknown";
                fso_setBstr(pResult, s);
                return S_OK;
            }
            wchar_t t[512];
            fso_typeName(o->path, 0, t, 512);
            fso_setBstr(pResult, t);
            return S_OK;
        }
        case FD_SIZE: {
            if (o->kind == FSO_K_FOLDER) {   // 递归内容总和（见 fso_dirSize 的读数注释）
                V_VT(pResult) = VT_R8;
                V_R8(pResult) = fso_dirSize(o->path);
                return S_OK;
            }
            WIN32_FILE_ATTRIBUTE_DATA fad;
            if (!GetFileAttributesExW(o->path, GetFileExInfoStandard, &fad)) { fso_setLong(pResult, 0); return S_OK; }
            fso_setLong(pResult, (long)fad.nFileSizeLow);
            return S_OK;
        }
        case FD_DATECREATED: case FD_DATELASTMODIFIED: case FD_DATELASTACCESSED: {
            WIN32_FILE_ATTRIBUTE_DATA fad;
            if (!GetFileAttributesExW(o->path, GetFileExInfoStandard, &fad)) { fso_setDate(pResult, 0.0); return S_OK; }
            const FILETIME* ft = (dispId == FD_DATECREATED) ? &fad.ftCreationTime
                               : (dispId == FD_DATELASTMODIFIED) ? &fad.ftLastWriteTime
                               : &fad.ftLastAccessTime;
            fso_setDate(pResult, fso_timeToDate(ft));
            return S_OK;
        }
        case FD_ATTRIBUTES: {
            DWORD a = GetFileAttributesW(o->path);
            fso_setLong(pResult, a == INVALID_FILE_ATTRIBUTES ? 0 : fso_attrs(a));
            return S_OK;
        }
        case FD_PARENTFOLDER: {
            if (o->kind == FSO_K_DRIVE) { fso_setObj(pResult, o); return S_OK; }   // 盘根没有父文件夹 → Nothing
            wchar_t parent[1024];
            wcsncpy(parent, o->path, 1023); parent[1023] = 0;
            fso_stripTrailSep(parent);
            wchar_t* cut = wcsrchr(parent, L'\\');
            if (!cut || cut == parent) { 
                if (cut == parent) { parent[1] = 0; }
                else { if (pResult) V_VT(pResult) = VT_EMPTY; return S_OK; }
            } else {
                *cut = 0;
                if (wcslen(parent) == 2 && parent[1] == L':') wcscat(parent, L"\\");
            }
            vb6_FsoObj* pf = fso_objNew(FSO_K_FOLDER);
            if (!pf) return E_OUTOFMEMORY;
            wcsncpy(pf->path, parent, 1023);
            fso_setObj(pResult, pf);
            return S_OK;
        }
        case FD_DRIVE: {
            wchar_t root[8] = L"C:\\";
            root[0] = o->path[0];
            vb6_FsoObj* d = fso_objNew(FSO_K_DRIVE);
            if (!d) return E_OUTOFMEMORY;
            wcsncpy(d->path, root, 1023);
            fso_setObj(pResult, d);
            return S_OK;
        }
        default: return E_NOTIMPL;
    }
}

// Drive 专有
static HRESULT fso_doDriveProp(vb6_FsoObj* o, int dispId, VARIANT* pResult) {
    switch (dispId) {
        case FD_DRIVELETTER: { wchar_t d[2] = { o->path[0], 0 }; fso_setBstr(pResult, d); return S_OK; }
        case FD_ROOTFOLDER: {
            vb6_FsoObj* f = fso_objNew(FSO_K_FOLDER);
            if (!f) return E_OUTOFMEMORY;
            wcsncpy(f->path, o->path, 1023);
            fso_setObj(pResult, f);
            return S_OK;
        }
        case FD_ISREADY: {
            wchar_t root[8] = L"C:\\";
            root[0] = o->path[0];
            fso_setBool(pResult, GetVolumeInformationW(root, NULL, 0, NULL, NULL, NULL, NULL, 0) != 0);
            return S_OK;
        }
        case FD_DRIVETYPE: {
            UINT t = GetDriveTypeW(o->path);
            int v = (t == DRIVE_REMOVABLE) ? 1 : (t == DRIVE_FIXED) ? 2 : (t == DRIVE_REMOTE) ? 3
                  : (t == DRIVE_CDROM) ? 4 : (t == DRIVE_RAMDISK) ? 5 : 0;
            fso_setLong(pResult, v);
            return S_OK;
        }
        case FD_FILESYSTEM: {
            wchar_t fs[64] = L"";
            GetVolumeInformationW(o->path, NULL, 0, NULL, NULL, NULL, fs, 64);
            fso_setBstr(pResult, fs);
            return S_OK;
        }
        case FD_VOLUMENAME: {
            wchar_t vol[256] = L"";
            GetVolumeInformationW(o->path, vol, 256, NULL, NULL, NULL, NULL, 0);
            fso_setBstr(pResult, vol);
            return S_OK;
        }
        case FD_SERIALNUMBER: {
            DWORD serial = 0;
            GetVolumeInformationW(o->path, NULL, 0, &serial, NULL, NULL, NULL, 0);
            fso_setLong(pResult, (long)serial);
            return S_OK;
        }
        case FD_TOTALSIZE: case FD_FREESPACE: case FD_AVAILABLESPACE: {
            ULARGE_INTEGER freeBytes, total, totalFree;
            if (!GetDiskFreeSpaceExW(o->path, &freeBytes, &total, &totalFree)) { fso_setLong(pResult, 0); return S_OK; }
            double v = (dispId == FD_TOTALSIZE) ? (double)total.QuadPart
                     : (dispId == FD_FREESPACE) ? (double)freeBytes.QuadPart
                     : (double)totalFree.QuadPart;
            V_VT(pResult) = VT_R8;
            V_R8(pResult) = v;
            return S_OK;
        }
        case FD_SHARENAME: {
            /* C30-11 第六刀（真引擎读数）: 本地盘 Drive.ShareName = ""（空串）；
             * UNC 盘（\\srv\share）给 share 名。这里按 o->path 的形态自己拆，不靠网络。 */
            wchar_t out[512] = L"";
            if (o->path[0] == L'\\' && o->path[1] == L'\\') {
                const wchar_t* s = o->path + 2;
                const wchar_t* cut1 = wcschr(s, L'\\');
                if (cut1) {
                    const wchar_t* s2 = cut1 + 1;
                    const wchar_t* cut2 = wcschr(s2, L'\\');
                    size_t n = cut2 ? (size_t)(cut2 - s2) : wcslen(s2);
                    if (n > 511) n = 511;
                    wcsncpy(out, s2, n); out[n] = 0;
                }
            }
            fso_setBstr(pResult, out);
            return S_OK;
        }
        default: return E_NOTIMPL;
    }
}

// ============================================================
// TextStream 成员
// ============================================================
static HRESULT ts_doMember(vb6_FsoObj* o, int dispId, DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    switch (dispId) {
        case FD_READ: {
            long n = fso_argLong(pdp, 0, 0);
            if (n < 0) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
            if (o->pos >= o->len) { fso_setBstr(pResult, L""); return S_OK; }
            if (o->pos + n > o->len) n = o->len - o->pos;
            wchar_t* tmp = (wchar_t*)calloc((size_t)n + 1, sizeof(wchar_t));
            if (!tmp) return E_OUTOFMEMORY;
            wmemcpy(tmp, o->buf + o->pos, (size_t)n);
            for (long i = 0; i < n; i++) if (tmp[i] == L'\n' || tmp[i] == L'\r') { o->line++; o->lineStart = o->pos + (int)i + 1; }
            o->pos += (int)n;
            fso_setBstr(pResult, tmp);
            free(tmp);
            return S_OK;
        }
        case FD_READLINE: {
            if (o->pos >= o->len) { fso_setBstr(pResult, L""); return S_OK; }
            int start = o->pos, i = o->pos;
            while (i < o->len && o->buf[i] != L'\r' && o->buf[i] != L'\n') i++;
            int n = i - start;
            wchar_t* tmp = (wchar_t*)calloc((size_t)n + 1, sizeof(wchar_t));
            if (!tmp) return E_OUTOFMEMORY;
            wmemcpy(tmp, o->buf + start, (size_t)n);
            // 吃掉行尾（CRLF / LF / CR）
            if (i < o->len && o->buf[i] == L'\r') { i++; if (i < o->len && o->buf[i] == L'\n') i++; }
            else if (i < o->len && o->buf[i] == L'\n') i++;
            o->pos = i; o->line++; o->lineStart = i;
            fso_setBstr(pResult, tmp);
            free(tmp);
            return S_OK;
        }
        case FD_READALL: {
            if (o->pos >= o->len) { fso_setBstr(pResult, L""); return S_OK; }
            int n = o->len - o->pos;
            wchar_t* tmp = (wchar_t*)calloc((size_t)n + 1, sizeof(wchar_t));
            if (!tmp) return E_OUTOFMEMORY;
            wmemcpy(tmp, o->buf + o->pos, (size_t)n);
            for (int i = 0; i < n; i++) if (tmp[i] == L'\n' || tmp[i] == L'\r') { o->line++; o->lineStart = o->pos + i + 1; }
            o->pos = o->len;
            fso_setBstr(pResult, tmp);
            free(tmp);
            return S_OK;
        }
        case FD_SKIP: {
            long n = fso_argLong(pdp, 0, 0);
            if (n < 0) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
            o->pos += (int)n;
            if (o->pos > o->len) o->pos = o->len;
            return S_OK;
        }
        case FD_SKIPLINE: {
            int i = o->pos;
            while (i < o->len && o->buf[i] != L'\r' && o->buf[i] != L'\n') i++;
            if (i < o->len && o->buf[i] == L'\r') { i++; if (i < o->len && o->buf[i] == L'\n') i++; }
            else if (i < o->len && o->buf[i] == L'\n') i++;
            o->pos = i; o->line++; o->lineStart = i;
            return S_OK;
        }
        case FD_WRITE: case FD_WRITELINE: {
            wchar_t* s = fso_argStr(pdp, 0);
            if (!s) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
            if (!o->writing) { free(s); return fso_fail(pexcep, FSO_ERR_PERMISSION); }
            ts_append(o, s, (int)wcslen(s));
            if (dispId == FD_WRITELINE) ts_append(o, L"\r\n", 2);
            free(s);
            return S_OK;
        }
        case FD_WRITEBLANKLINES: {
            long n = fso_argLong(pdp, 0, 0);
            for (long i = 0; i < n; i++) ts_append(o, L"\r\n", 2);
            return S_OK;
        }
        case FD_CLOSE: {
            if (o->writing && !ts_flush(o)) return fso_fail(pexcep, FSO_ERR_PERMISSION);
            o->writing = 0; o->reading = 0;
            return S_OK;
        }
        case FD_ATENDOFSTREAM: fso_setBool(pResult, o->pos >= o->len); return S_OK;
        case FD_ATENDOFLINE: {
            int atEnd = (o->pos >= o->len) || o->buf[o->pos] == L'\r' || o->buf[o->pos] == L'\n';
            fso_setBool(pResult, atEnd);
            return S_OK;
        }
        case FD_LINE: fso_setLong(pResult, o->line); return S_OK;
        case FD_COLUMN: fso_setLong(pResult, o->pos - o->lineStart + 1); return S_OK;
        default: return E_NOTIMPL;
    }
}

// ============================================================
// 集合：Count / Item / _NewEnum
// ============================================================
static int fso_collectPaths(vb6_FsoObj* coll, wchar_t*** pPaths) {
    // 返回条数；路径数组由调用方 free
    *pPaths = NULL;
    if (coll->kind == FSO_K_DRIVES) {
        int n = 0;
        wchar_t** arr = (wchar_t**)calloc(26, sizeof(wchar_t*));
        for (wchar_t c = L'A'; c <= L'Z'; c++) {
            wchar_t root[8] = L"C:\\";
            root[0] = c;
            if (GetDriveTypeW(root) != DRIVE_NO_ROOT_DIR) arr[n++] = _wcsdup(root);
        }
        *pPaths = arr;
        return n;
    }
    int wantDir = (coll->kind == FSO_K_FOLDERS);
    int cnt = 0;
    wchar_t** names = fso_listNames(coll->path, wantDir, &cnt);
    wchar_t** paths = (wchar_t**)calloc((size_t)(cnt > 0 ? cnt : 1), sizeof(wchar_t*));
    for (int i = 0; i < cnt; i++) {
        wchar_t full[1200];
        fso_join(full, 1200, coll->path, names[i]);
        paths[i] = _wcsdup(full);
    }
    fso_freeNames(names, cnt);
    *pPaths = paths;
    return cnt;
}

// ============================================================
// IDispatch（所有种类共用）
// ============================================================
static HRESULT STDMETHODCALLTYPE fso_QueryInterface(IDispatch* pDisp, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = pDisp;
        ((vb6_FsoObj*)pDisp)->refCount++;
        return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE fso_AddRef(IDispatch* pDisp) {
    return (ULONG)++((vb6_FsoObj*)pDisp)->refCount;
}

static ULONG STDMETHODCALLTYPE fso_Release(IDispatch* pDisp) {
    vb6_FsoObj* o = (vb6_FsoObj*)pDisp;
    LONG n = --o->refCount;
    if (n <= 0) {
        free(o->buf);
        free(o);
    }
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

// 名字 → DISPID（按对象种类）
static int fso_nameToDispid(int kind, const wchar_t* n) {
    if (kind == FSO_K_FSO) {
        if (!_wcsicmp(n, L"BuildPath")) return FD_BUILDPATH;
        if (!_wcsicmp(n, L"GetFileName")) return FD_GETFILENAME;
        if (!_wcsicmp(n, L"GetBaseName")) return FD_GETBASENAME;
        if (!_wcsicmp(n, L"GetExtensionName")) return FD_GETEXTENSIONNAME;
        if (!_wcsicmp(n, L"GetParentFolderName")) return FD_GETPARENTFOLDER;
        if (!_wcsicmp(n, L"GetAbsolutePathName")) return FD_GETABSOLUTEPATH;
        if (!_wcsicmp(n, L"GetDriveName")) return FD_GETDRIVENAME;
        if (!_wcsicmp(n, L"FileExists")) return FD_FILEEXISTS;
        if (!_wcsicmp(n, L"FolderExists")) return FD_FOLDEREXISTS;
        if (!_wcsicmp(n, L"DriveExists")) return FD_DRIVEEXISTS;
        if (!_wcsicmp(n, L"GetTempName")) return FD_GETTEMPNAME;
        if (!_wcsicmp(n, L"GetFolder")) return FD_GETFOLDER;
        if (!_wcsicmp(n, L"GetFile")) return FD_GETFILE;
        if (!_wcsicmp(n, L"GetDrive")) return FD_GETDRIVE;
        if (!_wcsicmp(n, L"GetSpecialFolder")) return FD_GETSPECIALFOLDER;
        if (!_wcsicmp(n, L"Drives")) return FD_DRIVES;
        if (!_wcsicmp(n, L"CreateFolder")) return FD_CREATEFOLDER;
        if (!_wcsicmp(n, L"DeleteFile")) return FD_DELETEFILE;
        if (!_wcsicmp(n, L"DeleteFolder")) return FD_DELETEFOLDER;
        if (!_wcsicmp(n, L"CopyFile")) return FD_COPYFILE;
        if (!_wcsicmp(n, L"MoveFile")) return FD_MOVEFILE;
        if (!_wcsicmp(n, L"CreateTextFile")) return FD_CREATETEXTFILE;
        if (!_wcsicmp(n, L"OpenTextFile")) return FD_OPENTEXTFILE;
        return 0;
    }
    if (kind == FSO_K_DRIVES || kind == FSO_K_FOLDERS || kind == FSO_K_FILES) {
        if (!_wcsicmp(n, L"Count")) return FD_COUNT;
        if (!_wcsicmp(n, L"Item")) return FD_ITEM;
        if (!_wcsicmp(n, L"_NewEnum")) return FD_NEWENUM;
        return 0;
    }
    if (kind == FSO_K_TEXTSTREAM) {
        if (!_wcsicmp(n, L"Read")) return FD_READ;
        if (!_wcsicmp(n, L"ReadLine")) return FD_READLINE;
        if (!_wcsicmp(n, L"ReadAll")) return FD_READALL;
        if (!_wcsicmp(n, L"Skip")) return FD_SKIP;
        if (!_wcsicmp(n, L"SkipLine")) return FD_SKIPLINE;
        if (!_wcsicmp(n, L"Write")) return FD_WRITE;
        if (!_wcsicmp(n, L"WriteLine")) return FD_WRITELINE;
        if (!_wcsicmp(n, L"WriteBlankLines")) return FD_WRITEBLANKLINES;
        if (!_wcsicmp(n, L"Close")) return FD_CLOSE;
        if (!_wcsicmp(n, L"AtEndOfStream")) return FD_ATENDOFSTREAM;
        if (!_wcsicmp(n, L"AtEndOfLine")) return FD_ATENDOFLINE;
        if (!_wcsicmp(n, L"Line")) return FD_LINE;
        if (!_wcsicmp(n, L"Column")) return FD_COLUMN;
        return 0;
    }
    // Folder / File / Drive 的公共属性
    if (!_wcsicmp(n, L"Name")) return FD_NAME;
    if (!_wcsicmp(n, L"Path")) return FD_PATH;
    if (!_wcsicmp(n, L"ShortPath")) return FD_SHORTPATH;
    if (!_wcsicmp(n, L"ShortName")) return FD_SHORTNAME;
    if (!_wcsicmp(n, L"Type")) return FD_TYPE;
    if (!_wcsicmp(n, L"Size")) return FD_SIZE;
    if (!_wcsicmp(n, L"DateCreated")) return FD_DATECREATED;
    if (!_wcsicmp(n, L"DateLastModified")) return FD_DATELASTMODIFIED;
    if (!_wcsicmp(n, L"DateLastAccessed")) return FD_DATELASTACCESSED;
    if (!_wcsicmp(n, L"Attributes")) return FD_ATTRIBUTES;
    if (!_wcsicmp(n, L"ParentFolder")) return FD_PARENTFOLDER;
    if (!_wcsicmp(n, L"Drive")) return FD_DRIVE;
    if (kind == FSO_K_FOLDER) {
        if (!_wcsicmp(n, L"IsRootFolder")) return FD_ISROOTFOLDER;
        if (!_wcsicmp(n, L"Files")) return FD_FILES;
        if (!_wcsicmp(n, L"SubFolders")) return FD_SUBFOLDERS;
        if (!_wcsicmp(n, L"CreateTextFile")) return FD_CREATETEXTFILE2;
        if (!_wcsicmp(n, L"Delete")) return FD_DELETE;
        if (!_wcsicmp(n, L"Copy")) return FD_COPY;
        if (!_wcsicmp(n, L"Move")) return FD_MOVE;
        return 0;
    }
    if (kind == FSO_K_FILE) {
        if (!_wcsicmp(n, L"OpenAsTextStream")) return FD_OPENASTEXTSTREAM;
        if (!_wcsicmp(n, L"Delete")) return FD_DELETE;
        if (!_wcsicmp(n, L"Copy")) return FD_COPY;
        if (!_wcsicmp(n, L"Move")) return FD_MOVE;
        return 0;
    }
    if (kind == FSO_K_DRIVE) {
        if (!_wcsicmp(n, L"DriveLetter")) return FD_DRIVELETTER;
        if (!_wcsicmp(n, L"RootFolder")) return FD_ROOTFOLDER;
        if (!_wcsicmp(n, L"DriveType")) return FD_DRIVETYPE;
        if (!_wcsicmp(n, L"FileSystem")) return FD_FILESYSTEM;
        if (!_wcsicmp(n, L"IsReady")) return FD_ISREADY;
        if (!_wcsicmp(n, L"VolumeName")) return FD_VOLUMENAME;
        if (!_wcsicmp(n, L"ShareName")) return FD_SHARENAME;
        if (!_wcsicmp(n, L"TotalSize")) return FD_TOTALSIZE;
        if (!_wcsicmp(n, L"FreeSpace")) return FD_FREESPACE;
        if (!_wcsicmp(n, L"AvailableSpace")) return FD_AVAILABLESPACE;
        if (!_wcsicmp(n, L"SerialNumber")) return FD_SERIALNUMBER;
        return 0;
    }
    return 0;
}

static HRESULT STDMETHODCALLTYPE fso_GetIDsOfNames(IDispatch* pDisp, REFIID riid,
                                                   LPOLESTR* rgszNames, UINT cNames,
                                                   LCID lcid, DISPID* rgDispId) {
    (void)riid; (void)lcid;
    if (!rgszNames || !rgDispId) return E_POINTER;
    vb6_FsoObj* o = (vb6_FsoObj*)pDisp;
    for (UINT i = 0; i < cNames; i++) {
        const wchar_t* n = rgszNames[i];
        int id = n ? fso_nameToDispid(o->kind, n) : 0;
        if (id == 0) { rgDispId[i] = DISPID_UNKNOWN; return DISP_E_UNKNOWNNAME; }
        rgDispId[i] = id;
    }
    return S_OK;
}

// 集合的 Item / Count
static HRESULT fso_collectionMember(vb6_FsoObj* o, int dispId, DISPPARAMS* pdp,
                                    VARIANT* pResult, EXCEPINFO* pexcep) {
    wchar_t** paths = NULL;
    int n = fso_collectPaths(o, &paths);
    HRESULT hr = S_OK;
    if (dispId == FD_COUNT) {
        fso_setLong(pResult, n);
    } else if (dispId == FD_ITEM) {
        /* C30-11 后记（遗留问题①）: scrrun 的 Files/Folders/Drives.Item 是**按名字**取项，
         * 不认序号 —— 真 VB 引擎读数（VBScript，与 VB6 同一套 oleaut32 晚绑定）:
         *   files.Item(1)        -> 运行错误 5「无效的过程调用或参数」
         *   files.Item("a.txt")  -> a.txt
         * 第二刀曾把数字索引当 1 基序号受理（比 VB6 宽松，读数是 a.txt）；
         * 现在按读数对齐: 非 BSTR 实参一律错误 5，按名字找不到也是 5。
         * 换言之 `coll(i)` / `coll.Item(i)` 在 VB6 里本来就不成立，正确形状是
         * `For Each` 或 `coll.Item(名字)`。 */
        VARIANT* v = fso_argRaw(pdp, 0);
        const wchar_t* key = NULL;
        if (v) {
            if (V_VT(v) == VT_BSTR) key = V_BSTR(v);
            else if (V_VT(v) == (VT_BSTR | VT_BYREF)) key = V_BSTRREF(v) ? *V_BSTRREF(v) : NULL;
        }
        int idx = -1;
        if (key) {
            for (int i = 0; i < n; i++) {
                const wchar_t* base = fso_lastComponent(paths[i]);
                if (_wcsicmp(base, key) == 0) { idx = i; break; }
            }
        }
        if (idx < 0) {
            /* 两档分开: 实参根本不是名字（数字等）-> 5「无效的过程调用或参数」；
             * 名字给了但集合里没有 -> 53「文件未找到」（真 VB 引擎读数:
             * files.Item("1") 就是 53，与数字实参的 5 是两个号）。 */
            hr = fso_fail(pexcep, key ? 53 : 5);
        }
        else {
            int kind = (o->kind == FSO_K_FOLDERS) ? FSO_K_FOLDER
                     : (o->kind == FSO_K_FILES) ? FSO_K_FILE : FSO_K_DRIVE;
            vb6_FsoObj* item = fso_objNew(kind);
            if (!item) hr = E_OUTOFMEMORY;
            else { wcsncpy(item->path, paths[idx], 1023); fso_setObj(pResult, item); }
        }
    }
    for (int i = 0; i < n; i++) free(paths[i]);
    free(paths);
    return hr;
}

static HRESULT STDMETHODCALLTYPE fso_Invoke(IDispatch* pDisp, DISPID dispIdMember,
                                            REFIID riid, LCID lcid, WORD wFlags,
                                            DISPPARAMS* pdp, VARIANT* pVarResult,
                                            EXCEPINFO* pexcep, UINT* puArgErr) {
    (void)riid; (void)lcid; (void)puArgErr;
    vb6_FsoObj* o = (vb6_FsoObj*)pDisp;
    int isCall = (wFlags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET)) != 0;

    if (o->kind == FSO_K_FSO) {
        if (!isCall) return fso_fail(pexcep, FSO_ERR_NO_MEMBER);
        switch (dispIdMember) {
            case FD_BUILDPATH:        return fso_doBuildPath(pdp, pVarResult, pexcep);
            case FD_GETFILENAME:      return fso_doGetFileName(pdp, pVarResult, pexcep);
            case FD_GETBASENAME:      return fso_doGetBaseName(pdp, pVarResult, pexcep);
            case FD_GETEXTENSIONNAME: return fso_doGetExtensionName(pdp, pVarResult, pexcep);
            case FD_GETPARENTFOLDER:  return fso_doGetParentFolderName(pdp, pVarResult, pexcep);
            case FD_GETABSOLUTEPATH:  return fso_doGetAbsolutePathName(pdp, pVarResult, pexcep);
            case FD_GETDRIVENAME:     return fso_doGetDriveName(pdp, pVarResult, pexcep);
            case FD_FILEEXISTS:       return fso_doFileExists(pdp, pVarResult, pexcep);
            case FD_FOLDEREXISTS:     return fso_doFolderExists(pdp, pVarResult, pexcep);
            case FD_DRIVEEXISTS:      return fso_doDriveExists(pdp, pVarResult, pexcep);
            case FD_GETTEMPNAME:      return fso_doGetTempName(pVarResult);
            case FD_GETFOLDER:        return fso_doGetFolder(pdp, pVarResult, pexcep);
            case FD_GETFILE:          return fso_doGetFile(pdp, pVarResult, pexcep);
            case FD_GETDRIVE:         return fso_doGetDrive(pdp, pVarResult, pexcep);
            case FD_GETSPECIALFOLDER: return fso_doGetSpecialFolder(pdp, pVarResult, pexcep);
            case FD_DRIVES:           return fso_doDrivesProp(pVarResult);
            case FD_CREATEFOLDER:     return fso_doCreateFolder(pdp, pVarResult, pexcep);
            case FD_DELETEFILE:       return fso_doDeleteFile(pdp, pexcep);
            case FD_DELETEFOLDER:     return fso_doDeleteFolder(pdp, pexcep);
            case FD_COPYFILE:         return fso_doCopyFile(pdp, pexcep);
            case FD_MOVEFILE:         return fso_doMoveFile(pdp, pexcep);
            case FD_CREATETEXTFILE:   return fso_doCreateTextFile(pdp, pVarResult, pexcep);
            case FD_OPENTEXTFILE:     return fso_doOpenTextFile(pdp, pVarResult, pexcep);
            default:                  return fso_fail(pexcep, FSO_ERR_NO_MEMBER);
        }
    }

    if (o->kind == FSO_K_TEXTSTREAM) {
        if (!isCall && !(wFlags & DISPATCH_PROPERTYPUT)) {
            // 属性读也要放行（AtEndOfStream 等）
        }
        if (dispIdMember == FD_WRITE || dispIdMember == FD_WRITELINE || dispIdMember == FD_WRITEBLANKLINES ||
            dispIdMember == FD_CLOSE || dispIdMember == FD_SKIP || dispIdMember == FD_SKIPLINE)
            if (!(wFlags & DISPATCH_METHOD)) return fso_fail(pexcep, FSO_ERR_NO_MEMBER);
        return ts_doMember(o, dispIdMember, pdp, pVarResult, pexcep);
    }

    if (o->kind == FSO_K_DRIVES || o->kind == FSO_K_FOLDERS || o->kind == FSO_K_FILES) {
        if (dispIdMember == FD_NEWENUM) { fso_enumMake(o, pVarResult); return S_OK; }
        /* C30-11 后记（遗留问题①）: `coll(x)` 在 VB6 里编到**默认成员** DISPID_VALUE(0)，
         * 而这三族的默认成员就是 Item。以前 0 号落到本函数后既不匹配 Count 也不匹配 Item，
         * 交回 Nothing（本机读数 `coll(1)` -> TypeName=Nothing）。收到 Item 上，
         * 后面就按 Item 的口径答（数字实参 = 错误 5，与 VB6 一致）。 */
        if (dispIdMember == DISPID_VALUE) dispIdMember = FD_ITEM;
        return fso_collectionMember(o, dispIdMember, pdp, pVarResult, pexcep);
    }

    /* C30-11 第六刀：写属性面（真引擎读数 .temp/leftover/oracle_members.vbs）
     *   File.Name = "x"           -> 改名；撞已有文件**覆盖**，err=0（R12）
     *   Folder.Name = "x"         -> 改名；撞已有目录 err=58（R20），自改也 58（R37）
     *   File/Folder.Attributes=n  -> SetFileAttributesW；目录读回恒含 16（R35：写 32 读 48）
     * 其余 put（Drive.ShareName / File.ShortName 之类只读）一律 438 —— 真引擎那边
     * 同样是「拒绝」（R40 直接抛错）。*/
    if ((o->kind == FSO_K_FOLDER || o->kind == FSO_K_FILE || o->kind == FSO_K_DRIVE)
        && (wFlags & DISPATCH_PROPERTYPUT)) {
        if (dispIdMember == FD_NAME && o->kind != FSO_K_DRIVE) {
            wchar_t* nm = fso_argStr(pdp, 0);
            if (!nm || !*nm) { free(nm); return fso_fail(pexcep, FSO_ERR_BAD_FILE_NAME); }
            if (wcschr(nm, L'\\') || wcschr(nm, L'/') || wcschr(nm, L':')) { free(nm); return fso_fail(pexcep, FSO_ERR_BAD_FILE_NAME); }
            wchar_t src[1024];
            wcsncpy(src, o->path, 1023); src[1023] = 0;
            fso_stripTrailSep(src);
            wchar_t* cut = wcsrchr(src, L'\\');
            if (!cut) { free(nm); return fso_fail(pexcep, FSO_ERR_INVALID_ARG); }
            wchar_t dst[1200];
            size_t prefix = (size_t)(cut - src) + 1;
            if (prefix >= 1024) { free(nm); return fso_fail(pexcep, FSO_ERR_INVALID_ARG); }
            wcsncpy(dst, src, prefix);
            size_t room = sizeof(dst) / sizeof(wchar_t) - 1 - prefix;
            wcsncpy(dst + prefix, nm, room);
            dst[prefix + room] = 0;
            free(nm);
            if (!_wcsicmp(dst, o->path) || !_wcsicmp(dst, src)) {
                if (o->kind == FSO_K_FILE) return S_OK;                 // 文件自改：不报错
                return fso_fail(pexcep, FSO_ERR_ALREADY_EXISTS);        // 目录撞名（含自改）：逐字读数 58
            }
            BOOL ok;
            if (o->kind == FSO_K_FOLDER) ok = MoveFileW(o->path, dst);
            else ok = MoveFileExW(o->path, dst, MOVEFILE_REPLACE_EXISTING);   // R12：文件撞名是覆盖
            if (!ok) {
                DWORD e = GetLastError();
                int err = (e == ERROR_ALREADY_EXISTS || e == ERROR_FILE_EXISTS) ? FSO_ERR_ALREADY_EXISTS
                        : (e == ERROR_FILE_NOT_FOUND)  ? FSO_ERR_FILE_NOT_FOUND
                        : (e == ERROR_PATH_NOT_FOUND)  ? FSO_ERR_PATH_NOT_FOUND
                        : FSO_ERR_PERMISSION;
                return fso_fail(pexcep, err);
            }
            wcsncpy(o->path, dst, 1023); o->path[1023] = 0;   // R11：改完 f.Path 就是新名
            return S_OK;
        }
        if (dispIdMember == FD_ATTRIBUTES && o->kind != FSO_K_DRIVE) {
            long want = fso_argLong(pdp, 0, -1);
            if (want < 0) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
            DWORD a = 0;
            if (want & 1)  a |= FILE_ATTRIBUTE_READONLY;
            if (want & 2)  a |= FILE_ATTRIBUTE_HIDDEN;
            if (want & 4)  a |= FILE_ATTRIBUTE_SYSTEM;
            if (want & 32) a |= FILE_ATTRIBUTE_ARCHIVE;
            if (o->kind == FSO_K_FOLDER) a |= FILE_ATTRIBUTE_DIRECTORY;   // R35/R36：目录位恒在
            if (!SetFileAttributesW(o->path, a)) return fso_fail(pexcep, FSO_ERR_PERMISSION);
            return S_OK;
        }
        return fso_fail(pexcep, FSO_ERR_NO_MEMBER);
    }

    // Folder / File / Drive
    if (dispIdMember == FD_FILES || dispIdMember == FD_SUBFOLDERS) {
        vb6_FsoObj* coll = fso_objNew(dispIdMember == FD_FILES ? FSO_K_FILES : FSO_K_FOLDERS);
        if (!coll) return E_OUTOFMEMORY;
        wcsncpy(coll->path, o->path, 1023);
        fso_setObj(pVarResult, coll);
        return S_OK;
    }
    if (dispIdMember == FD_ISROOTFOLDER) {
        wchar_t tmp[1024]; wcsncpy(tmp, o->path, 1023); tmp[1023] = 0;
        fso_stripTrailSep(tmp);      // "C:\" 会被剥成 "C:"，下面两种形状都算根
        size_t n = wcslen(tmp);
        fso_setBool(pVarResult, n >= 2 && tmp[1] == L':' && (n == 2 || n == 3));
        return S_OK;
    }
    if (dispIdMember == FD_CREATETEXTFILE2) {
        wchar_t* name = fso_argStr(pdp, 0);
        long overwrite = fso_argLong(pdp, 1, -1);
        long unicode = fso_argLong(pdp, 2, 0);
        if (!name) return fso_fail(pexcep, FSO_ERR_INVALID_ARG);
        wchar_t full[1200];
        fso_join(full, 1200, o->path, name);
        free(name);
        if (!fso_existsKind(o->path, 1)) return fso_fail(pexcep, FSO_ERR_PATH_NOT_FOUND);
        vb6_FsoObj* ts = ts_open(full, 2, 1, unicode ? -1 : 0);
        if (!ts) return fso_fail(pexcep, FSO_ERR_PERMISSION);
        (void)overwrite;
        fso_setObj(pVarResult, ts);
        return S_OK;
    }
    if (dispIdMember == FD_OPENASTEXTSTREAM) {
        long iomode = fso_argLong(pdp, 0, 1);
        vb6_FsoObj* ts = ts_open(o->path, (int)iomode, 0, -2);
        if (!ts) return fso_fail(pexcep, FSO_ERR_FILE_NOT_FOUND);
        fso_setObj(pVarResult, ts);
        return S_OK;
    }
    if (dispIdMember == FD_DELETE) {
        int ok = (o->kind == FSO_K_FOLDER) ? fso_deleteTree(o->path) : fso_deleteOne(o->path, 0);
        return ok ? S_OK : fso_fail(pexcep, o->kind == FSO_K_FOLDER ? FSO_ERR_PATH_NOT_FOUND : FSO_ERR_FILE_NOT_FOUND);
    }
    HRESULT hr = fso_doPathProp(o, dispIdMember, pVarResult);
    if (hr == E_NOTIMPL && o->kind == FSO_K_DRIVE) hr = fso_doDriveProp(o, dispIdMember, pVarResult);
    if (hr == E_NOTIMPL) return fso_fail(pexcep, FSO_ERR_NO_MEMBER);
    return hr;
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
// IEnumVARIANT
// ============================================================
static HRESULT STDMETHODCALLTYPE fe_QueryInterface(IEnumVARIANT* pEnum, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IEnumVARIANT)) {
        *ppv = pEnum; ((vb6_FsoEnum*)pEnum)->refCount++;
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE fe_AddRef(IEnumVARIANT* pEnum) {
    return (ULONG)++((vb6_FsoEnum*)pEnum)->refCount;
}

static ULONG STDMETHODCALLTYPE fe_Release(IEnumVARIANT* pEnum) {
    vb6_FsoEnum* e = (vb6_FsoEnum*)pEnum;
    LONG n = --e->refCount;
    if (n <= 0) {
        for (int i = 0; i < e->count; i++) free(e->paths[i]);
        free(e->paths);
        free(e);
    }
    return (ULONG)n;
}

static HRESULT STDMETHODCALLTYPE fe_Next(IEnumVARIANT* pEnum, ULONG celt, VARIANT* rgVar, ULONG* pCeltFetched) {
    vb6_FsoEnum* e = (vb6_FsoEnum*)pEnum;
    ULONG got = 0;
    while (got < celt && e->index < e->count) {
        vb6_FsoObj* item = fso_objNew(e->kind);
        if (!item) break;
        wcsncpy(item->path, e->paths[e->index], 1023);
        V_VT(&rgVar[got]) = VT_DISPATCH;
        V_DISPATCH(&rgVar[got]) = (IDispatch*)item;
        e->index++; got++;
    }
    if (pCeltFetched) *pCeltFetched = got;
    return (got == celt) ? S_OK : S_FALSE;
}

static HRESULT STDMETHODCALLTYPE fe_Skip(IEnumVARIANT* pEnum, ULONG celt) {
    vb6_FsoEnum* e = (vb6_FsoEnum*)pEnum;
    e->index += (int)celt;
    if (e->index > e->count) e->index = e->count;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE fe_Reset(IEnumVARIANT* pEnum) {
    ((vb6_FsoEnum*)pEnum)->index = 0;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE fe_Clone(IEnumVARIANT* pEnum, IEnumVARIANT** ppEnum) {
    vb6_FsoEnum* e = (vb6_FsoEnum*)pEnum;
    if (!ppEnum) return E_POINTER;
    *ppEnum = NULL;
    vb6_FsoEnum* c = (vb6_FsoEnum*)calloc(1, sizeof(vb6_FsoEnum));
    if (!c) return E_OUTOFMEMORY;
    c->lpVtbl = &g_fsoEnumVtbl;
    c->refCount = 1;
    c->kind = e->kind;
    c->count = e->count;
    c->index = e->index;
    c->paths = (wchar_t**)calloc((size_t)(e->count > 0 ? e->count : 1), sizeof(wchar_t*));
    for (int i = 0; i < e->count; i++) c->paths[i] = _wcsdup(e->paths[i]);
    *ppEnum = (IEnumVARIANT*)c;
    return S_OK;
}

static IEnumVARIANTVtbl g_fsoEnumVtbl = {
    fe_QueryInterface, fe_AddRef, fe_Release,
    fe_Next, fe_Skip, fe_Reset, fe_Clone,
};

static void fso_enumMake(vb6_FsoObj* coll, VARIANT* pResult) {
    wchar_t** paths = NULL;
    int n = fso_collectPaths(coll, &paths);
    vb6_FsoEnum* e = (vb6_FsoEnum*)calloc(1, sizeof(vb6_FsoEnum));
    if (!e) { for (int i = 0; i < n; i++) free(paths[i]); free(paths); return; }
    e->lpVtbl = &g_fsoEnumVtbl;
    e->refCount = 1;
    e->kind = (coll->kind == FSO_K_FOLDERS) ? FSO_K_FOLDER
            : (coll->kind == FSO_K_FILES) ? FSO_K_FILE : FSO_K_DRIVE;
    e->paths = paths;
    e->count = n;
    e->index = 0;
    if (pResult) {
        V_VT(pResult) = VT_UNKNOWN;
        V_UNKNOWN(pResult) = (IUnknown*)e;
    } else {
        fe_Release((IEnumVARIANT*)e);
    }
}

// ============================================================
// 对外入口
// ============================================================
// 内建实现是否启用：**默认开**（2026-11-18 用户要求翻面）。
// C3_FSO_PREFER_MS=1 → 让回真 scrrun（A/B 对照）。
int vb6_Fso_NativeEnabled(void) {
    static int on = -1;
    if (on < 0) on = (GetEnvironmentVariableW(L"C3_FSO_PREFER_MS", NULL, 0) > 0) ? 0 : 1;
    return on;
}

int32_t vb6_Fso_IsProgramId(const wchar_t* progId) {
    if (!progId || !*progId) return 0;
    if (_wcsicmp(progId, L"Scripting.FileSystemObject") == 0) return 1;
    if (_wcsicmp(progId, L"Scripting.FileSystemObject.1") == 0) return 1;
    return 0;
}

void* vb6_Fso_New(void) {
    vb6_FsoObj* o = fso_objNew(FSO_K_FSO);
    if (!o) return NULL;
    vb6_ComRememberProgID(o, L"Scripting.FileSystemObject");
    return (void*)o;
}

int32_t vb6_Fso_IsClassName(const wchar_t* className) {
    if (!className || !*className) return 0;
    if (_wcsicmp(className, L"FileSystemObject") == 0) return 1;
    if (_wcsicmp(className, L"Scripting.FileSystemObject") == 0) return 1;
    return 0;
}

// 只认 FSO 根对象：刀2 起 Folder/File/Drive/TextStream/集合都共用这张 vtable，
// 若按 vtable 比对，`TypeName` / 类名判定会把它们一律读成 FileSystemObject。
int vb6_Fso_IsInstance(void* obj) {
    if (!obj) return 0;
    if (*(void***)obj != (void**)&g_fsoVtbl) return 0;
    return ((vb6_FsoObj*)obj)->kind == FSO_K_FSO;
}
