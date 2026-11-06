// vb6com_manifest.c - 免注册 COM：组件内嵌 assembly manifest 取证 + ActCtx 原语
//
// 背景（2026-11-04 实测结论，见 docs/免注册COM-SxS两条路径.md）：
//   ① 工具包参考清单（MSDatGrd.ocx.manifest / TabCtl32.Ocx.manifest）**没有 assemblyIdentity**，
//      只有 <file name=".."> + <comClass clsid tlbid>；进程清单（CreateActCtx 传入或 EXE 内嵌）
//      里同样不需要 identity —— 实测有/无都成功，且无 14001。
//   ② 单放在组件旁的 <组件文件名>.manifest **不被 SxS 自动采用**（CoCreateInstance=0x80040154），
//      它只是素材，必须并进进程清单才生效（收口在 vb6com.c 的 comComposeManifest）。
//   ③ Codejock v20.3 全家有 RT_MANIFEST #2 但只有 trustInfo；RICHTX32/NewTab01 连 RT_MANIFEST 都没有。
//      ⇒ manifest 模式不依赖厂商 identity，缺 identity 是事实而不是失败。
//
// 本文件只做两件事：
//   ① 取证/诊断：读文件的内嵌 RT_MANIFEST（资源类型 RT_MANIFEST = 24，id 2 = DLL 清单），
//      抽取 <assemblyIdentity> 的 name/version/publicKeyToken/language，缺哪一项就报哪一项；
//   ② ActCtx 原语：CreateActCtxW + ActivateActCtx + 释放（策略/模式判定归 vb6com.c）。
// 纯 Win32 + CRT，无 RTL 依赖 —— tests\tools\com_manifest_probe.c 也直接编这份源文件，
// 保证"取证工具"和"运行期诊断"用的是**同一份判据**，不会两处口径漂移。

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>

#include "vb6com.h"

#define VB6_MANIFEST_RES_ID     2u          /* DLL 的清单资源 id（EXE 是 1） */
#define VB6_MANIFEST_MAX_BYTES  (1024u * 1024u)

/* ------------------------------------------------------------
 * 文本工具：manifest 是 XML 文本（UTF-8 或 ANSI），这里只做**属性级**扫描，
 * 不引入 XML 解析器 —— 判据只需要"这一项在不在、值是什么"，全量解析是第二期的事。
 * ---------------------------------------------------------- */
static const char* ci_find(const char* hay, const char* needle) {
    if (!hay || !needle || !*needle) return NULL;
    size_t n = strlen(needle);
    for (const char* p = hay; *p; p++) {
        if (_strnicmp(p, needle, n) == 0) return p;
    }
    return NULL;
}

/* tag: 指向元素起点（如 "<assemblyIdentity"），元素范围 = 到 '>' 为止。
 * attr: 属性名（不带 =）。命中返回 1 并把值写进 out（去引号）。 */
static int tag_attr(const char* tagStart, const char* attr, char* out, int cch) {
    if (out && cch > 0) out[0] = 0;
    if (!tagStart || !attr) return 0;
    const char* tagEnd = strchr(tagStart, '>');
    if (!tagEnd) return 0;
    for (const char* p = tagStart; p < tagEnd; ) {
        p = ci_find(p, attr);
        if (!p || p >= tagEnd) return 0;
        /* 必须是完整属性名：前面是空白，后面紧跟 '=' */
        const char* after = p + strlen(attr);
        int boundary = (p == tagStart) || (*(p - 1) == ' ' || *(p - 1) == '\t' || *(p - 1) == '\n' || *(p - 1) == '\r');
        while (after < tagEnd && (*after == ' ' || *after == '\t')) after++;
        if (boundary && after < tagEnd && *after == '=') {
            after++;
            while (after < tagEnd && (*after == ' ' || *after == '\t')) after++;
            char quote = 0;
            if (after < tagEnd && (*after == '"' || *after == '\'')) { quote = *after; after++; }
            int n = 0;
            while (after < tagEnd && *after && *after != quote && *after != '>' && *after != ' ') {
                if (out && n < cch - 1) out[n] = *after;
                n++; after++;
            }
            if (out && cch > 0) out[n < cch ? n : cch - 1] = 0;
            return 1;
        }
        p = after;
    }
    return 0;
}

static void missing_add(wchar_t* out, int cch, const wchar_t* item) {
    if (!out || cch <= 0) return;
    if (out[0]) {
        size_t used = wcslen(out);
        if (used + 3 < (size_t)cch) wcscat(out, L"; ");
    }
    size_t used = wcslen(out);
    size_t need = wcslen(item);
    if (used + need >= (size_t)cch) return;
    wcscat(out, item);
}

/* 分析一段 manifest 文本：抽取第一条 <assemblyIdentity>，把缺失项写进 missing。
 * 返回 1 = identity 完整（name/version/publicKeyToken 齐）；language 缺失只**报告**不判死，
 * 因为 win32 私有程序集的 assemblyIdentity 本来就可以不写 language（缺省即 "*"）。 */
static int analyze_identity(const char* text, wchar_t* ident, int identCch,
                            wchar_t* missing, int missCch) {
    if (ident && identCch > 0) ident[0] = 0;
    if (!text) { missing_add(missing, missCch, L"manifest 文本为空"); return 0; }
    const char* tag = ci_find(text, "<assemblyIdentity");
    if (!tag) {
        /* 区分两种"没有"：完全没有 assemblyIdentity vs 有 RT_MANIFEST 但只有 trustInfo */
        if (ci_find(text, "<trustInfo")) missing_add(missing, missCch, L"assemblyIdentity（资源只有 UAC trustInfo）");
        else missing_add(missing, missCch, L"assemblyIdentity");
        return 0;
    }
    char name[256] = {0}, ver[64] = {0}, tok[64] = {0}, lang[32] = {0};
    int hasName = tag_attr(tag, "name", name, (int)sizeof(name));
    int hasVer  = tag_attr(tag, "version", ver, (int)sizeof(ver));
    int hasTok  = tag_attr(tag, "publicKeyToken", tok, (int)sizeof(tok));
    int hasLang = tag_attr(tag, "language", lang, (int)sizeof(lang));
    wchar_t item[512];
    if (!hasName) missing_add(missing, missCch, L"assemblyIdentity.name");
    if (!hasVer)  missing_add(missing, missCch, L"assemblyIdentity.version");
    /* publicKeyToken / language 不是硬性项：私有程序集（无强名）的 assemblyIdentity 本来就可以不写，
     * SxS 也接受。缺了照实报告 (ident 里写 (none))，但不判死 —— 判死的只有 name/version。 */
    if (ident && identCch > 0 && hasName) {
        _snwprintf(ident, (size_t)identCch - 1,
                   L"name=%hs version=%hs publicKeyToken=%hs language=%hs",
                   name, hasVer ? ver : "?",
                   hasTok ? tok : "(none)",
                   hasLang ? lang : "(none)");
        ident[identCch - 1] = 0;
    }
    return hasName && hasVer;
}

/* ------------------------------------------------------------
 * 读文件 / 读内嵌 RT_MANIFEST
 * ---------------------------------------------------------- */
static char* read_file_bytes(const wchar_t* path, size_t* pLen) {
    if (pLen) *pLen = 0;
    HANDLE hf = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
    if (hf == INVALID_HANDLE_VALUE) return NULL;
    LARGE_INTEGER sz;
    if (!GetFileSizeEx(hf, &sz) || sz.QuadPart <= 0 || sz.QuadPart > (LONGLONG)VB6_MANIFEST_MAX_BYTES) {
        CloseHandle(hf);
        return NULL;
    }
    size_t n = (size_t)sz.QuadPart;
    char* buf = (char*)malloc(n + 1);
    if (!buf) { CloseHandle(hf); return NULL; }
    DWORD got = 0;
    if (!ReadFile(hf, buf, (DWORD)n, &got, NULL)) { free(buf); CloseHandle(hf); return NULL; }
    CloseHandle(hf);
    buf[got] = 0;
    if (pLen) *pLen = got;
    return buf;
}

static char* read_embedded_manifest(const wchar_t* path) {
    HMODULE hMod = LoadLibraryExW(path, NULL, LOAD_LIBRARY_AS_DATAFILE);
    if (!hMod) return NULL;
    HRSRC hRes = FindResourceW(hMod, MAKEINTRESOURCEW(VB6_MANIFEST_RES_ID), MAKEINTRESOURCEW(24) /* RT_MANIFEST */);
    if (!hRes) { FreeLibrary(hMod); return NULL; }
    HGLOBAL hGlb = LoadResource(hMod, hRes);
    DWORD size = SizeofResource(hMod, hRes);
    if (!hGlb || size == 0 || size > VB6_MANIFEST_MAX_BYTES) { FreeLibrary(hMod); return NULL; }
    const void* pv = LockResource(hGlb);
    if (!pv) { FreeLibrary(hMod); return NULL; }
    char* buf = (char*)malloc((size_t)size + 1);
    if (!buf) { FreeLibrary(hMod); return NULL; }
    memcpy(buf, pv, size);
    buf[size] = 0;
    FreeLibrary(hMod);
    return buf;
}

/* ------------------------------------------------------------
 * 公开：取证一个组件文件（OCX/DLL/EXE）
 *   返回 1 = 有内嵌 RT_MANIFEST 且 <assemblyIdentity> 完整（name/version/publicKeyToken）
 *   返回 0 = 不可用；missing 写清缺什么（可直接打给用户看）
 *   ident（可选）写 "<identity 四字段>"
 * ---------------------------------------------------------- */
int vb6_ComManifestInspect(const wchar_t* filePath, wchar_t* ident, int identCch,
                           wchar_t* missing, int missCch) {
    if (ident && identCch > 0) ident[0] = 0;
    if (missing && missCch > 0) missing[0] = 0;
    if (!filePath || !*filePath) { missing_add(missing, missCch, L"组件路径为空"); return 0; }
    if (GetFileAttributesW(filePath) == INVALID_FILE_ATTRIBUTES) {
        missing_add(missing, missCch, L"组件文件不存在");
        return 0;
    }
    char* text = read_embedded_manifest(filePath);
    if (!text) {
        missing_add(missing, missCch, L"内嵌 RT_MANIFEST（资源 #2 不存在）");
        return 0;
    }
    int ok = analyze_identity(text, ident, identCch, missing, missCch);
    free(text);
    return ok;
}

/* 公开：校验客户端 manifest 文件（-ComManifest 指的那一份）
 *   返回 1 = 文件存在，且按下列两种口径之一可用：
 *     a) 工具包口径（外部清单代写）：<file name="..."> 里有 <comClass>，**不需要 identity**；
 *     b) 依赖程序集口径：<dependentAssembly> 的 <assemblyIdentity> name/version 齐。
 *   返回 0 = 不可用；missing 写清缺什么 */
int vb6_ComManifestValidateClient(const wchar_t* manifestPath,
                                  wchar_t* missing, int missCch) {
    if (missing && missCch > 0) missing[0] = 0;
    if (!manifestPath || !*manifestPath) {
        missing_add(missing, missCch, L"manifest 文件路径（未给 -ComManifest / C3_COM_MANIFEST）");
        return 0;
    }
    if (GetFileAttributesW(manifestPath) == INVALID_FILE_ATTRIBUTES) {
        missing_add(missing, missCch, L"manifest 文件存在（指向的路径找不到文件）");
        return 0;
    }
    char* text = read_file_bytes(manifestPath, NULL);
    if (!text) {
        missing_add(missing, missCch, L"manifest 文件可读（打开/读取失败）");
        return 0;
    }
    const char* dep = ci_find(text, "<dependentAssembly");
    if (dep) {
        /* 只看 dependentAssembly 段：顶层 assemblyIdentity 是客户端自己的身份，不算组件身份 */
        const char* depEnd = ci_find(dep, "</dependentAssembly>");
        char depBuf[4096];
        size_t n = depEnd ? (size_t)(depEnd - dep) : 0;
        if (n == 0 || n >= sizeof(depBuf)) n = 0;
        if (n) {
            memcpy(depBuf, dep, n);
            depBuf[n] = 0;
        } else {
            /* 段太长/未闭合：退化为整文扫描（宁可误报也要给出点名） */
            strncpy(depBuf, dep, sizeof(depBuf) - 1);
            depBuf[sizeof(depBuf) - 1] = 0;
        }
        int ok = analyze_identity(depBuf, NULL, 0, missing, missCch);
        free(text);
        return ok;
    }
    /* 外部清单代写口径：<file name="x.ocx"> + <comClass clsid=.../> 就够，identity 可有可无 */
    const char* fileTag = ci_find(text, "<file");
    if (fileTag) {
        char name[512] = {0};
        int ok = 1;
        if (!tag_attr(fileTag, "name", name, (int)sizeof(name)))
            missing_add(missing, missCch, L"<file name>（<file> 没写文件名）"), ok = 0;
        if (!ci_find(fileTag, "<comClass"))
            missing_add(missing, missCch, L"<comClass>（<file> 里没有 COM 类映射）"), ok = 0;
        free(text);
        return ok;
    }
    missing_add(missing, missCch, L"<file>+<comClass> 或 <dependentAssembly>（客户端 manifest 两者都没有）");
    free(text);
    return 0;
}

/* 公开：从外部组件清单（<组件文件名>.manifest，工具包那种 per-file 清单）抽出
 * <file ...>...</file> 块，供合成聚合清单时原样复用（"优先用现成的 <文件名>.manifest"）。
 * 返回 1 = 抽到；out 为 UTF-8 文本的宽字符副本，未改写 name 属性（由调用方按目标目录重写）。 */
int vb6_ComManifestExtractFileBlock(const wchar_t* path, wchar_t* out, int cch) {
    if (out && cch > 0) out[0] = 0;
    char* text = read_file_bytes(path, NULL);
    if (!text) return 0;
    int ok = 0;
    const char* start = ci_find(text, "<file");
    if (start) {
        const char* end = ci_find(start, "</file>");
        if (end) {
            end += 7;   /* "</file>" */
            size_t n = (size_t)(end - start);
            if (cch > 0 && n < (size_t)cch) {
                for (size_t i = 0; i < n; i++) out[i] = (wchar_t)start[i];
                out[n] = 0;
                ok = 1;
            }
        }
    }
    free(text);
    return ok;
}

/* ------------------------------------------------------------
 * 公开：ActCtx 原语（进程级单例；句柄与 cookie 的释放归 vb6_ComActCtxRelease）
 * ---------------------------------------------------------- */
static HANDLE      g_comActCtx = NULL;
static ULONG_PTR   g_comActCookie = 0;

static void format_win32_error(DWORD code, wchar_t* out, int cch) {
    if (!out || cch <= 0) return;
    wchar_t msg[512];
    msg[0] = 0;
    FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
                   NULL, code, 0, msg, 512, NULL);
    // 去掉行尾的 CR/LF
    for (wchar_t* p = msg; *p; p++) if (*p == L'\r' || *p == L'\n') { *p = 0; break; }
    _snwprintf(out, (size_t)cch - 1, L"error %lu (%ls)", (unsigned long)code, msg[0] ? msg : L"?");
    out[cch - 1] = 0;
}

int vb6_ComActCtxCreate(const wchar_t* manifestPath, wchar_t* err, int errCch) {
    if (err && errCch > 0) err[0] = 0;
    if (g_comActCtx) return 1;               /* 幂等：一次进程只建一枚 */
    if (!manifestPath || !*manifestPath) {
        if (err && errCch > 0) wcsncpy(err, L"manifest 路径为空", (size_t)errCch - 1);
        return 0;
    }
    ACTCTXW ac;
    memset(&ac, 0, sizeof(ac));
    ac.cbSize = sizeof(ac);
    ac.dwFlags = 0;
    ac.lpSource = manifestPath;
    ac.lpAssemblyDirectory = NULL;
    HANDLE h = CreateActCtxW(&ac);
    // 注意: CreateActCtx 失败返回 INVALID_HANDLE_VALUE((HANDLE)-1), 不是 NULL —
    // 只判 !h 会拿 INVALID_HANDLE_VALUE 去 ActivateActCtx, 得到 error 87 掩盖真因。
    if (h == INVALID_HANDLE_VALUE || !h) {
        DWORD code = GetLastError();
        wchar_t why[512];
        format_win32_error(code, why, 512);
        if (err && errCch > 0) {
            _snwprintf(err, (size_t)errCch - 1, L"CreateActCtx failed: %ls", why);
            err[errCch - 1] = 0;
        }
        return 0;
    }
    ULONG_PTR cookie = 0;
    if (!ActivateActCtx(h, &cookie)) {
        DWORD code = GetLastError();
        wchar_t why[512];
        format_win32_error(code, why, 512);
        ReleaseActCtx(h);
        if (err && errCch > 0) {
            _snwprintf(err, (size_t)errCch - 1, L"ActivateActCtx failed: %ls", why);
            err[errCch - 1] = 0;
        }
        return 0;
    }
    g_comActCtx = h;
    g_comActCookie = cookie;
    return 1;
}

void vb6_ComActCtxRelease(void) {
    if (!g_comActCtx) return;
    DeactivateActCtx(0, g_comActCookie);
    ReleaseActCtx(g_comActCtx);
    g_comActCtx = NULL;
    g_comActCookie = 0;
}

int vb6_ComActCtxActive(void) {
    return g_comActCtx != NULL;
}
