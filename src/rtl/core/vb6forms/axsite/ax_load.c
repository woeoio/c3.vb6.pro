// ax_load.c - vb6forms_axsite 拆分片：OCX 实例化多路径解析 (免注册 LoadLibrary / 同目录 / CWD / 注册表) + 缇↔HIMETRIC
//
// 内容 = 拆分前 vb6forms_axsite.c 第 406~522 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_axsite_internal.h

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <shellapi.h>
#include <commctrl.h>
#endif

#include "vb6forms_axsite_internal.h"
#include <stdio.h>
#include <stdarg.h>

#include <stdlib.h>   /* malloc, free */
#include <stddef.h>   /* offsetof */
#include <oleauto.h>  /* SysAllocString, BSTR */
#include <olectl.h>   /* IPicture, OleLoadPicture, OLE_HANDLE */

#ifdef __cplusplus
extern "C" {
#endif

/* --- 按 CLSID 字符串实例化: 优先免注册 LoadLibrary(ocxPath) --- */

static HRESULT ocxCreateFromPath(const wchar_t* ocxPath, REFCLSID rclsid, void** ppUnk) {
    *ppUnk = NULL;
    if (!ocxPath || !*ocxPath) return ((HRESULT)0x800401F1L);
    HMODULE hMod = LoadLibraryW(ocxPath);
    if (!hMod) return ((HRESULT)0x800401F1L);
    typedef HRESULT (__stdcall *PFN_DllGetClassObject)(REFCLSID, REFIID, void**);
    PFN_DllGetClassObject pGet = (PFN_DllGetClassObject)GetProcAddress(hMod, "DllGetClassObject");
    if (!pGet) { FreeLibrary(hMod); return ((HRESULT)0x800401F1L); }
    IClassFactory* cf = NULL;
    HRESULT hr = pGet(rclsid, &IID_IClassFactory, (void**)&cf);
    if (FAILED(hr) || !cf) { /* FreeLibrary 不做: DllGetClassObject 成功后对象可能回引 DLL */ return hr; }
    hr = cf->lpVtbl->CreateInstance(cf, NULL, &IID_IUnknown, ppUnk);
    cf->lpVtbl->Release(cf);
    return hr;
}

/* 多路径实例化 (Fix 143):
 *   1) <exe目录>\<ocx文件名>  — 便携分发 (ocx 随 exe 走, 免注册)
 *   2) 编译期 bake 的 ocx 绝对路径 — 开发机场景
 *   3) CoCreateInstance — 已注册场景 (INPROC + LOCAL_SERVER 兼容 surrogate)
 *   4) 裸文件名 LoadLibrary — 走系统 DLL 搜索路径
 * 顺序保证: 同一 CLSID 已注册时也优先用随 exe 的文件, 与 VB6
 * "注册表指向哪用哪"不同 — 便携部署下注册表常指向不存在/旧版路径.
 * 用环境变量 C3_OCX_PREFER_REG=1 可改为注册表优先. */
HRESULT vb6_ocxCreateAny(const wchar_t* ocxPath, REFCLSID rclsid, void** ppUnk) {
    *ppUnk = NULL;
    HRESULT hr;

    /* exe 所在目录 (结尾带反斜杠) — 用 GetModuleFileName 取, 与 CWD 无关 */
    wchar_t exeDir[MAX_PATH];
    exeDir[0] = 0;
    DWORD n = GetModuleFileNameW(NULL, exeDir, MAX_PATH);
    if (n > 0 && n < MAX_PATH) {
        wchar_t* es = wcsrchr(exeDir, L'\\');
        if (es) es[1] = 0;   /* 保留结尾 '\' */
        else exeDir[0] = 0;
    } else {
        exeDir[0] = 0;
    }

    /* 从 bake 的 ocxPath 提取纯文件名 */
    const wchar_t* fname = ocxPath ? ocxPath : L"";
    if (fname && *fname) {
        const wchar_t* s = wcsrchr(fname, L'\\');
        if (!s) s = wcsrchr(fname, L'/');
        if (s) fname = s + 1;
    }

    /* 候选 1: <exe目录>\<文件名>  (OCX 与 exe 同目录, 兼容旧 bake) */
    wchar_t exeSide[MAX_PATH * 2];
    exeSide[0] = 0;
    if (exeDir[0] && fname[0] && wcslen(exeDir) + wcslen(fname) + 1 < MAX_PATH * 2) {
        wcscpy(exeSide, exeDir);
        wcscat(exeSide, fname);
    }

    /* 候选 2: <exe目录>\<ocxPath相对部分>  (VBP 写 "bin\X.ocx" → exe目录\bin\X.ocx, 免注册相对 exe 加载, 不依赖 CWD) */
    wchar_t exeRel[MAX_PATH * 2];
    exeRel[0] = 0;
    if (exeDir[0] && ocxPath && *ocxPath) {
        /* 仅对相对路径拼接 exe 目录; 绝对/UNC/盘符根路径保持原样, 不重复拼 */
        int isAbs = (wcschr(ocxPath, L':') != NULL) || (ocxPath[0] == L'\\');
        if (!isAbs && wcslen(exeDir) + wcslen(ocxPath) + 1 < MAX_PATH * 2) {
            wcscpy(exeRel, exeDir);
            wcscat(exeRel, ocxPath);
        }
    }

    static int preferReg = -1;
    if (preferReg < 0) {
        preferReg = (GetEnvironmentVariableW(L"C3_OCX_PREFER_REG", NULL, 0) > 0) ? 1 : 0;
    }

    if (!preferReg) {
        /* 便携优先 (默认): 先试"相对 exe"解析 (不依赖 CWD), 再试同目录, 再试 CWD, 最后注册表 */
        if (exeRel[0]) {
            hr = ocxCreateFromPath(exeRel, rclsid, ppUnk);
            if (SUCCEEDED(hr) && *ppUnk) return hr;
        }
        if (exeSide[0]) {
            hr = ocxCreateFromPath(exeSide, rclsid, ppUnk);
            if (SUCCEEDED(hr) && *ppUnk) return hr;
        }
        hr = ocxCreateFromPath(ocxPath, rclsid, ppUnk);
        if (SUCCEEDED(hr) && *ppUnk) return hr;
        hr = CoCreateInstance(rclsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                              &IID_IUnknown, ppUnk);
        if (SUCCEEDED(hr) && *ppUnk) return hr;
        /* 4) 裸文件名 (系统搜索路径) */
        if (ocxPath) {
            const wchar_t* slash = wcsrchr(ocxPath, L'\\');
            if (!slash) slash = wcsrchr(ocxPath, L'/');
            if (slash) return ocxCreateFromPath(slash + 1, rclsid, ppUnk);
        }
        return hr;
    } else {
        /* 注册表优先 (C3_OCX_PREFER_REG=1) */
        hr = CoCreateInstance(rclsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                              &IID_IUnknown, ppUnk);
        if (SUCCEEDED(hr) && *ppUnk) return hr;
        if (exeRel[0]) {
            hr = ocxCreateFromPath(exeRel, rclsid, ppUnk);
            if (SUCCEEDED(hr) && *ppUnk) return hr;
        }
        if (exeSide[0]) {
            hr = ocxCreateFromPath(exeSide, rclsid, ppUnk);
            if (SUCCEEDED(hr) && *ppUnk) return hr;
        }
        return ocxCreateFromPath(ocxPath, rclsid, ppUnk);
    }
}

/* twips → HIMETRIC。账 #298 (§B129): 这条比例以前在 axsite 里各写一遍
 * (这里一份、ax_site_ext 一份、vb6forms_olecon 两份), 现在只转调 vb6forms.c 那一处。
 * 名字与签名照旧 —— 它是 ax_host.c 三处调用者认的那个出口。 */
long vb6_twipsToHimetric(long twips) { return vb6_TwipsToHimetric(twips); }


#ifdef __cplusplus
} // extern "C"
#endif
