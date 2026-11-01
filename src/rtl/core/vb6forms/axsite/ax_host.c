// ax_host.c - vb6forms_axsite 拆分片：对外入口: 窗体 Dispatch 属性 / Controls.Add / 宿主创建 10 步流程 / 消息转发与绘制
//
// 内容 = 拆分前 vb6forms_axsite.c 第 22~28 / 207~276 / 532~590 / 833~1029 行，纯搬移零重排无行为改动
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

static const wchar_t* g_FormDispatchProp = L"VB6_Form_IDispatch";

/* 账 (Outlook2013 左面板 VisualTheme): 控件类名的来源 —— CLSID 注册的 ProgID 尾段。
 * 站点创建时记进 site->typeName, 由 vb6_OcxHost_TypeNameOf 交出, vb6_ComObjectTypeCast
 * 拿它复刻 VB6 `Set <接口类型变量> = <控件>` 的类型检查 (不符 = Error 13)。 */
static void vb6_AxTypeNameFromProgId(const wchar_t* progId, wchar_t* out, int cch) {
    out[0] = 0;
    if (!progId || !*progId || cch <= 0) return;
    /* ProgID 形如 "Codejock.ShortcutCaption.20.3.0" —— 尾段是**版本号**, 不是类名。
     * 取最后一个「非纯数字」的段 = 控件类名 (与 .frm 的 `Begin XtremeShortcutBar.
     * ShortcutCaption` 尾段同一口径)。 */
    const wchar_t* best = progId;
    for (const wchar_t* p = progId; *p; ) {
        const wchar_t* dot = wcschr(p, L'.');
        size_t len = dot ? (size_t)(dot - p) : wcslen(p);
        int isNum = len > 0;
        for (size_t i = 0; i < len; i++) if (p[i] < L'0' || p[i] > L'9') { isNum = 0; break; }
        if (len > 0 && !isNum) best = p;
        if (!dot) break;
        p = dot + 1;
    }
    const wchar_t* dot2 = wcschr(best, L'.');
    size_t n = dot2 ? (size_t)(dot2 - best) : wcslen(best);
    if (n >= (size_t)cch) n = (size_t)cch - 1;
    wcsncpy(out, best, n);
    out[n] = 0;
}

static void vb6_AxTypeNameFromClsid(const wchar_t* clsidStr, wchar_t* out, int cch) {
    out[0] = 0;
    if (!clsidStr || !*clsidStr) return;
    wchar_t key[256];
    _snwprintf(key, 255, L"CLSID\\%s\\ProgID", clsidStr);
    key[255] = 0;
    HKEY hk = NULL;
    if (RegOpenKeyExW(HKEY_CLASSES_ROOT, key, 0, KEY_QUERY_VALUE, &hk) != ERROR_SUCCESS) return;
    wchar_t buf[256];
    DWORD sz = sizeof(buf) - sizeof(wchar_t);
    DWORD typ = 0;
    if (RegQueryValueExW(hk, NULL, NULL, &typ, (LPBYTE)buf, &sz) == ERROR_SUCCESS && typ == REG_SZ) {
        buf[255] = 0;
        vb6_AxTypeNameFromProgId(buf, out, cch);
    }
    RegCloseKey(hk);
}


void vb6_Form_SetDispatch(void* hwnd, void* pDispatch) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, g_FormDispatchProp, (HANDLE)pDispatch);
}

/* Fix <vbeclipse> 2026-10-06: 运行期 Controls.Add 免注册 OCX 表 (vbp Object= 同款机制).
 * cgen 在入口点烘焙 c3_ocx_libs[], 由 vb6_OcxRefRegister 登记。vb6_Form_ControlsAdd
 * 按 ProgID 命中后改走 vb6_ocxCreateAny (LoadLibrary+DllGetClassObject, 绕开注册表);
 * 未命中 → 原有 CLSIDFromProgID+CoCreateInstance 注册表路径 (零回归: 表空时恒回退). */
static Vb6OcxRef g_ocxRefs[VB6_MAX_OCXREFS];
static int       g_ocxRefCount = 0;

void vb6_OcxRefRegister(const Vb6OcxRef* libs, int count) {
    if (count > VB6_MAX_OCXREFS) {
        fwprintf(stderr, L"[C3_OCX] ocx ref table full (%d > %d); truncating\n",
                 count, VB6_MAX_OCXREFS);
        count = VB6_MAX_OCXREFS;
    }
    for (int i = 0; i < count; i++) g_ocxRefs[i] = libs[i];
    g_ocxRefCount = count;
}

/* 按 ProgID 或 CLSID 字符串查免注册表; 返回 NULL = 未命中. */
static const Vb6OcxRef* vb6_OcxRefLookup(const wchar_t* progId, const wchar_t* clsidStr) {
    if (g_ocxRefCount == 0) return NULL;
    /* 1) ProgID 精确 (大小写不敏感) */
    if (progId) {
        for (int i = 0; i < g_ocxRefCount; i++) {
            if (g_ocxRefs[i].progId && _wcsicmp(g_ocxRefs[i].progId, progId) == 0)
                return &g_ocxRefs[i];
        }
        /* 2) ProgID 去版本后缀 (.1/.2) 再试 (typelib 导出版本无关 ProgID 时兜底) */
        size_t n = wcslen(progId);
        if (n > 2 && progId[n - 2] == L'.' && progId[n - 1] >= L'0' && progId[n - 1] <= L'9') {
            wchar_t base[128];
            if (n < 128) {
                wcscpy(base, progId);
                base[n - 2] = 0;
                for (int i = 0; i < g_ocxRefCount; i++)
                    if (g_ocxRefs[i].progId && _wcsicmp(g_ocxRefs[i].progId, base) == 0)
                        return &g_ocxRefs[i];
            }
        }
    }
    /* 3) CLSID 字符串精确 (注册表可解析 CLSID 时兜底) */
    if (clsidStr) {
        for (int i = 0; i < g_ocxRefCount; i++) {
            if (g_ocxRefs[i].clsidStr && _wcsicmp(g_ocxRefs[i].clsidStr, clsidStr) == 0)
                return &g_ocxRefs[i];
        }
    }
    return NULL;
}

void* vb6_Form_ControlsAdd(void* hwnd, const wchar_t* progId, const wchar_t* ctrlName) {
    if (!hwnd || !progId) return NULL;

    /* C3_OCX_TRACE=1: 打印实例化/激活各步 HRESULT (诊断宿主问题) */
    int ocxTrace = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0);

    /* 0. 免注册表优先 (按 ProgID 查, 不依赖注册表).
     *    命中 → 用表里记录的 coclass CLSID + 相对 exe 路径走 vb6_ocxCreateAny,
     *    完全绕开 HKLM 注册表 —— CI runner 未注册 OCX 也能实例化. */
    const Vb6OcxRef* hit = vb6_OcxRefLookup(progId, NULL);
    if (!hit) {
        /* 注册表可解析 CLSID 时, 再按 CLSID 试一次 (typelib progId 与用户写的
         * ProgID 不一致时兜底; 此分支仅在机器已注册该 ProgID 时才会用到). */
        CLSID c0;
        if (SUCCEEDED(CLSIDFromProgID(progId, &c0))) {
            wchar_t cs[40];
            if (StringFromGUID2(&c0, cs, 40) > 0)
                hit = vb6_OcxRefLookup(NULL, cs);
        }
    }

    IUnknown* pUnk = NULL;
    HRESULT hr = ((HRESULT)0x800401F1L);
    if (hit && hit->clsidStr && hit->fileName) {
        CLSID c;
        if (SUCCEEDED(CLSIDFromString((LPOLESTR)hit->clsidStr, &c))) {
            hr = vb6_ocxCreateAny(hit->fileName, &c, (void**)&pUnk);
            if (ocxTrace)
                fwprintf(stderr,
                         L"[C3_OCX] Controls.Add regfree hit progId=%ls ocx=%ls hr=0x%08lX pUnk=%p\n",
                         progId, hit->fileName, (unsigned long)hr, (void*)pUnk);
        }
    }

    /* 1. 注册表兜底 (原有路径): 表未命中 / vb6_ocxCreateAny 失败 (OCX 不在 exe 旁). */
    if (!pUnk) {
        CLSID clsid;
        hr = CLSIDFromProgID(progId, &clsid);
        if (FAILED(hr)) {
            if (ocxTrace)
                fwprintf(stderr, L"[C3_OCX] Controls.Add: 免注册表未命中且 CLSIDFromProgID 失败 progId=%ls (需注册表或 Object= 引用)\n", progId);
            return NULL;
        }
        hr = CoCreateInstance(&clsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
                              &IID_IUnknown, (void**)&pUnk);
        if (FAILED(hr) || !pUnk) {
            if (ocxTrace)
                fwprintf(stderr, L"[C3_OCX] Controls.Add registry fallback progId=%ls hr=0x%08lX\n", progId, (unsigned long)hr);
            return NULL;
        }
        if (ocxTrace)
            fwprintf(stderr, L"[C3_OCX] Controls.Add registry fallback progId=%ls pUnk=%p\n", progId, (void*)pUnk);
    }

    /* 3. Get IDispatch */
    IDispatch* pDisp = NULL;
    hr = pUnk->lpVtbl->QueryInterface(pUnk, &IID_IDispatch, (void**)&pDisp);
    if (FAILED(hr) || !pDisp) {
        pUnk->lpVtbl->Release(pUnk);
        return NULL;
    }

    /* 4. Get IOleObject */
    IOleObject* pOleObj = NULL;
    hr = pDisp->lpVtbl->QueryInterface(pDisp, &IID_IOleObject, (void**)&pOleObj);
    if (FAILED(hr) || !pOleObj) {
        /* Non-ActiveX COM object: just return IDispatch (no embedding) */
        pUnk->lpVtbl->Release(pUnk);
        return (void*)pDisp;
    }

    /* Fix <vbeclipse> 2026-10-06: 运行期 Controls.Add 原先只有 SetClientSite +
     * 写死的 extent/rect (26700x20000 HIMETRIC ≈ 1009x756px, 注释里的
     * "400x300px≈267x200HM" 换算还是错的), 不 DoVerb 激活、不登记转发/绘制 ——
     * windowless 控件运行期加上去就是一片空白。现在走与 vb6_OcxHost_Create
     * (设计期 OCX) 同一套完整宿主流程, 初始几何 = 窗体客户区。 */

    /* 5. 完整 site (含 windowless 三件套 + Dispatch) */
    Vb6AxSite* site = (Vb6AxSite*)calloc(1, sizeof(Vb6AxSite));
    if (!site) { pOleObj->lpVtbl->Release(pOleObj); pUnk->lpVtbl->Release(pUnk); return (void*)pDisp; }
    site->lpVtblClientSite = &g_axSiteClientSiteVtbl;
    site->lpVtblInPlaceSite = &g_axSiteInPlaceSiteVtbl;
    site->lpVtblInPlaceFrame = &g_axSiteInPlaceFrameVtbl;
    site->lpVtblControlSite = &g_axSiteControlSiteVtbl;
    site->lpVtblInPlaceSiteWindowless = &g_axSiteInPlaceSiteWindowlessVtbl;
    site->lpVtblDispatch = &g_axSiteDispatchVtbl;
    site->ref = 1;
    site->hwndForm = (HWND)hwnd;
    site->pOleObj = pOleObj;
    if (ctrlName) { wcsncpy(site->ctrlName, ctrlName, 127); site->ctrlName[127] = 0; }
    vb6_AxTypeNameFromProgId(progId, site->typeName, 128);  /* 控件类名 —— 见上面的注释 */
    RECT rcClient;
    GetClientRect((HWND)hwnd, &rcClient);
    site->rcCtrl = rcClient;
    axSiteRegister(site);

    /* 6. SetClientSite */
    pOleObj->lpVtbl->SetClientSite(pOleObj, (IOleClientSite*)&site->lpVtblClientSite);

    /* 7. 初始 extent = 窗体客户区 (HIMETRIC; px→twips→HM, 与设计期同一换算) */
    SIZEL sz = { vb6_twipsToHimetric((rcClient.right - rcClient.left) * 15),
                 vb6_twipsToHimetric((rcClient.bottom - rcClient.top) * 15) };
    pOleObj->lpVtbl->SetExtent(pOleObj, DVASPECT_CONTENT, &sz);

    /* 8. 原地激活 (与设计期同口径: 只 INPLACEACTIVATE, UIACTIVATE 需要完整的
     * IOleInPlaceUIWindow/菜单合并, VB6 UserControl 会挂起) */
    hr = pOleObj->lpVtbl->DoVerb(pOleObj, OLEIVERB_INPLACEACTIVATE, NULL,
                                 (IOleClientSite*)&site->lpVtblClientSite,
                                 -1, (HWND)hwnd, NULL);
    if (FAILED(hr)) {
        pOleObj->lpVtbl->DoVerb(pOleObj, OLEIVERB_SHOW, NULL,
                                (IOleClientSite*)&site->lpVtblClientSite,
                                -1, (HWND)hwnd, NULL);
    }

    /* 9. SetObjectRects = 窗体客户区 (不是写死的 400x300) */
    IOleInPlaceObject* pIPO = NULL;
    hr = pOleObj->lpVtbl->QueryInterface(pOleObj, &IID_IOleInPlaceObject, (void**)&pIPO);
    if (SUCCEEDED(hr) && pIPO) {
        pIPO->lpVtbl->SetObjectRects(pIPO, &rcClient, &rcClient);
        pIPO->lpVtbl->Release(pIPO);
    }

    /* 10. windowless 消息转发 + IViewObject 宿主绘制 (与设计期同: 从登记表回找 site) */
    IOleInPlaceObjectWindowless* pIPOW = NULL;
    if (SUCCEEDED(pDisp->lpVtbl->QueryInterface(pDisp, &IID_IOleInPlaceObjectWindowless,
                                                (void**)&pIPOW)) && pIPOW) {
        for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
            if (g_axSites[i] == site) { g_axSites[i]->pInPlaceObj = pIPOW; break; }
        }
    }
    IViewObject* pView = NULL;
    if (SUCCEEDED(pDisp->lpVtbl->QueryInterface(pDisp, &IID_IViewObject, (void**)&pView)) && pView) {
        for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
            if (g_axSites[i] == site) { g_axSites[i]->pViewObj = pView; break; }
        }
    }

    InvalidateRect((HWND)hwnd, NULL, TRUE);

    pUnk->lpVtbl->Release(pUnk);
    return (void*)pDisp;
}

/* 把窗体消息转发给该窗体上所有无窗口 OCX 控件.
 * 返回 1 = 至少一个控件处理了该消息 (调用方应把 *plResult 作为结果).
 * windowless 控件不接收窗口消息: 绘制 (WM_PAINT)、鼠标、键盘全靠宿主转发,
 * 不转发则控件永不刷新 — 表现为窗体上一片空白. */
int32_t vb6_OcxHost_ForwardMessage(void* hwndForm, unsigned int msg, uintptr_t wp,
                                   intptr_t lp, intptr_t* plResult) {
    int handled = 0;
    for (int i = 0; i < vb6_axSiteCount; i++) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->hwndForm != (HWND)hwndForm || !s->pInPlaceObj) continue;
        if (s->hwndCtrl && IsWindow(s->hwndCtrl)) continue;   /* 窗口化控件不走宿主转发 */
        if (!s->visible) continue;   /* 容器 Extender 隐藏了它: 不再收消息 */
        IOleInPlaceObjectWindowless* p = (IOleInPlaceObjectWindowless*)s->pInPlaceObj;
        LRESULT lr = 0;
        HRESULT hr = p->lpVtbl->OnWindowMessage(p, (UINT)msg, (WPARAM)wp, (LPARAM)lp, &lr);
        if (hr == S_OK) {
            if (plResult) *plResult = (intptr_t)lr;
            handled = 1;
        }
    }
    return handled;
}

/* 主动让该窗体上的无窗口控件重绘 (窗体首次呈现 / 尺寸变化后) */
void vb6_OcxHost_InvalidateAll(void* hwndForm) {
    for (int i = 0; i < vb6_axSiteCount; i++) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->hwndForm != (HWND)hwndForm) continue;
        InvalidateRect(s->hwndForm, &s->rcCtrl, TRUE);
    }
}

/* Fix 143d: 把该窗体上所有 OCX 控件画到给定 DC (在 WM_PAINT 的
 * BeginPaint/EndPaint 之间调用). 用 IViewObject::Draw — 这是窗口化与
 * 无窗口控件都支持的通用绘制入口, 比转发 WM_PAINT 更可靠.
 * 注: windowless 控件的绘制矩形必须换算成 DC 坐标 (DC 原点=窗体客户区左上),
 * 我们的 rcCtrl 本身就是客户区坐标, 直接用. */
void vb6_OcxHost_PaintAll(void* hwndForm, void* hdc) {
    HDC dc = (HDC)hdc;
    if (!dc) return;
    const int paintTrace = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0);
    for (int i = 0; i < vb6_axSiteCount; i++) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->hwndForm != (HWND)hwndForm || !s->pViewObj) continue;
        /* 容器 Extender 隐藏了它: 不再画 (窗口化控件由 ShowWindow 盖住, 这一条只管
           windowless —— 它们的"隐藏"没有任何窗口状态可依, 只能靠这份标记)。 */
        if (!s->visible) continue;
        /* 只画真正的无窗口控件! 有自己窗口的控件 (VB6 UserControl 绝大多数是
         * windowed, 窗口类 ThunderRT6UserControlDC) 会自己绘制; 宿主再对它调
         * IViewObject::Draw 会与控件自绘冲突, 实测点击按钮触发重绘时
         * 异常从 WndProc 逃逸 → 进程被 user32 杀掉 (0xC000041D). */
        if (!s->pInPlaceObj) continue;
        if (s->hwndCtrl && IsWindow(s->hwndCtrl)) continue;   /* 窗口化控件自绘 */
        IViewObject* pView = (IViewObject*)s->pViewObj;
        RECTL rc;
        rc.left   = s->rcCtrl.left;
        rc.top    = s->rcCtrl.top;
        rc.right  = s->rcCtrl.right;
        rc.bottom = s->rcCtrl.bottom;
        int saved = SaveDC(dc);
        IntersectClipRect(dc, (int)rc.left, (int)rc.top, (int)rc.right, (int)rc.bottom);
        HRESULT dhr = pView->lpVtbl->Draw(pView, DVASPECT_CONTENT, -1, NULL, NULL, NULL, dc, &rc, NULL, NULL, 0);
        if (paintTrace) {
            fprintf(stderr, "[C3_OCX] PaintAll ctrl='%ls' rect=(%ld,%ld)-(%ld,%ld) hr=0x%08lX\n",
                    s->ctrlName, (long)rc.left, (long)rc.top, (long)rc.right, (long)rc.bottom,
                    (unsigned long)dhr);
            fflush(stderr);
        }
        RestoreDC(dc, saved);
    }
}

/* ============================================================
 * 账 (Outlook2013 主区 Report 盖住 Calendar): 容器 Extender 的三位主人。
 * VB6 里 `控件.Visible / .Left/.Top/.Width/.Height / .Move` 都不是控件自己的属性,
 * 是 VB6 窗体 (容器) 的 Extender 账 —— 它们改的是控件**窗口**的可见性与矩形。
 * C3 生成码把这些直接发给控件对象 (vb6_hwnd_<name> = 控件 IDispatch*), 而 RTL 的
 * vb6_GetControlVisible / vb6_ControlMove 一族按"原生控件"口径收 HWND, IsWindow()
 * 对 IDispatch* 判假 ⇒ 静默不动作 (实测: `wndReportControl.Visible = False` 无声
 * 失败, Report 一直盖在 Calendar 上; Form_Resize 里两处 Move 也全部落空)。
 * 下面按控件对象反查站点, 交出它激活后那个窗口; 找不到就是 NULL (旧行为不变)。
 * ============================================================ */
/* 账 (Outlook2013 frmInbox cmdSearch 的 buddy 贴靠): 容器 Extender 只带 (窗体, 实例名),
 * 要读回几何/Move 就得按名反查站点交出控件对象 —— 名字与站点账是同一份, 不另存。 */
void* vb6_OcxHost_DispByName(void* hwndForm, const wchar_t* name) {
    if (!hwndForm || !name || !name[0]) return NULL;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (s && s->pDisp && s->hwndForm == (HWND)hwndForm
            && s->ctrlName[0] && _wcsicmp(s->ctrlName, name) == 0)
            return s->pDisp;
    }
    return NULL;
}

void* vb6_OcxHost_WindowOf(void* ctlDisp) {
    if (!ctlDisp) return NULL;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (s && s->pDisp == ctlDisp && s->hwndCtrl && IsWindow(s->hwndCtrl))
            return (void*)s->hwndCtrl;
    }
    return NULL;
}

/* 账 (Outlook2013 左面板 VisualTheme): vb6_OcxHost_WindowOf 的**反向** —— 由窗口
 * 回找宿主托管的控件 IDispatch。
 * 场景: `For Each Ctrl In Form.Controls: Set CtrlCaption = Ctrl: CtrlCaption.VisualTheme = v`
 * 里 Controls 枚举交出的是控件窗口 (Codejock 还会带内部子窗口), 晚绑定写属性于是落到
 * 宿主模型 (它只认窗口层 Extender 成员) ⇒ 控件自己的属性被静默丢弃。
 * 判据: 自下而上找最近的、等于某站点 hwndCtrl 的祖先; 另兼容 hwndCtrl 是内层面
 * (Codejock 外面套 AfxWnd/AfxOleControl) 的情形。窗体本身不算 (返回 NULL)。 */
void* vb6_OcxHost_DispatchOf(void* hwndOrCtl) {
    if (!hwndOrCtl) return NULL;
    /* 1) 已经是控件对象 (windowless 站点) */
    for (int i = vb6_axSiteCount - 1; i >= 0; i--)
        if (g_axSites[i] && g_axSites[i]->pDisp == hwndOrCtl) return g_axSites[i]->pDisp;

    HWND h = (HWND)hwndOrCtl;
    if (!IsWindow(h)) return NULL;

    /* 2) 自下而上: 最近的控件窗口 (含控件内部子窗口) */
    for (HWND w = h; w; w = GetParent(w)) {
        for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
            Vb6AxSite* s = g_axSites[i];
            if (s && s->pDisp && s->hwndCtrl && s->hwndCtrl == w) return s->pDisp;
        }
    }

    /* 3) 容器窗口: hwndCtrl 是它的内层子窗口 */
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || !s->pDisp || !s->hwndCtrl || !IsWindow(s->hwndCtrl)) continue;
        if (h == s->hwndForm) continue;   /* 窗体本身不是控件 */
        for (HWND w = GetParent(s->hwndCtrl); w; w = GetParent(w))
            if (w == h) return s->pDisp;
    }
    return NULL;
}

/* 账 (Outlook2013 左面板 VisualTheme): VB6 的 `Form.Controls` 包含**所有**控件,
 * 无窗口 (windowless) 的 OCX 也在其中; C3 的集合只按窗口枚举, 于是
 * `For Each Ctrl In Form.Controls: Ctrl.VisualTheme = ...` 永远遍历不到这些控件。
 * 这里把该窗体上无窗口的站点控件对象补进集合 (有窗的已由窗口枚举覆盖)。 */
int32_t vb6_OcxHost_CollectWindowless(void* hwndForm, void** out, int32_t max) {
    int32_t n = 0;
    if (!hwndForm || !out) return 0;
    for (int i = 0; i < vb6_axSiteCount && n < max; i++) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || !s->pDisp || s->hwndForm != (HWND)hwndForm) continue;
        if (s->hwndCtrl && IsWindow(s->hwndCtrl)) continue;   /* 有窗口: 已覆盖 */
        int dup = 0;
        for (int32_t j = 0; j < n; j++) if (out[j] == s->pDisp) { dup = 1; break; }
        if (!dup) {
            out[n++] = s->pDisp;
            if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
                fwprintf(stderr, L"[C3_OCX] CollectWindowless form=%p ctl=%p name=\"%ls\"\n",
                         (void*)hwndForm, s->pDisp, s->ctrlName ? s->ctrlName : L"?");
        }
    }
    return n;
}

/* 控件类名 (站点创建时按 CLSID 的注册 ProgID 记下)。返回 1 = 这枚对象确是本宿主
 * 托管的 OCX 控件且类名已知; 其余一律 0 (调用方据此不改判)。 */
int32_t vb6_OcxHost_TypeNameOf(void* hwndOrCtl, wchar_t* out, int32_t cch) {
    if (!hwndOrCtl || !out || cch <= 0) return 0;
    /* 1) 宿主托管的 OCX: 对象本身或其窗口 (含内部子窗口 / 外层容器窗口) 都算 */
    void* ctl = vb6_OcxHost_DispatchOf(hwndOrCtl);
    if (ctl) {
        for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
            Vb6AxSite* s = g_axSites[i];
            if (!s || !s->typeName[0] || s->pDisp != ctl) continue;
            wcsncpy(out, s->typeName, (size_t)(cch - 1));
            out[cch - 1] = 0;
            return 1;
        }
    }
    /* 2) VB6 原生控件/窗体: 宿主对象登记表里有 `.frm` 声明的类型名
     *    (`vb6_HostObj_Register(hwnd, name, "XtremeSuiteControls.TreeView")` 尾段) */
    {
        const wchar_t* nm = vb6_Host_TypeNameOf(hwndOrCtl);
        if (nm && *nm) {
            wcsncpy(out, nm, (size_t)(cch - 1));
            out[cch - 1] = 0;
            return 1;
        }
    }
    return 0;
}

/* 读/写容器 Extender 的 Visible。返回 1 = 这枚对象确实是本宿主托管的 OCX 控件
 * (调用方据此决定要不要走普通 IDispatch 路)。windowless 控件没有窗口状态,
 * 一律答站点里那份标记。 */
int32_t vb6_OcxHost_GetVisible(void* ctlDisp, int32_t* outVisible) {
    if (!ctlDisp) return 0;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->pDisp != ctlDisp) continue;
        if (outVisible) {
            if (s->hwndCtrl && IsWindow(s->hwndCtrl))
                *outVisible = IsWindowVisible(s->hwndCtrl) ? 1 : 0;
            else
                *outVisible = s->visible;
        }
        return 1;
    }
    return 0;
}

int32_t vb6_OcxHost_SetVisible(void* ctlDisp, int32_t visible) {
    if (!ctlDisp) return 0;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->pDisp != ctlDisp) continue;
        s->visible = visible ? 1 : 0;
        if (s->hwndCtrl && IsWindow(s->hwndCtrl)) {
            ShowWindow(s->hwndCtrl, s->visible ? SW_SHOW : SW_HIDE);
            if (s->visible && s->hwndForm) InvalidateRect(s->hwndForm, NULL, TRUE);
        } else if (s->hwndForm) {
            InvalidateRect(s->hwndForm, &s->rcCtrl, TRUE);
        }
        return 1;
    }
    return 0;
}

/* 账 (Outlook2013 日历只铺了 561px): 容器只 SetWindowPos 控件窗口还不够 ——
 * 原地激活的 OCX 会按 IOleInPlaceObject::SetObjectRects 给的新矩形重排自己的
 * 客户区。VB6 的窗体容器在 Move 之后就是走这一发 (实测 ref 里 XTPCalendar
 * 铺满主区, 而 C3 里内层 AfxWnd 一直停在设计期 8415 缇 = 561px)。
 * 这里按控件窗口**当前**的矩形补发这一发; 没有窗口 (windowless) 或窗口没变
 * 就什么都不做。返回 1 = 是本宿主托管的 OCX 控件。 */
int32_t vb6_OcxHost_SyncRects(void* ctlDisp) {
    if (!ctlDisp) return 0;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->pDisp != ctlDisp) continue;
        if (!s->hwndCtrl || !IsWindow(s->hwndCtrl)) return 1;
        RECT rc;
        GetWindowRect(s->hwndCtrl, &rc);
        POINT p0 = { rc.left, rc.top }, p1 = { rc.right, rc.bottom };
        ScreenToClient(s->hwndForm, &p0);
        ScreenToClient(s->hwndForm, &p1);
        RECT r = { p0.x, p0.y, p1.x, p1.y };
        s->rcCtrl = r;
        IOleInPlaceObject* pIPO = NULL;
        IDispatch* pD = (IDispatch*)s->pDisp;
        if (SUCCEEDED(pD->lpVtbl->QueryInterface(pD, &IID_IOleInPlaceObject, (void**)&pIPO)) && pIPO) {
            pIPO->lpVtbl->SetObjectRects(pIPO, &r, &r);
            pIPO->lpVtbl->Release(pIPO);
        }
        return 1;
    }
    return 0;
}

/* 账 #305 第三百五十三刀: 对本窗体上每个 OCX 站点尝试晚绑定调用一个无参方法
 * (GUISample: CommandBars.RecalcLayout —— 它的停靠带宽度由这个重算, maximize
 * 后窗格/菜单带不回流就是没人调它)。调用失败的静默 (绝大多数控件没有该方法)。 */
void vb6_OcxHost_InvokeAll(void* hwndForm, const wchar_t* method) {
    if (!hwndForm || !method || !method[0]) return;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->hwndForm != (HWND)hwndForm || !s->pDisp) continue;
        IDispatch* pD = (IDispatch*)s->pDisp;
        DISPID dispid = DISPID_UNKNOWN;
        OLECHAR* nm = (OLECHAR*)method;
        if (FAILED(pD->lpVtbl->GetIDsOfNames(pD, &IID_NULL, &nm, 1,
                       GetUserDefaultLCID(), &dispid)) || dispid == DISPID_UNKNOWN)
            continue;
        if (GetEnvironmentVariableA("C3_CRASH_TRACE", NULL, 0) > 0) {
            fprintf(stderr, "[C3_INVOK] site=%ls method=%ls dispid=%ld\n",
                    s->ctrlName, method, (long)dispid); fflush(stderr);
        }
        DISPPARAMS dp; memset(&dp, 0, sizeof(dp));
        VARIANT res; VariantInit(&res);
        pD->lpVtbl->Invoke(pD, dispid, &IID_NULL, GetUserDefaultLCID(),
            DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, &res, NULL, NULL);
        VariantClear(&res);
    }
}

/* 账 #305 第三百四十八刀: 见 vb6forms_controls.h 的声明 —— MDIForm 尺寸变化
 * 后给所有 windowless 站点补发 SetObjectRects(全客户区)。有窗口的控件不归这里
 * (它们有自己的设计矩形, 由 vb6_ControlMove 那条路管)。 */
void vb6_OcxHost_FormResized(void* hwndForm, int w, int h) {
    int touched = 0;
    int dbg348 = GetEnvironmentVariableA("C3_CRASH_TRACE", NULL, 0) > 0;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->hwndForm != (HWND)hwndForm) continue;
        if (s->hwndCtrl && IsWindow(s->hwndCtrl)) continue;
        if (dbg348) {
            fprintf(stderr, "[C3_FR] site=%ls form=%p rect=(%ld,%ld)-(%ld,%ld)\n",
                    s->ctrlName, (void*)s->hwndForm,
                    (long)s->rcCtrl.left, (long)s->rcCtrl.top,
                    (long)s->rcCtrl.right, (long)s->rcCtrl.bottom);
            fflush(stderr);
        }
        RECT r = {0, 0, w, h};
        s->rcCtrl = r;
        if (s->pInPlaceObj) {
            IOleInPlaceObjectWindowless* p = (IOleInPlaceObjectWindowless*)s->pInPlaceObj;
            p->lpVtbl->SetObjectRects(p, &r, &r);
            touched = 1;
        } else if (s->pOleObj) {
            IOleInPlaceObject* pIPO = NULL;
            if (SUCCEEDED(s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IOleInPlaceObject, (void**)&pIPO)) && pIPO) {
                pIPO->lpVtbl->SetObjectRects(pIPO, &r, &r);
                pIPO->lpVtbl->Release(pIPO);
                touched = 1;
            }
        }
    }
    if (touched) InvalidateRect((HWND)hwndForm, NULL, TRUE);
}

/* 账 (Outlook2013 左面板 My Calendars 树跑到顶端): **windowless** 控件没有窗口,
 * 但 VB6 的容器照样按 Extender 存着它的 Left/Top/Width/Height (frmPaneCalendar
 * 的 UpdateLayout 就读 ItemCaption.top/Height 来摆 tree)。站点里那份 rcCtrl 就是
 * 容器的账本 (像素, 相对 hwndForm 客户区), 下面两位给"无窗口"那条路用:
 *   GetRect → out[4] = x,y,w,h; *outForm = 容器窗体 (ScaleMode 用); 1 = 命中;
 *   SetRect → 按 mask 改账本 (位 1/2/4/8 同 vb6_ControlMove), 重画新旧并集。 */
int32_t vb6_OcxHost_GetRect(void* ctlDisp, int32_t out[4], void** outForm) {
    if (!ctlDisp) return 0;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->pDisp != ctlDisp) continue;
        if (s->hwndCtrl && IsWindow(s->hwndCtrl)) return 0;   /* 有窗口: 走窗口那条 */
        if (out) {
            out[0] = s->rcCtrl.left; out[1] = s->rcCtrl.top;
            out[2] = s->rcCtrl.right - s->rcCtrl.left;
            out[3] = s->rcCtrl.bottom - s->rcCtrl.top;
        }
        if (outForm) *outForm = (void*)s->hwndForm;
        return 1;
    }
    return 0;
}

/* windowless 控件没有窗口 ⇒ 没有 GeomCache 那一格，几何只能按站点账本折给用户。
 * 四个 getter 各折一遍就是四本账，收成这一处；账本住在这里，折算就跟着账本走。
 * slot 即 VB6_GEOM_* (0..3)，slot & 1 = 竖直那一轴，与 vb6_GeomCacheRead 同口径。 */
int32_t vb6_OcxHost_GetUserCoord(void* ctlDisp, int32_t slot) {
    int32_t r[4]; void* form = NULL;
    if (!vb6_OcxHost_GetRect(ctlDisp, r, &form)) return 0;
    int32_t cm = vb6_ContainerScaleMode(form);
    int32_t u = (int32_t)vb6_ScalePxToUser((double)r[slot], cm, (int)(slot & 1));
    if (GetEnvironmentVariableW(L"C3_GEOM_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_GEOM] GetUserCoord ctl=%p slot=%d px=%d cm=%d user=%d\n",
                ctlDisp, (int)slot, (int)r[slot], (int)cm, (int)u);
    return u;
}

int32_t vb6_OcxHost_SetRect(void* ctlDisp, int32_t x, int32_t y, int32_t w, int32_t h, int32_t mask) {
    if (!ctlDisp) return 0;
    for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
        Vb6AxSite* s = g_axSites[i];
        if (!s || s->pDisp != ctlDisp) continue;
        if (s->hwndCtrl && IsWindow(s->hwndCtrl)) return 0;   /* 有窗口: 走窗口那条 */
        if (GetEnvironmentVariableW(L"C3_GEOM_TRACE", NULL, 0) > 0) {
            fprintf(stderr, "[C3_GEOM] SetRect ctrl='%ls' mask=%d x=%d y=%d w=%d h=%d old=(%ld,%ld,%ld,%ld)\n",
                    s->ctrlName, (int)mask, (int)x, (int)y, (int)w, (int)h,
                    (long)s->rcCtrl.left, (long)s->rcCtrl.top,
                    (long)s->rcCtrl.right, (long)s->rcCtrl.bottom);
            fflush(stderr);
        }
        RECT old = s->rcCtrl, nw = s->rcCtrl;
        /* 宽高先存下来: 下面先写 left/top, 再用 left/top 算右/下边 —— 若现场取
           (nw.right - nw.left), 拿到的已经是改过之后的值, 每挪一次宽高就被吃掉
           一个偏移量 (实测 MainCaption 19px → 17 → -144 → -167, DatePicker 宽
           被算成 5385 缇)。 */
        int w0 = nw.right - nw.left, h0 = nw.bottom - nw.top;
        if (mask & 1) { nw.left = x; nw.right = nw.left + w0; }
        if (mask & 2) { nw.top = y;  nw.bottom = nw.top + h0; }
        if (mask & 4) { nw.right = nw.left + w; }
        if (mask & 8) { nw.bottom = nw.top + h; }
        s->rcCtrl = nw;
        if (s->hwndForm) {
            RECT u = old;
            if (nw.left < u.left) u.left = nw.left;
            if (nw.top < u.top) u.top = nw.top;
            if (nw.right > u.right) u.right = nw.right;
            if (nw.bottom > u.bottom) u.bottom = nw.bottom;
            InvalidateRect(s->hwndForm, &u, TRUE);
        }
        return 1;
    }
    return 0;
}

/* A2 诊断: 激活各步之后枚举窗体直接子窗, 定位 AfxOleControl 包裹层出现的时机 */
static void axDiagChildren(const char* step, HWND form) {
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) <= 0) return;
    int n = 0;
    for (HWND cw = FindWindowExW(form, NULL, NULL, NULL); cw;
         cw = FindWindowExW(form, cw, NULL, NULL)) {
        wchar_t cls[128] = {0};
        RECT rr;
        GetClassNameW(cw, cls, 128);
        GetWindowRect(cw, &rr);
        fprintf(stderr, "[C3_STEP] %s child%d '%ls' (%ld,%ld)\n", step, n++, cls,
                (long)rr.left, (long)rr.top);
    }
    fprintf(stderr, "[C3_STEP] %s total=%d\n", step, n);
    fflush(stderr);
}

void* vb6_OcxHost_Create(void* hwndForm, const wchar_t* clsidStr, const wchar_t* altClsidStr,
                         const wchar_t* ocxPath, int x, int y, int w, int h,
                         const wchar_t* ctrlName, const Vb6OcxProp* props, int propCount) {
    if (!hwndForm || !clsidStr) return NULL;

    /* C3_OCX_TRACE=1: 打印实例化/激活各步 HRESULT (诊断宿主问题) */
    int ocxTrace = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0);

    /* 1. 实例化 (多路径 + 多候选 CLSID)
     * 候选 1 = clsidStr: 来自 OCX typelib 的 coclass GUID (typelib 真实类, 最可靠)
     * 候选 2 = altClsidStr: 来自 vbp Object= 的 GUID (可能指向旧版本/已注册副本)
     * 两者都可能对也可能过期 — 依次尝试, 谁先成功用谁. */
    IUnknown* pUnk = NULL;
    HRESULT hr = ((HRESULT)0x800401F1L);
    const wchar_t* usedClsid = NULL;
    {
        const wchar_t* cand[2];
        int nCand = 0;
        cand[nCand++] = clsidStr;
        if (altClsidStr && *altClsidStr && wcscmp(altClsidStr, clsidStr) != 0)
            cand[nCand++] = altClsidStr;
        for (int ci = 0; ci < nCand && !pUnk; ci++) {
            CLSID c;
            if (FAILED(CLSIDFromString((LPOLESTR)cand[ci], &c))) continue;
            hr = vb6_ocxCreateAny(ocxPath, &c, (void**)&pUnk);
            if (pUnk) usedClsid = cand[ci];
            if ((FAILED(hr) || !pUnk) && nCand > 1 && ci + 1 < nCand) {
                /* 第一个候选失败 → 试下一个 */
                if (ocxTrace) fprintf(stderr, "[C3_OCX]   clsid[%d] failed hr=0x%08lX, trying next\n",
                                      ci, (unsigned long)hr);
            }
        }
    }
    if (ocxTrace) fprintf(stderr, "[C3_OCX] %ls instantiate hr=0x%08lX pUnk=%p\n",
                          ctrlName ? ctrlName : L"?", (unsigned long)hr, (void*)pUnk);
    if (FAILED(hr) || !pUnk) return NULL;

    /* VB6 在 CoCreateInstance 之后必调 OleSetContainedObject(TRUE): 告知服务器
     * "你被某个容器装进来了" (服务器据此决定是否走上层激活/通知路径).
     * A2: 补齐这步 (缺失时服务器的激活判断与本机不同). */
    {
        HRESULT hrCont = OleSetContainedObject((IUnknown*)pUnk, TRUE);
        if (ocxTrace) fprintf(stderr, "[C3_OCX]   OleSetContainedObject hr=0x%08lX\n",
                              (unsigned long)hrCont);
    }

    /* 2. IDispatch (返回值, vb6_hwnd_<name> 语义升级为"控件对象") */
    IDispatch* pDisp = NULL;
    {
        HRESULT hrQI = pUnk->lpVtbl->QueryInterface(pUnk, &IID_IDispatch, (void**)&pDisp);
        if (ocxTrace) fprintf(stderr, "[C3_OCX]   QI(IDispatch) hr=0x%08lX pDisp=%p\n",
                              (unsigned long)hrQI, (void*)pDisp);
        if (FAILED(hrQI) || !pDisp) {
            pUnk->lpVtbl->Release(pUnk);
            return NULL;
        }
    }

    /* 3. IOleObject + site */
    IOleObject* pOleObj = NULL;
    {
        HRESULT hrOle = pDisp->lpVtbl->QueryInterface(pDisp, &IID_IOleObject, (void**)&pOleObj);
        if (ocxTrace) fprintf(stderr, "[C3_OCX]   QI(IOleObject) hr=0x%08lX p=%p\n",
                              (unsigned long)hrOle, (void*)pOleObj);
    }
    if (pOleObj) {
        Vb6AxSite* site = (Vb6AxSite*)calloc(1, sizeof(Vb6AxSite));
        IOleClientSite* siteCS = NULL;
        if (site) {
            site->lpVtblClientSite = &g_axSiteClientSiteVtbl;
            site->lpVtblInPlaceSite = &g_axSiteInPlaceSiteVtbl;
            site->lpVtblInPlaceFrame = &g_axSiteInPlaceFrameVtbl;
            /* Fix 143d: windowless 控件宿主三件套 */
            site->lpVtblControlSite = &g_axSiteControlSiteVtbl;
            site->lpVtblInPlaceSiteWindowless = &g_axSiteInPlaceSiteWindowlessVtbl;
            site->lpVtblDispatch = &g_axSiteDispatchVtbl;
            site->lpVtblAdviseSink = &g_axSiteAdviseSinkVtbl;
            site->ref = 1;
            site->hwndForm = (HWND)hwndForm;
            site->pOleObj = pOleObj;
            /* 容器 Extender 的账: 按控件对象 (生成码里的 vb6_hwnd_<name>) 反查站点 */
            site->pDisp = (void*)pDisp;
            site->visible = 1;
            if (ctrlName) { wcsncpy(site->ctrlName, ctrlName, 127); site->ctrlName[127] = 0; }
            vb6_AxTypeNameFromClsid(usedClsid, site->typeName, 128);
            site->rcCtrl.left = x; site->rcCtrl.top = y;
            site->rcCtrl.right = x + w; site->rcCtrl.bottom = y + h;
            siteCS = (IOleClientSite*)&site->lpVtblClientSite;
            axSiteRegister(site);   /* Fix 143d: 登记以供窗体消息转发 */
            if (ocxTrace) fprintf(stderr, "[C3_OCX]   SetClientSite begin\n");
            pOleObj->lpVtbl->SetClientSite(pOleObj, siteCS);
            if (ocxTrace) fprintf(stderr, "[C3_OCX]   SetClientSite done\n");
            /* VB6 还会调 IOleObject::SetHostNames(容器应用名, 控件名) —— 服务器用
             * 它给自建窗口命名/区分多个同类实例. A2: 补齐. */
            {
                wchar_t hostApp[160] = {0};
                GetWindowTextW((HWND)hwndForm, hostApp, 159);
                HRESULT hrHN = pOleObj->lpVtbl->SetHostNames(
                    pOleObj, hostApp, ctrlName ? ctrlName : L"");
                if (ocxTrace) fprintf(stderr, "[C3_OCX]   SetHostNames('%ls','%ls') hr=0x%08lX\n",
                                      hostApp, ctrlName ? ctrlName : L"", (unsigned long)hrHN);
            }
            /* VB6 容器对每个控件都 IOleObject::Advise 一个 IAdviseSink
             * (OnViewChange / OnClose / OnShowWindow 由此送到容器). A2: 补齐. */
            {
                HRESULT hrAdv = pOleObj->lpVtbl->Advise(pOleObj,
                                    (IAdviseSink*)&site->lpVtblAdviseSink,
                                    &site->adviseCookie);
                if (ocxTrace) fprintf(stderr, "[C3_OCX]   Advise hr=0x%08lX cookie=%lu\n",
                                      (unsigned long)hrAdv, (unsigned long)site->adviseCookie);
            }
            if (ocxTrace && ctrlName && wcscmp(ctrlName, L"cmdSearch") == 0)
                axDiagChildren("aftersetclient", (HWND)hwndForm);
        }
        /* 4. 设计期大小 (HIMETRIC) */
        SIZEL sz = { vb6_twipsToHimetric(w * 15), vb6_twipsToHimetric(h * 15) };
        pOleObj->lpVtbl->SetExtent(pOleObj, DVASPECT_CONTENT, &sz);
        if (ocxTrace) fprintf(stderr, "[C3_OCX]   SetExtent done\n");
        /* VB6 在 SetExtent 之后还会 GetExtent 回读一次 (容器拿真实尺寸做布局). A2: 补齐. */
        {
            SIZEL szBack = {0, 0};
            HRESULT hrGE = pOleObj->lpVtbl->GetExtent(pOleObj, DVASPECT_CONTENT, &szBack);
            if (ocxTrace) fprintf(stderr, "[C3_OCX]   GetExtent hr=0x%08lX -> %ld,%ld (asked %ld,%ld)\n",
                                  (unsigned long)hrGE, (long)szBack.cx, (long)szBack.cy,
                                  (long)sz.cx, (long)sz.cy);
        }
        if (ocxTrace && ctrlName && wcscmp(ctrlName, L"cmdSearch") == 0)
            axDiagChildren("aftersetextent", (HWND)hwndForm);

        /* 4b. Fix 143c 踩坑记录: 不要对 VB6 UserControl 调
         * IPersistStreamInit::InitNew —— 实测 NewTab01.ocx 返回 0x800A9C68
         * (VB6 运行时错误), 且之后 DoVerb 也一并失败, 控件彻底不可用.
         * VB6 容器的初始化由下方 IPersistPropertyBag::Load 承担. */

        /* 5. 设计期属性 (IPersistPropertyBag) */
        if (props && propCount > 0) {
            IPersistPropertyBag* pPPB = NULL;
            if (SUCCEEDED(pDisp->lpVtbl->QueryInterface(pDisp, &IID_IPersistPropertyBag, (void**)&pPPB)) && pPPB) {
                Vb6PropBag* bag = (Vb6PropBag*)calloc(1, sizeof(Vb6PropBag));
                if (bag) {
                    bag->lpVtbl = &g_pbVtbl;
                    bag->ref = 1;
                    bag->props = props;
                    bag->count = propCount;
                    bag->extentX = vb6_twipsToHimetric(w * 15);
                    bag->extentY = vb6_twipsToHimetric(h * 15);
                    if (ocxTrace) {
                        fprintf(stderr, "[C3_OCX]   PropertyBag::Load begin (%d props):", propCount);
                        for (int pi = 0; pi < propCount; pi++)
                            fprintf(stderr, " %ls", props[pi].name);
                        fprintf(stderr, "\n");
                    }
                    pPPB->lpVtbl->Load(pPPB, (IPropertyBag*)bag, NULL);
                    if (ocxTrace) fprintf(stderr, "[C3_OCX]   PropertyBag::Load done\n");
                    if (ocxTrace && ctrlName && wcscmp(ctrlName, L"cmdSearch") == 0)
                        axDiagChildren("afterload", (HWND)hwndForm);
                    ((IPropertyBag*)bag)->lpVtbl->Release((IPropertyBag*)bag);
                }
                pPPB->lpVtbl->Release(pPPB);
            } else if (ocxTrace) {
                fprintf(stderr, "[C3_OCX]   no IPersistPropertyBag\n");
            }
        }

        /* 6. 原地激活 (控件由此创建自己的子窗口并渲染).
         * 注意: 只做 INPLACEACTIVATE. OLEIVERB_UIACTIVATE 会让 VB6 UserControl
         * 挂起 (实测 NewTab01.ocx: DoVerb 不返回) — 它需要容器实现完整的
         * IOleInPlaceUIWindow / 菜单合并 / IOleInPlaceActiveObject 服务. */
        /* VB6 在激活前就先把控件摆到设计期位置 (IOleInPlaceObject::SetObjectRects).
         * A2: 补齐激活前的这一刀 (以前只在激活后才设, 控件激活时算布局可能拿不到
         * 自己的矩形). */
        {
            IOleInPlaceObject* pIPO0 = NULL;
            if (SUCCEEDED(pDisp->lpVtbl->QueryInterface(pDisp, &IID_IOleInPlaceObject,
                                                        (void**)&pIPO0)) && pIPO0) {
                RECT rc0 = { x, y, x + w, y + h };
                pIPO0->lpVtbl->SetObjectRects(pIPO0, &rc0, &rc0);
                if (ocxTrace) fprintf(stderr, "[C3_OCX]   SetObjectRects(pre-activate) (%ld,%ld)-(%ld,%ld)\n",
                                      (long)rc0.left, (long)rc0.top, (long)rc0.right, (long)rc0.bottom);
                pIPO0->lpVtbl->Release(pIPO0);
            }
        }
        hr = pOleObj->lpVtbl->DoVerb(pOleObj, OLEIVERB_INPLACEACTIVATE, NULL, siteCS,
                                     -1, (HWND)hwndForm, NULL);
        if (ocxTrace) fprintf(stderr, "[C3_OCX] DoVerb(INPLACEACTIVATE) hr=0x%08lX ctrl='%ls'\n",
                              (unsigned long)hr, ctrlName ? ctrlName : L"");
        if (FAILED(hr)) {
            hr = pOleObj->lpVtbl->DoVerb(pOleObj, OLEIVERB_SHOW, NULL, siteCS,
                                         -1, (HWND)hwndForm, NULL);
            if (ocxTrace) fprintf(stderr, "[C3_OCX] DoVerb(SHOW) hr=0x%08lX ctrl='%ls'\n",
                                  (unsigned long)hr, ctrlName ? ctrlName : L"");
        }
        if (ocxTrace) {
            int nSub = 0;
            HWND cw = FindWindowExW((HWND)hwndForm, NULL, NULL, NULL);
            while (cw) {
                wchar_t cls[128] = {0}, txt[128] = {0};
                GetClassNameW(cw, cls, 128);
                GetWindowTextW(cw, txt, 128);
                RECT rr; GetWindowRect(cw, &rr);
                fprintf(stderr, "[C3_OCX]   child %d: class='%ls' text='%ls' rect=(%ld,%ld,%ld,%ld) vis=%d\n",
                        nSub, cls, txt, (long)rr.left, (long)rr.top,
                        (long)(rr.right - rr.left), (long)(rr.bottom - rr.top),
                        IsWindowVisible(cw) ? 1 : 0);
                nSub++;
                cw = FindWindowExW((HWND)hwndForm, cw, NULL, NULL);
            }
            fprintf(stderr, "[C3_OCX] form child windows after activate: %d\n", nSub);
        }
    }

    /* 7. 位置 */
    {
        IOleInPlaceObject* pIPO = NULL;
        if (SUCCEEDED(pDisp->lpVtbl->QueryInterface(pDisp, &IID_IOleInPlaceObject, (void**)&pIPO)) && pIPO) {
            RECT rc = { x, y, x + w, y + h };
            pIPO->lpVtbl->SetObjectRects(pIPO, &rc, &rc);
            /* 账 (Outlook2013): 记下控件窗口 —— 容器的 Visible / Move 都落在它身上。
               windowless 控件这里给 0 (或与窗体同址), 一律当"无窗口"处理。 */
            HWND hwndCtl = NULL;
            pIPO->lpVtbl->GetWindow(pIPO, &hwndCtl);
            if (ocxTrace) {
                wchar_t cls[128] = {0};
                if (hwndCtl) GetClassNameW(hwndCtl, cls, 128);
                fprintf(stderr, "[C3_OCX]   GetWindow(IPO) hwnd=%p class='%ls'\n", (void*)hwndCtl, cls);
                fflush(stderr);
            }
            if (hwndCtl && hwndCtl != (HWND)hwndForm) {
                for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
                    if (g_axSites[i] && g_axSites[i]->hwndForm == (HWND)hwndForm &&
                        g_axSites[i]->pDisp == (void*)pDisp) {
                        g_axSites[i]->hwndCtrl = hwndCtl;
                        break;
                    }
                }
            }
            pIPO->lpVtbl->Release(pIPO);
        }
    }

    /* 8. Fix 143d: 保存 IOleInPlaceObjectWindowless (消息转发用) + 登记宿主.
     * windowless 控件激活后不会自己收窗口消息, 由宿主在窗体 WndProc 里
     * 调 vb6_OcxHost_ForwardMessage 转发 WM_PAINT/鼠标/键盘. */
    {
        IOleInPlaceObjectWindowless* pIPOW = NULL;
        if (SUCCEEDED(pDisp->lpVtbl->QueryInterface(pDisp, &IID_IOleInPlaceObjectWindowless,
                                                    (void**)&pIPOW)) && pIPOW) {
            /* 找一个属于本窗体的 site 记录 (site 在步骤 3 创建, 用 QI 反查不可靠,
             * 故这里按 hwndForm 从登记表里找最新一条) */
            for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
                if (g_axSites[i] && g_axSites[i]->hwndForm == (HWND)hwndForm) {
                    g_axSites[i]->pInPlaceObj = pIPOW;
                    break;
                }
            }
            if (ocxTrace) fprintf(stderr, "[C3_OCX]   IOleInPlaceObjectWindowless OK (windowless 控件)\n");
        } else if (ocxTrace) {
            fprintf(stderr, "[C3_OCX]   no IOleInPlaceObjectWindowless (windowed 控件)\n");
        }
    }

    /* 9. Fix 143d: 保存 IViewObject — 宿主 WM_PAINT 里用它绘制控件 */
    {
        IViewObject* pView = NULL;
        if (SUCCEEDED(pDisp->lpVtbl->QueryInterface(pDisp, &IID_IViewObject, (void**)&pView)) && pView) {
            for (int i = vb6_axSiteCount - 1; i >= 0; i--) {
                if (g_axSites[i] && g_axSites[i]->hwndForm == (HWND)hwndForm) {
                    g_axSites[i]->pViewObj = pView;
                    break;
                }
            }
            if (ocxTrace) fprintf(stderr, "[C3_OCX]   IViewObject OK (可用于宿主绘制)\n");
        } else if (ocxTrace) {
            fprintf(stderr, "[C3_OCX]   no IViewObject\n");
        }
    }

    /* 10. 请求首帧重绘 */
    InvalidateRect((HWND)hwndForm, NULL, TRUE);

    pUnk->lpVtbl->Release(pUnk);
    return (void*)pDisp;
}

#ifdef __cplusplus
} // extern "C"
#endif
