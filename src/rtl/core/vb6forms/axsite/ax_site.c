// ax_site.c - vb6forms_axsite 拆分片：站点本体: IOleClientSite / IOleInPlaceSite / IOleInPlaceFrame 三套 vtable + 已创建宿主登记表
//
// 内容 = 拆分前 vb6forms_axsite.c 第 29~33 / 72~206 / 523~531 行，纯搬移零重排无行为改动
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

/* ===== 简化 ActiveX 控件宿主 ===== */

/* Minimal IOleClientSite + IOleInPlaceSite + IOleInPlaceFrame
   用于 Controls.Add 动态加载 ActiveX 控件并嵌入 Form 窗口 */

/* --- IOleClientSite --- */

/* Fix 143d 诊断: C3_OCX_TRACE=1 时打印容器 QI 收到的 IID + 应答结果.
 * 用于判断控件到底在向宿主索要哪些接口 (Windowless/ControlSite/Ambient 等). */
static void ocxTraceQI(const char* iface, const wchar_t* ctrl, REFIID riid, HRESULT hr) {
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) <= 0) return;
    wchar_t guid[64] = {0};
    StringFromGUID2(riid, guid, 64);
    fprintf(stderr, "[C3_OCX]   QI %s %ls -> 0x%08lX ctrl='%ls'\n", iface, guid,
            (unsigned long)hr, ctrl ? ctrl : L"");
}

HRESULT STDMETHODCALLTYPE axSite_CS_QueryInterface(IOleClientSite* This, REFIID riid, void** ppv) {
    Vb6AxSite* s = SITE_OF(This, lpVtblClientSite);
    HRESULT r = E_NOINTERFACE;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IOleClientSite)) {
        *ppv = &s->lpVtblClientSite; s->ref++; r = S_OK;
    } else if (IsEqualIID(riid, &IID_IOleInPlaceSite)) {
        *ppv = &s->lpVtblInPlaceSite; s->ref++; r = S_OK;
    } else if (IsEqualIID(riid, &IID_IOleInPlaceFrame)) {
        *ppv = &s->lpVtblInPlaceFrame; s->ref++; r = S_OK;
    } else if (IsEqualIID(riid, &IID_IDispatch)) {
        /* Fix 143d: ambient 属性通道 (控件据此取 BackColor/Font/UserMode) */
        *ppv = &s->lpVtblDispatch; s->ref++; r = S_OK;
    } else if (IsEqualIID(riid, &IID_IOleControlSite)) {
        *ppv = &s->lpVtblControlSite; s->ref++; r = S_OK;
    } else if (IsEqualIID(riid, &IID_IOleInPlaceSiteWindowless)) {
        /* Fix 143d: 无窗口控件激活必需 */
        *ppv = &s->lpVtblInPlaceSiteWindowless; s->ref++; r = S_OK;
    } else if (IsEqualIID(riid, &IID_IOleInPlaceSiteEx)) {
        /* IOleInPlaceSiteEx 是 IOleInPlaceSiteWindowless 的前缀 (前 8 槽同表),
         * 直接交同一张表. 漏答这枚: MFC 控件 QI 失败后走旧式激活 → 强制窗口化
         * (实测 Codejock PushButton cmdSearch 多出一层 AfxOleControl140s 宿主窗口,
         * 自带主题边框的 1px 三边就画不出来). */
        *ppv = &s->lpVtblInPlaceSiteWindowless; s->ref++; r = S_OK;
    } else {
        *ppv = NULL;
    }
    ocxTraceQI("site", s->ctrlName, riid, r);
    return r;
}
ULONG STDMETHODCALLTYPE axSite_CS_AddRef(IOleClientSite* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblClientSite); return InterlockedIncrement(&s->ref);
}
ULONG STDMETHODCALLTYPE axSite_CS_Release(IOleClientSite* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblClientSite); LONG r = InterlockedDecrement(&s->ref);
    if (r <= 0) { free(s); } return r;
}
static HRESULT STDMETHODCALLTYPE axSite_CS_SaveObject(IOleClientSite* This) { return E_NOTIMPL; }
static HRESULT STDMETHODCALLTYPE axSite_CS_GetMoniker(IOleClientSite* This, DWORD a, DWORD b, IMoniker** ppM) { *ppM=NULL; return E_NOTIMPL; }
static IOleContainer* axContainerGet(HWND hwndForm);   /* 定义在下方 (§容器对象) */
static int axTrace(void);                              /* 定义在下方 (§容器对象) */
static HRESULT STDMETHODCALLTYPE axSite_CS_GetContainer(IOleClientSite* This, IOleContainer** ppC) {
    Vb6AxSite* s = SITE_OF(This, lpVtblClientSite);
    if (!ppC) return E_POINTER;
    /* Codejock PushButton 按名解析 BuddyControl 走这里 (§容器对象) */
    *ppC = axContainerGet(s->hwndForm);
    if (axTrace()) {
        fprintf(stderr, "[C3_OCX] GetContainer ctrl='%ls' -> %p\n", s->ctrlName, (void*)*ppC);
        fflush(stderr);
    }
    return *ppC ? S_OK : E_OUTOFMEMORY;
}
static HRESULT STDMETHODCALLTYPE axSite_CS_ShowObject(IOleClientSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_CS_OnShowWindow(IOleClientSite* This, BOOL f) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_CS_RequestNewObjectLayout(IOleClientSite* This) { return E_NOTIMPL; }

const IOleClientSiteVtbl g_axSiteClientSiteVtbl = {
    axSite_CS_QueryInterface, axSite_CS_AddRef, axSite_CS_Release,
    axSite_CS_SaveObject, axSite_CS_GetMoniker, axSite_CS_GetContainer,
    axSite_CS_ShowObject, axSite_CS_OnShowWindow, axSite_CS_RequestNewObjectLayout
};

/* --- IOleInPlaceSite --- */
static HRESULT STDMETHODCALLTYPE axSite_IPS_QueryInterface(IOleInPlaceSite* This, REFIID riid, void** ppv) {
    return axSite_CS_QueryInterface((IOleClientSite*)This, riid, ppv);
}
static ULONG STDMETHODCALLTYPE axSite_IPS_AddRef(IOleInPlaceSite* This) {
    return axSite_CS_AddRef((IOleClientSite*)SITE_OF(This, lpVtblInPlaceSite));
}
static ULONG STDMETHODCALLTYPE axSite_IPS_Release(IOleInPlaceSite* This) {
    return axSite_CS_Release((IOleClientSite*)SITE_OF(This, lpVtblInPlaceSite));
}
static HRESULT STDMETHODCALLTYPE axSite_IPS_GetWindow(IOleInPlaceSite* This, HWND* phwnd) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSite); *phwnd = s->hwndForm; return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPS_ContextSensitiveHelp(IOleInPlaceSite* This, BOOL f) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_CanInPlaceActivate(IOleInPlaceSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnInPlaceActivate(IOleInPlaceSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnUIActivate(IOleInPlaceSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_GetWindowContext(IOleInPlaceSite* This,
    IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc,
    LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSite);
    *ppFrame = (IOleInPlaceFrame*)&s->lpVtblInPlaceFrame;
    s->ref++; /* AddRef for the out param */
    *ppDoc = NULL;
    /* Fix 143: 用控件真实矩形, 不再硬编码 400x300 */
    RECT rFallback; rFallback.left=0; rFallback.top=0; rFallback.right=400; rFallback.bottom=300;
    const RECT* rc = (s->rcCtrl.right > s->rcCtrl.left) ? &s->rcCtrl : &rFallback;
    if (lprcPosRect)  memcpy(lprcPosRect, rc, sizeof(RECT));
    if (lprcClipRect) memcpy(lprcClipRect, rc, sizeof(RECT));
    if (lpFrameInfo) { /* 账 #305 第三百五十一刀: fMDIApp 按**宿主是否真有 MDIClient** 答 —— Codejock
     * 的 CommandBars/DockingPaneManager 靠它进 MDI 布局模式 (接管 MDI 客体,
     * 随窗体 resize 回流窗格); 恒 FALSE 便退化为普通窗体, maximize 后窗格
     * 不回流 (GUISample 实测)。 */
    lpFrameInfo->fMDIApp = (GetPropW(s->hwndForm, L"VB6_MDIClient") != NULL);
    lpFrameInfo->hwndFrame=s->hwndForm; lpFrameInfo->haccel=0; lpFrameInfo->cAccelEntries=0; }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPS_Scroll(IOleInPlaceSite* This, SIZE s) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnUIDeactivate(IOleInPlaceSite* This, BOOL f) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnInPlaceDeactivate(IOleInPlaceSite* This) { return S_OK; }
/* 账 (Outlook2013 frmPaneCalendar.SetFlatStyle): 这张表以前**漏了 DiscardUndoState**
 * (IOleInPlaceSite 的第 13 槽), 于是 DeactivateAndUndo / OnPosRectChange 各前移一格、
 * 末槽 OnPosRectChange 留空 —— 控件按真实槽位调 OnPosRectChange 就是调 NULL。
 * 槽位必须与 oleidl.h 的声明**逐格对齐**, 少一格不是"少一个功能", 是整条尾巴错位。 */
static HRESULT STDMETHODCALLTYPE axSite_IPS_DiscardUndoState(IOleInPlaceSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_DeactivateAndUndo(IOleInPlaceSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnPosRectChange(IOleInPlaceSite* This, LPCRECT lprc) { return S_OK; }

const IOleInPlaceSiteVtbl g_axSiteInPlaceSiteVtbl = {
    axSite_IPS_QueryInterface, axSite_IPS_AddRef, axSite_IPS_Release,
    axSite_IPS_GetWindow, axSite_IPS_ContextSensitiveHelp,
    axSite_IPS_CanInPlaceActivate, axSite_IPS_OnInPlaceActivate, axSite_IPS_OnUIActivate,
    axSite_IPS_GetWindowContext, axSite_IPS_Scroll,
    axSite_IPS_OnUIDeactivate, axSite_IPS_OnInPlaceDeactivate,
    axSite_IPS_DiscardUndoState,
    axSite_IPS_DeactivateAndUndo, axSite_IPS_OnPosRectChange
};

/* --- IOleInPlaceFrame --- */
static HRESULT STDMETHODCALLTYPE axSite_IPF_QueryInterface(IOleInPlaceFrame* This, REFIID riid, void** ppv) {
    return axSite_CS_QueryInterface((IOleClientSite*)This, riid, ppv);
}
static ULONG STDMETHODCALLTYPE axSite_IPF_AddRef(IOleInPlaceFrame* This) {
    return axSite_CS_AddRef((IOleClientSite*)SITE_OF(This, lpVtblInPlaceFrame));
}
static ULONG STDMETHODCALLTYPE axSite_IPF_Release(IOleInPlaceFrame* This) {
    return axSite_CS_Release((IOleClientSite*)SITE_OF(This, lpVtblInPlaceFrame));
}
static HRESULT STDMETHODCALLTYPE axSite_IPF_GetWindow(IOleInPlaceFrame* This, HWND* phwnd) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceFrame); *phwnd = s->hwndForm; return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPF_ContextSensitiveHelp(IOleInPlaceFrame* This, BOOL f) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_GetBorder(IOleInPlaceFrame* This, LPRECT lprc) { return E_NOTIMPL; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_RequestBorderSpace(IOleInPlaceFrame* This, LPCBORDERWIDTHS lpbw) { return E_NOTIMPL; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_SetBorderSpace(IOleInPlaceFrame* This, LPCBORDERWIDTHS lpbw) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_SetActiveObject(IOleInPlaceFrame* This, IOleInPlaceActiveObject* pObj, LPCOLESTR pszObj) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_InsertMenus(IOleInPlaceFrame* This, HMENU hShared, LPOLEMENUGROUPWIDTHS mw) { return E_NOTIMPL; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_SetMenu(IOleInPlaceFrame* This, HMENU hShared, HOLEMENU hOLEMenu, HWND hwndActiveObj) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_RemoveMenus(IOleInPlaceFrame* This, HMENU hShared) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_SetStatusText(IOleInPlaceFrame* This, LPCOLESTR psz) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_EnableModeless(IOleInPlaceFrame* This, BOOL f) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPF_TranslateAccelerators(IOleInPlaceFrame* This, LPMSG pMsg, WORD wID) { return S_FALSE; }

const IOleInPlaceFrameVtbl g_axSiteInPlaceFrameVtbl = {
    axSite_IPF_QueryInterface, axSite_IPF_AddRef, axSite_IPF_Release,
    axSite_IPF_GetWindow, axSite_IPF_ContextSensitiveHelp,
    axSite_IPF_GetBorder, axSite_IPF_RequestBorderSpace, axSite_IPF_SetBorderSpace,
    axSite_IPF_SetActiveObject,
    axSite_IPF_InsertMenus, axSite_IPF_SetMenu, axSite_IPF_RemoveMenus,
    axSite_IPF_SetStatusText, axSite_IPF_EnableModeless, axSite_IPF_TranslateAccelerators
};

/* ===== Fix 143d: 已创建的 OCX 宿主登记表 (windowless 控件消息转发用) ===== */
#define VB6_MAX_AXSITES 64
Vb6AxSite* g_axSites[VB6_MAX_AXSITES];
int vb6_axSiteCount = 0;

void axSiteRegister(Vb6AxSite* s) {
    if (vb6_axSiteCount < VB6_MAX_AXSITES) g_axSites[vb6_axSiteCount++] = s;
}

/* ===== 容器对象 (IOleContainer + IDispatch): 控件按名解析同窗体其他控件 =====
 * 账 (Outlook2013 cmdSearch 的 BuddyControl): frmInbox.Form_Resize 里
 * `cmdSearch.BuddyControl = "txtSearchBar"` 之后, 参考 exe 把按钮贴到 buddy 右缘
 * ((1129,164)-(1147,184)); C3 里按钮留在 .frm 设计位置。Codejock 的 PushButton
 * 解析 buddy 走的是 IOleClientSite::GetContainer → IOleContainer → QI(IDispatch)
 * → 容器分派的 GetIDsOfNames/Invoke (VB6 窗体就是这么给控件名的)。以前
 * GetContainer 恒返回 E_NOINTERFACE, buddy 找不到 ⇒ 按钮不动。
 * 这里补一具最小容器: 分派按名认出本窗体的控件实例, 交回控件自己的 IDispatch
 * (= VB6 容器 Extender 与控件同一条通道)。 */
static int axTrace(void) { return GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0; }

static Vb6AxSite* axContainerFindSite(HWND hwndForm, const wchar_t* name) {
    if (!name) return NULL;
    for (int i = 0; i < vb6_axSiteCount; i++) {
        Vb6AxSite* s = g_axSites[i];
        if (s && s->hwndForm == hwndForm && _wcsicmp(s->ctrlName, name) == 0) return s;
    }
    return NULL;
}

/* 托管控件交回容器的那份代理: 日志 + 把 hWnd 这档补齐, 其余原样转给控件自己.
 * 只为一具多出来, 所以 ref 独立计; QI 到别的接口直接问控件, 不参与计数. */
typedef struct Vb6AxCtlProxy {
    const IDispatchVtbl* lpVtbl;
    LONG ref;
    Vb6AxSite* site;
} Vb6AxCtlProxy;

typedef struct Vb6AxContainer {
    /* VB6 窗体交出去的是 **IOleItemContainer** (9 槽): 它继承 IParseDisplayName +
     * IOleContainer, 控件按名取另一枚控件走 ParseDisplayName/GetObject。
     * 只给 5 槽的 IOleContainer 时, 调用方按子接口槽位调用会越过表尾 (实测卡死)。
     * 槽位次序 (oleidl.h): QI/AddRef/Release, ParseDisplayName, EnumObjects,
     * LockContainer, GetObject, GetObjectStorage, IsRunning —— ParseDisplayName 夹在
     * 中间, 所以 IOleContainer 面**不是**本表的前 5 槽, 必须另给一张 5 槽表。 */
    const IOleItemContainerVtbl* lpVtblContainer;
    const IOleContainerVtbl* lpVtblOle;
    const IDispatchVtbl* lpVtblDispatch;
    LONG ref;
    HWND hwndForm;
    struct { wchar_t name[128]; DISPID dispid; } map[32];
    int mapCount;
} Vb6AxContainer;

#define CONTAINER_OF(p, f) ((Vb6AxContainer*)((char*)(p) - offsetof(Vb6AxContainer, f)))
#define PROXY_OF(p) ((Vb6AxCtlProxy*)(p))
#define AXCONTAINER_BASE  0x1000
#define AXCONTAINER_HWND  (AXCONTAINER_BASE - 1)
#define AXPROXY_HWND      0x1FFF
#define AXPROXY_NAME      0x2000

static const IOleItemContainerVtbl g_axContainerVtbl;
static const IDispatchVtbl    g_axContainerDispatchVtbl;

/* --- 代理 (buddy 等被容器交出去的控件) --- */
static HRESULT STDMETHODCALLTYPE axProxy_QueryInterface(IDispatch* This, REFIID riid, void** ppv) {
    Vb6AxCtlProxy* p = PROXY_OF(This);
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = &p->lpVtbl; p->ref++; return S_OK;
    }
    IDispatch* d = (IDispatch*)p->site->pDisp;
    HRESULT hr = d ? d->lpVtbl->QueryInterface(d, riid, ppv) : E_NOINTERFACE;
    if (axTrace()) {
        wchar_t g[64] = {0}; StringFromGUID2(riid, g, 64);
        fprintf(stderr, "[C3_AXC] proxy('%ls') QI %ls -> 0x%08lX\n", p->site->ctrlName, g, (unsigned long)hr);
        fflush(stderr);
    }
    return hr;
}
static ULONG STDMETHODCALLTYPE axProxy_AddRef(IDispatch* This) { return InterlockedIncrement(&PROXY_OF(This)->ref); }
static ULONG STDMETHODCALLTYPE axProxy_Release(IDispatch* This) {
    LONG r = InterlockedDecrement(&PROXY_OF(This)->ref);
    if (r <= 0) free(PROXY_OF(This));
    return r;
}
static HRESULT STDMETHODCALLTYPE axProxy_GetTypeInfoCount(IDispatch* This, UINT* p) { (void)This; *p = 0; return S_OK; }
static HRESULT STDMETHODCALLTYPE axProxy_GetTypeInfo(IDispatch* This, UINT i, LCID l, ITypeInfo** p) {
    (void)This; (void)i; (void)l; *p = NULL; return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE axProxy_GetIDsOfNames(IDispatch* This, REFIID riid, LPOLESTR* names,
                                                        UINT cNames, LCID lcid, DISPID* dispids) {
    Vb6AxCtlProxy* p = PROXY_OF(This);
    (void)lcid;
    if (!names || !dispids) return E_POINTER;
    IDispatch* d = (IDispatch*)p->site->pDisp;
    HRESULT hr = d ? d->lpVtbl->GetIDsOfNames(d, riid, names, cNames, lcid, dispids) : E_FAIL;
    for (UINT i = 0; i < cNames; i++) {
        /* 控件自己不认 hWnd / Name 时补上 (VB6 容器 Extender 的两档:
         * hWnd 是控件窗口, Name 是 .frm 里的实例名 —— 控件找 buddy 按名匹配靠它) */
        if ((FAILED(hr) || dispids[i] == DISPID_UNKNOWN) && names[i]) {
            if (_wcsicmp(names[i], L"hWnd") == 0) { dispids[i] = AXPROXY_HWND; hr = S_OK; }
            else if (_wcsicmp(names[i], L"Name") == 0) { dispids[i] = AXPROXY_NAME; hr = S_OK; }
        }
        if (axTrace()) {
            fprintf(stderr, "[C3_AXC] proxy('%ls') GetIDsOfNames '%ls' -> %ld\n",
                    p->site->ctrlName, names[i] ? names[i] : L"", (long)dispids[i]);
            fflush(stderr);
        }
    }
    return hr;
}
static HRESULT STDMETHODCALLTYPE axProxy_Invoke(IDispatch* This, DISPID dispid, REFIID riid, LCID lcid,
                                                WORD flags, DISPPARAMS* dp, VARIANT* result,
                                                EXCEPINFO* ei, UINT* ae) {
    Vb6AxCtlProxy* p = PROXY_OF(This);
    if (axTrace()) {
        fprintf(stderr, "[C3_AXC] proxy('%ls') Invoke dispid=%ld flags=0x%04X\n",
                p->site->ctrlName, (long)dispid, (unsigned)flags);
        fflush(stderr);
    }
    if (dispid == AXPROXY_HWND) {
        if (result) {
            VariantInit(result);
            result->vt = VT_I4;
            result->lVal = (LONG)(LONG_PTR)vb6_OcxHost_WindowOf(p->site->pDisp);
        }
        return S_OK;
    }
    if (dispid == AXPROXY_NAME) {
        if (result) {
            VariantInit(result);
            result->vt = VT_BSTR;
            result->bstrVal = SysAllocString(p->site->ctrlName);
        }
        return S_OK;
    }
    IDispatch* d = (IDispatch*)p->site->pDisp;
    if (!d) return DISP_E_MEMBERNOTFOUND;
    return d->lpVtbl->Invoke(d, dispid, riid, lcid, flags, dp, result, ei, ae);
}
static const IDispatchVtbl g_axProxyVtbl = {
    axProxy_QueryInterface, axProxy_AddRef, axProxy_Release,
    axProxy_GetTypeInfoCount, axProxy_GetTypeInfo,
    axProxy_GetIDsOfNames, axProxy_Invoke
};

/* --- 容器分派 --- */
static HRESULT STDMETHODCALLTYPE axC_D_QueryInterface(IDispatch* This, REFIID riid, void** ppv);
static ULONG STDMETHODCALLTYPE axC_D_AddRef(IDispatch* This);
static ULONG STDMETHODCALLTYPE axC_D_Release(IDispatch* This);

static HRESULT STDMETHODCALLTYPE axC_D_GetTypeInfoCount(IDispatch* This, UINT* p) { (void)This; *p = 0; return S_OK; }
static HRESULT STDMETHODCALLTYPE axC_D_GetTypeInfo(IDispatch* This, UINT i, LCID l, ITypeInfo** p) {
    (void)This; (void)i; (void)l; *p = NULL; return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE axC_D_GetIDsOfNames(IDispatch* This, REFIID riid, LPOLESTR* names,
                                                     UINT cNames, LCID lcid, DISPID* dispids) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblDispatch);
    (void)lcid;
    if (!names || !dispids) return E_POINTER;
    for (UINT i = 0; i < cNames; i++) {
        const wchar_t* n = names[i];
        DISPID d = DISPID_UNKNOWN;
        if (n && _wcsicmp(n, L"hWnd") == 0) {
            d = AXCONTAINER_HWND;   /* 容器自己的窗口 = 窗体 */
        } else if (n) {
            for (int k = 0; k < c->mapCount; k++)
                if (_wcsicmp(n, c->map[k].name) == 0) { d = c->map[k].dispid; break; }
            if (d == DISPID_UNKNOWN && axContainerFindSite(c->hwndForm, n)) {
                if (c->mapCount < 32) {
                    wcsncpy(c->map[c->mapCount].name, n, 127);
                    c->map[c->mapCount].name[127] = 0;
                    c->map[c->mapCount].dispid = AXCONTAINER_BASE + c->mapCount;
                    d = c->map[c->mapCount].dispid;
                    c->mapCount++;
                }
            }
        }
        dispids[i] = d;
        if (axTrace()) {
            fprintf(stderr, "[C3_AXC] container GetIDsOfNames '%ls' -> %ld\n",
                    n ? n : L"", (long)d);
            fflush(stderr);
        }
    }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axC_D_Invoke(IDispatch* This, DISPID dispid, REFIID riid, LCID lcid,
                                              WORD flags, DISPPARAMS* dp, VARIANT* result,
                                              EXCEPINFO* ei, UINT* ae) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblDispatch);
    (void)riid; (void)lcid; (void)flags; (void)dp; (void)ei; (void)ae;
    if (axTrace()) {
        fprintf(stderr, "[C3_AXC] container Invoke dispid=%ld flags=0x%04X\n",
                (long)dispid, (unsigned)flags);
        fflush(stderr);
    }
    if (result) VariantInit(result);
    if (dispid == AXCONTAINER_HWND) {
        if (result) { result->vt = VT_I4; result->lVal = (LONG)(LONG_PTR)c->hwndForm; }
        return S_OK;
    }
    int idx = (int)(dispid - AXCONTAINER_BASE);
    if (idx >= 0 && idx < c->mapCount) {
        Vb6AxSite* s = axContainerFindSite(c->hwndForm, c->map[idx].name);
        if (s && s->pDisp) {
            IDispatch* d = (IDispatch*)s->pDisp;
            /* 交代理: 记日志, 并给控件补 hWnd 这一档 (§axProxy_*) */
            Vb6AxCtlProxy* px = (Vb6AxCtlProxy*)calloc(1, sizeof(Vb6AxCtlProxy));
            if (px) {
                px->lpVtbl = &g_axProxyVtbl;
                px->ref = 1;
                px->site = s;
                if (axTrace()) {
                    fprintf(stderr, "[C3_AXC] container resolve '%ls' -> proxy pDisp=%p hwnd=%p\n",
                            c->map[idx].name, (void*)d, (void*)vb6_OcxHost_WindowOf(d));
                    fflush(stderr);
                }
                if (result) { result->vt = VT_DISPATCH; result->pdispVal = (IDispatch*)px; }
                else px->lpVtbl->Release((IDispatch*)px);
                return S_OK;
            }
        }
    }
    if (result) result->vt = VT_EMPTY;
    return DISP_E_MEMBERNOTFOUND;
}
static const IDispatchVtbl g_axContainerDispatchVtbl = {
    axC_D_QueryInterface, axC_D_AddRef, axC_D_Release,
    axC_D_GetTypeInfoCount, axC_D_GetTypeInfo,
    axC_D_GetIDsOfNames, axC_D_Invoke
};

/* --- 容器本体 --- */
/* 名字形态: 控件可能写 "\\1!txtSearchBar" / "!txtSearchBar" 这类 OLE 项名,
 * 只取 `!` 之后或最后一个 `\\` 之后那段. */
static const wchar_t* axC_ItemName(const wchar_t* s) {
    const wchar_t* r = s;
    for (const wchar_t* p = s; p && *p; p++) if (*p == L'!' || *p == L'\\') r = p + 1;
    return r ? r : L"";
}

static HRESULT STDMETHODCALLTYPE axC_QueryInterface(IOleItemContainer* This, REFIID riid, void** ppv) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblContainer);
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IOleItemContainer)) {
        *ppv = &c->lpVtblContainer; c->ref++;
        if (axTrace()) { fprintf(stderr, "[C3_AXC] container QI IUnknown/IOleItemContainer -> S_OK\n"); fflush(stderr); }
        return S_OK;
    }
    if (IsEqualIID(riid, &IID_IOleContainer)) {
        /* IOleContainer 面单独一张 5 槽表 (ParseDisplayName 插在中间, 不能共用) */
        *ppv = &c->lpVtblOle; c->ref++;
        if (axTrace()) { fprintf(stderr, "[C3_AXC] container QI IOleContainer -> S_OK\n"); fflush(stderr); }
        return S_OK;
    }
    if (IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = &c->lpVtblDispatch; c->ref++;
        if (axTrace()) { fprintf(stderr, "[C3_AXC] container QI IDispatch -> S_OK\n"); fflush(stderr); }
        return S_OK;
    }
    *ppv = NULL;
    if (axTrace()) {
        wchar_t g[64] = {0}; StringFromGUID2(riid, g, 64);
        fprintf(stderr, "[C3_AXC] container QI %ls -> E_NOINTERFACE\n", g);
        fflush(stderr);
    }
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE axC_AddRef(IOleItemContainer* This) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblContainer);
    LONG r = InterlockedIncrement(&c->ref);
    if (axTrace()) { fprintf(stderr, "[C3_AXC] container AddRef -> %ld\n", (long)r); fflush(stderr); }
    return r;
}
static ULONG STDMETHODCALLTYPE axC_Release(IOleItemContainer* This) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblContainer);
    LONG r = InterlockedDecrement(&c->ref);
    if (r <= 0) {
        if (GetPropW(c->hwndForm, L"VB6_AxContainer") == (HANDLE)c)
            RemovePropW(c->hwndForm, L"VB6_AxContainer");
        free(c);
    }
    return r;
}
/* --- IEnumUnknown: 控件找 buddy 时先枚举容器里的对象 (实测 Codejock
 * PushButton 在 SetProp BuddyControl 里正是 GetContainer → EnumObjects) --- */
typedef struct Vb6AxEnum {
    const IEnumUnknownVtbl* lpVtbl;
    LONG ref;
    int idx;
    int count;
    IDispatch* items[64];
} Vb6AxEnum;

static HRESULT STDMETHODCALLTYPE axE_QueryInterface(IEnumUnknown* This, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IEnumUnknown)) {
        *ppv = This; ((Vb6AxEnum*)This)->ref++; return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE axE_AddRef(IEnumUnknown* This) {
    return InterlockedIncrement(&((Vb6AxEnum*)This)->ref);
}
static ULONG STDMETHODCALLTYPE axE_Release(IEnumUnknown* This) {
    Vb6AxEnum* e = (Vb6AxEnum*)This;
    LONG r = InterlockedDecrement(&e->ref);
    if (r <= 0) {
        for (int i = 0; i < e->count; i++) if (e->items[i]) e->items[i]->lpVtbl->Release(e->items[i]);
        free(e);
    }
    return r;
}
static HRESULT STDMETHODCALLTYPE axE_Next(IEnumUnknown* This, ULONG celt, IUnknown** rgelt, ULONG* pceltFetched) {
    Vb6AxEnum* e = (Vb6AxEnum*)This;
    ULONG got = 0;
    if (!rgelt) return E_POINTER;
    while (got < celt && e->idx < e->count) {
        IDispatch* d = e->items[e->idx++];
        if (!d) continue;
        d->lpVtbl->AddRef(d);
        rgelt[got++] = (IUnknown*)d;
        if (axTrace()) { fprintf(stderr, "[C3_AXC] enum Next idx=%d disp=%p\n", e->idx - 1, (void*)d); fflush(stderr); }
    }
    if (pceltFetched) *pceltFetched = got;
    return got == celt ? S_OK : S_FALSE;
}
static HRESULT STDMETHODCALLTYPE axE_Skip(IEnumUnknown* This, ULONG celt) {
    Vb6AxEnum* e = (Vb6AxEnum*)This;
    e->idx += (int)celt;
    if (e->idx > e->count) e->idx = e->count;
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axE_Reset(IEnumUnknown* This) {
    ((Vb6AxEnum*)This)->idx = 0;
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axE_Clone(IEnumUnknown* This, IEnumUnknown** ppE) {
    (void)This; if (ppE) *ppE = NULL; return E_NOTIMPL;
}
static const IEnumUnknownVtbl g_axEnumVtbl = {
    axE_QueryInterface, axE_AddRef, axE_Release,
    axE_Next, axE_Skip, axE_Reset, axE_Clone
};

static IEnumUnknown* axEnumCreate(HWND hwndForm) {
    Vb6AxEnum* e = (Vb6AxEnum*)calloc(1, sizeof(Vb6AxEnum));
    if (!e) return NULL;
    e->lpVtbl = &g_axEnumVtbl;
    e->ref = 1;
    for (int i = 0; i < vb6_axSiteCount && e->count < 64; i++) {
        Vb6AxSite* s = g_axSites[i];
        if (s && s->hwndForm == hwndForm && s->pDisp) {
            /* 交代理而不是裸 IDispatch: VB6 容器枚举出来的对象认得 Name/hWnd
             * 这两档 Extender 属性 (§axProxy_*), 控件据此按名匹配 buddy。 */
            Vb6AxCtlProxy* px = (Vb6AxCtlProxy*)calloc(1, sizeof(Vb6AxCtlProxy));
            if (!px) continue;
            px->lpVtbl = &g_axProxyVtbl;
            px->ref = 1;
            px->site = s;
            e->items[e->count++] = (IDispatch*)px;
        }
    }
    return (IEnumUnknown*)e;
}

static HRESULT STDMETHODCALLTYPE axC_EnumObjects(IOleItemContainer* This, DWORD flags, IEnumUnknown** ppe) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblContainer);
    (void)flags;
    if (!ppe) return E_POINTER;
    *ppe = axEnumCreate(c->hwndForm);
    if (axTrace()) {
        fprintf(stderr, "[C3_AXC] container EnumObjects -> %p\n", (void*)*ppe);
        fflush(stderr);
    }
    return *ppe ? S_OK : E_OUTOFMEMORY;
}
static HRESULT STDMETHODCALLTYPE axC_LockContainer(IOleItemContainer* This, BOOL f) { (void)This; (void)f; return S_OK; }

/* 控件按名要 moniker: 交一枚空 moniker 没有意义, 直接说"没这个对象"
 * (调用方只把它当失败分支, 不影响宿主)。 */
static HRESULT STDMETHODCALLTYPE axC_ParseDisplayName(IOleItemContainer* This,
                                                      IBindCtx* pbc, LPOLESTR pszName,
                                                      ULONG* pchEaten, IMoniker** ppmkOut) {
    (void)This; (void)pbc;
    if (pchEaten) *pchEaten = pszName ? (ULONG)wcslen(pszName) : 0;
    if (ppmkOut) *ppmkOut = NULL;
    if (axTrace()) { fprintf(stderr, "[C3_AXC] container ParseDisplayName '%ls' -> MK_E_NOOBJECT\n",
                                pszName ? pszName : L""); fflush(stderr); }
    return MK_E_NOOBJECT;
}

/* 真正的按名取对象: 交控件自己的接口 (buddy 定位走这条) */
static HRESULT STDMETHODCALLTYPE axC_GetObject(IOleItemContainer* This, LPOLESTR pszItem,
                                               DWORD dwSpeedNeeded, IBindCtx* pbc,
                                               REFIID riid, void** ppvObject) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblContainer);
    (void)dwSpeedNeeded; (void)pbc;
    if (!ppvObject) return E_POINTER;
    *ppvObject = NULL;
    const wchar_t* n = axC_ItemName(pszItem);
    Vb6AxSite* s = axContainerFindSite(c->hwndForm, n);
    HRESULT hr = MK_E_NOOBJECT;
    if (s && s->pDisp) {
        IDispatch* d = (IDispatch*)s->pDisp;
        hr = d->lpVtbl->QueryInterface(d, riid, ppvObject);
    }
    if (axTrace()) {
        wchar_t g[64] = {0}; StringFromGUID2(riid, g, 64);
        fprintf(stderr, "[C3_AXC] container GetObject '%ls' QI %ls -> 0x%08lX\n",
                n, g, (unsigned long)hr);
        fflush(stderr);
    }
    return hr;
}
static HRESULT STDMETHODCALLTYPE axC_GetObjectStorage(IOleItemContainer* This, LPOLESTR pszItem,
                                                      IBindCtx* pbc, REFIID riid, void** ppvStorage) {
    (void)This; (void)pszItem; (void)pbc; (void)riid;
    if (ppvStorage) *ppvStorage = NULL;
    if (axTrace()) { fprintf(stderr, "[C3_AXC] container GetObjectStorage '%ls' -> E_NOTIMPL\n",
                                pszItem ? pszItem : L""); fflush(stderr); }
    return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE axC_IsRunning(IOleItemContainer* This, LPOLESTR pszItem) {
    Vb6AxContainer* c = CONTAINER_OF(This, lpVtblContainer);
    const wchar_t* n = axC_ItemName(pszItem);
    HRESULT hr = axContainerFindSite(c->hwndForm, n) ? S_OK : MK_E_NOOBJECT;
    if (axTrace()) { fprintf(stderr, "[C3_AXC] container IsRunning '%ls' -> 0x%08lX\n",
                                n, (unsigned long)hr); fflush(stderr); }
    return hr;
}

static const IOleItemContainerVtbl g_axContainerVtbl = {
    axC_QueryInterface, axC_AddRef, axC_Release,
    axC_ParseDisplayName, axC_EnumObjects, axC_LockContainer,
    axC_GetObject, axC_GetObjectStorage, axC_IsRunning
};

/* --- IOleContainer 面 (IOleClientSite::GetContainer 交出去的就是这一面) --- */
static HRESULT STDMETHODCALLTYPE axO_QueryInterface(IOleContainer* This, REFIID riid, void** ppv) {
    return axC_QueryInterface((IOleItemContainer*)CONTAINER_OF(This, lpVtblOle), riid, ppv);
}
static ULONG STDMETHODCALLTYPE axO_AddRef(IOleContainer* This) {
    return axC_AddRef((IOleItemContainer*)CONTAINER_OF(This, lpVtblOle));
}
static ULONG STDMETHODCALLTYPE axO_Release(IOleContainer* This) {
    return axC_Release((IOleItemContainer*)CONTAINER_OF(This, lpVtblOle));
}
static HRESULT STDMETHODCALLTYPE axO_EnumObjects(IOleContainer* This, DWORD flags, IEnumUnknown** ppe) {
    return axC_EnumObjects((IOleItemContainer*)CONTAINER_OF(This, lpVtblOle), flags, ppe);
}
static HRESULT STDMETHODCALLTYPE axO_LockContainer(IOleContainer* This, BOOL f) {
    return axC_LockContainer((IOleItemContainer*)CONTAINER_OF(This, lpVtblOle), f);
}
static const IOleContainerVtbl g_axContainerOleVtbl = {
    axO_QueryInterface, axO_AddRef, axO_Release,
    axO_EnumObjects, axO_LockContainer
};

static HRESULT STDMETHODCALLTYPE axC_D_QueryInterface(IDispatch* This, REFIID riid, void** ppv) {
    return axC_QueryInterface((IOleItemContainer*)CONTAINER_OF(This, lpVtblDispatch), riid, ppv);
}
static ULONG STDMETHODCALLTYPE axC_D_AddRef(IDispatch* This) { return axC_AddRef((IOleItemContainer*)CONTAINER_OF(This, lpVtblDispatch)); }
static ULONG STDMETHODCALLTYPE axC_D_Release(IDispatch* This) { return axC_Release((IOleItemContainer*)CONTAINER_OF(This, lpVtblDispatch)); }

/* 每窗体一具; 引用由窗口属性持 1 份 (随窗体存活), 交出去的那份各自 AddRef. */
static IOleContainer* axContainerGet(HWND hwndForm) {
    Vb6AxContainer* c = (Vb6AxContainer*)GetPropW(hwndForm, L"VB6_AxContainer");
    if (!c) {
        c = (Vb6AxContainer*)calloc(1, sizeof(Vb6AxContainer));
        if (!c) return NULL;
        c->lpVtblContainer = &g_axContainerVtbl;
        c->lpVtblOle = &g_axContainerOleVtbl;
        c->lpVtblDispatch = &g_axContainerDispatchVtbl;
        c->ref = 1;
        c->hwndForm = hwndForm;
        SetPropW(hwndForm, L"VB6_AxContainer", (HANDLE)c);
    }
    c->ref++;
    /* 实测: 按 IOleContainer 面 (5 槽) 交出去时, 调用方在本槽位之后越界取到空
     * 函数指针并 call (0x7033F151 读地址 0 崩溃)。交 9 槽的 IOleItemContainer 面
     * (它的 QI 仍能换到 IOleContainer 面) 以覆盖两种假定。 */
    return (IOleContainer*)&c->lpVtblContainer;
}


#ifdef __cplusplus
} // extern "C"
#endif
