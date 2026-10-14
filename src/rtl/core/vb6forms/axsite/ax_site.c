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
static void ocxTraceQI(const char* iface, REFIID riid, HRESULT hr) {
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) <= 0) return;
    wchar_t guid[64] = {0};
    StringFromGUID2(riid, guid, 64);
    fprintf(stderr, "[C3_OCX]   QI %s %ls -> 0x%08lX\n", iface, guid, (unsigned long)hr);
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
    } else {
        *ppv = NULL;
    }
    ocxTraceQI("site", riid, r);
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
static HRESULT STDMETHODCALLTYPE axSite_CS_GetContainer(IOleClientSite* This, IOleContainer** ppC) { *ppC=NULL; return E_NOINTERFACE; }
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
    if (lpFrameInfo) { lpFrameInfo->fMDIApp=FALSE; lpFrameInfo->hwndFrame=s->hwndForm; lpFrameInfo->haccel=0; lpFrameInfo->cAccelEntries=0; }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPS_Scroll(IOleInPlaceSite* This, SIZE s) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnUIDeactivate(IOleInPlaceSite* This, BOOL f) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnInPlaceDeactivate(IOleInPlaceSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_DeactivateAndUndo(IOleInPlaceSite* This) { return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPS_OnPosRectChange(IOleInPlaceSite* This, LPCRECT lprc) { return S_OK; }

const IOleInPlaceSiteVtbl g_axSiteInPlaceSiteVtbl = {
    axSite_IPS_QueryInterface, axSite_IPS_AddRef, axSite_IPS_Release,
    axSite_IPS_GetWindow, axSite_IPS_ContextSensitiveHelp,
    axSite_IPS_CanInPlaceActivate, axSite_IPS_OnInPlaceActivate, axSite_IPS_OnUIActivate,
    axSite_IPS_GetWindowContext, axSite_IPS_Scroll,
    axSite_IPS_OnUIDeactivate, axSite_IPS_OnInPlaceDeactivate,
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


#ifdef __cplusplus
} // extern "C"
#endif
