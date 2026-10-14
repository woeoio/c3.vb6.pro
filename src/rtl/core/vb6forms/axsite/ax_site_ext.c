// ax_site_ext.c - vb6forms_axsite 拆分片：windowless 控件宿主三件套: IDispatch(ambient) / IOleControlSite / IOleInPlaceSiteWindowless
//
// 内容 = 拆分前 vb6forms_axsite.c 第 591~592 / 596~832 行，纯搬移零重排无行为改动
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

/* ===== Fix 143d: 无窗口控件宿主的三件套接口 =====
 * 前置声明 (实现中互相引用) */

/* --- IDispatch: ambient 属性 (控件通过它取容器的 BackColor/Font/UserMode 等) --- */
static HRESULT STDMETHODCALLTYPE axSite_D_QueryInterface(IDispatch* This, REFIID riid, void** ppv) {
    return axSite_CS_QueryInterface((IOleClientSite*)This, riid, ppv);
}
static ULONG STDMETHODCALLTYPE axSite_D_AddRef(IDispatch* This) { return axSite_CS_AddRef((IOleClientSite*)SITE_OF(This, lpVtblDispatch)); }
static ULONG STDMETHODCALLTYPE axSite_D_Release(IDispatch* This) { return axSite_CS_Release((IOleClientSite*)SITE_OF(This, lpVtblDispatch)); }
static HRESULT STDMETHODCALLTYPE axSite_D_GetTypeInfoCount(IDispatch* This, UINT* p) { (void)This; *p = 0; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_D_GetTypeInfo(IDispatch* This, UINT i, LCID l, ITypeInfo** p) {
    (void)This; (void)i; (void)l; *p = NULL; return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE axSite_D_GetIDsOfNames(IDispatch* This, REFIID riid, LPOLESTR* names,
                                                        UINT cNames, LCID lcid, DISPID* dispids) {
    (void)This; (void)riid; (void)lcid;
    if (!names || !dispids) return E_POINTER;
    for (UINT i = 0; i < cNames; i++) {
        const wchar_t* n = names[i];
        DISPID d = DISPID_UNKNOWN;
        if (_wcsicmp(n, L"BackColor") == 0) d = DISPID_AMBIENT_BACKCOLOR;
        else if (_wcsicmp(n, L"ForeColor") == 0) d = DISPID_AMBIENT_FORECOLOR;
        else if (_wcsicmp(n, L"Font") == 0) d = DISPID_AMBIENT_FONT;
        else if (_wcsicmp(n, L"UserMode") == 0) d = DISPID_AMBIENT_USERMODE;
        else if (_wcsicmp(n, L"Appearance") == 0) d = DISPID_AMBIENT_APPEARANCE;
        else if (_wcsicmp(n, L"ScaleUnits") == 0) d = DISPID_AMBIENT_SCALEUNITS;
        else if (_wcsicmp(n, L"LocaleID") == 0) d = DISPID_AMBIENT_LOCALEID;
        else if (_wcsicmp(n, L"DisplayName") == 0) d = DISPID_AMBIENT_DISPLAYNAME;
        else if (_wcsicmp(n, L"TextAlign") == 0) d = DISPID_AMBIENT_TEXTALIGN;
        else if (_wcsicmp(n, L"DisplayAsDefault") == 0) d = DISPID_AMBIENT_DISPLAYASDEFAULT;
        else if (_wcsicmp(n, L"UIDead") == 0) d = DISPID_AMBIENT_UIDEAD;
        else if (_wcsicmp(n, L"SupportsMnemonics") == 0) d = DISPID_AMBIENT_SUPPORTSMNEMONICS;
        else if (_wcsicmp(n, L"ShowGrabHandles") == 0) d = DISPID_AMBIENT_SHOWGRABHANDLES;
        else if (_wcsicmp(n, L"ShowHatching") == 0) d = DISPID_AMBIENT_SHOWHATCHING;
        else if (_wcsicmp(n, L"MessageReflect") == 0) d = DISPID_AMBIENT_MESSAGEREFLECT;
        else if (_wcsicmp(n, L"AutoClip") == 0) d = -713;
        else if (_wcsicmp(n, L"Palette") == 0) d = DISPID_AMBIENT_PALETTE;
        dispids[i] = d;
    }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_D_Invoke(IDispatch* This, DISPID dispid, REFIID riid, LCID lcid,
                                                 WORD flags, DISPPARAMS* dp, VARIANT* result,
                                                 EXCEPINFO* ei, UINT* ae) {
    (void)This; (void)riid; (void)lcid; (void)flags; (void)dp; (void)ei; (void)ae;
    if (!result) return S_OK;
    VariantInit(result);
    switch (dispid) {
        case DISPID_AMBIENT_BACKCOLOR:          result->vt = VT_I4;  result->lVal = 0x8000000F; return S_OK;
        case DISPID_AMBIENT_FORECOLOR:          result->vt = VT_I4;  result->lVal = 0x80000012; return S_OK;
        case DISPID_AMBIENT_APPEARANCE:         result->vt = VT_I2;  result->iVal = 1;          return S_OK;
        case DISPID_AMBIENT_USERMODE:           result->vt = VT_BOOL; result->boolVal = VARIANT_TRUE; return S_OK;
        case DISPID_AMBIENT_LOCALEID:           result->vt = VT_I4;  result->lVal = 0x409;      return S_OK;
        case DISPID_AMBIENT_TEXTALIGN:          result->vt = VT_I2;  result->iVal = 0;          return S_OK;
        case DISPID_AMBIENT_DISPLAYASDEFAULT:   result->vt = VT_BOOL; result->boolVal = VARIANT_FALSE; return S_OK;
        case DISPID_AMBIENT_UIDEAD:             result->vt = VT_BOOL; result->boolVal = VARIANT_FALSE; return S_OK;
        case DISPID_AMBIENT_SUPPORTSMNEMONICS:  result->vt = VT_BOOL; result->boolVal = VARIANT_TRUE;  return S_OK;
        case DISPID_AMBIENT_SHOWGRABHANDLES:    result->vt = VT_BOOL; result->boolVal = VARIANT_FALSE; return S_OK;
        case DISPID_AMBIENT_SHOWHATCHING:       result->vt = VT_BOOL; result->boolVal = VARIANT_FALSE; return S_OK;
        case DISPID_AMBIENT_MESSAGEREFLECT:     result->vt = VT_BOOL; result->boolVal = VARIANT_FALSE; return S_OK;
        case DISPID_AMBIENT_SCALEUNITS:         result->vt = VT_BSTR; result->bstrVal = SysAllocString(L"twip"); return S_OK;
        case DISPID_AMBIENT_DISPLAYNAME:        result->vt = VT_BSTR; result->bstrVal = SysAllocString(L"");     return S_OK;
        default:                                result->vt = VT_EMPTY; return DISP_E_MEMBERNOTFOUND;
    }
}

const IDispatchVtbl g_axSiteDispatchVtbl = {
    axSite_D_QueryInterface, axSite_D_AddRef, axSite_D_Release,
    axSite_D_GetTypeInfoCount, axSite_D_GetTypeInfo,
    axSite_D_GetIDsOfNames, axSite_D_Invoke
};

/* --- IOleControlSite --- */
static HRESULT STDMETHODCALLTYPE axSite_OCS_QueryInterface(IOleControlSite* This, REFIID riid, void** ppv) {
    return axSite_CS_QueryInterface((IOleClientSite*)This, riid, ppv);
}
static ULONG STDMETHODCALLTYPE axSite_OCS_AddRef(IOleControlSite* This) { return axSite_CS_AddRef((IOleClientSite*)SITE_OF(This, lpVtblControlSite)); }
static ULONG STDMETHODCALLTYPE axSite_OCS_Release(IOleControlSite* This) { return axSite_CS_Release((IOleClientSite*)SITE_OF(This, lpVtblControlSite)); }
static HRESULT STDMETHODCALLTYPE axSite_OCS_OnControlInfoChanged(IOleControlSite* This) { (void)This; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_OCS_LockInPlaceActive(IOleControlSite* This, BOOL f) { (void)This; (void)f; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_OCS_GetExtendedControl(IOleControlSite* This, IDispatch** ppDisp) {
    Vb6AxSite* s = SITE_OF(This, lpVtblControlSite);
    if (!ppDisp) return E_POINTER;
    /* Fix 148: VB6 控件用 UserControl.Extender 访问容器
     * (Extender.Container / Extender.TabIndex). 返回 E_NOTIMPL 时 VB6 运行时
     * 在其内部解引用空指针 — 实测 NewTab01.ocx 读 TDIMode 属性即崩
     * (0xC000041D 从 WndProc 逃逸). 这里给出轻量 Extender. */
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_OCX]   GetExtendedControl called (ctrl='%ls')\n", s->ctrlName);
    if (!s->pExtender) {
        s->pExtender = vb6_AxContainer_CreateExtender(s->hwndForm, s->ctrlName);
    }
    if (!s->pExtender) { *ppDisp = NULL; return E_OUTOFMEMORY; }
    *ppDisp = (IDispatch*)s->pExtender;
    ((IDispatch*)s->pExtender)->lpVtbl->AddRef((IDispatch*)s->pExtender);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_OCS_TransformCoords(IOleControlSite* This, POINTL* ptlHimetric,
                                                            POINTF* ptfContainer, DWORD flags) {
    (void)This;
    /* VB6 的 XFORMCOORDS_* 语义: HIMETRIC <-> 容器像素。折算走 vb6forms.c 那一对 float 权威
     * (账 #298 §B129 并表) —— 以前这一行把 DPI 写死成 96（自己拿一个常数去除那一档长度），
     * 换 DPI 的机器上容器坐标就跟着缩（与 #184 那条 VBFlexGrid 网格缩 20% 同一族）。
     * 两轴各问各的 DPI；表示留在浮点里，交出去的是 POINTF，量化成整数 px 会替容器改掉决定。 */
    if (!ptlHimetric || !ptfContainer) return E_POINTER;
    if (flags & XFORMCOORDS_HIMETRICTOCONTAINER) {
        ptfContainer->x = vb6_HimetricToPxF((float)ptlHimetric->x, (float)vb6_DpiX());
        ptfContainer->y = vb6_HimetricToPxF((float)ptlHimetric->y, (float)vb6_DpiY());
    } else {
        ptlHimetric->x = (LONG)vb6_PxToHimetricF(ptfContainer->x, (float)vb6_DpiX());
        ptlHimetric->y = (LONG)vb6_PxToHimetricF(ptfContainer->y, (float)vb6_DpiY());
    }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_OCS_TranslateAccelerator(IOleControlSite* This, MSG* pMsg, DWORD grfModifiers) {
    (void)This; (void)pMsg; (void)grfModifiers;
    return S_FALSE;   /* 不处理, 交给容器默认流程 */
}
static HRESULT STDMETHODCALLTYPE axSite_OCS_OnFocus(IOleControlSite* This, BOOL fGotFocus) { (void)This; (void)fGotFocus; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_OCS_ShowPropertyFrame(IOleControlSite* This) { (void)This; return E_NOTIMPL; }

const IOleControlSiteVtbl g_axSiteControlSiteVtbl = {
    axSite_OCS_QueryInterface, axSite_OCS_AddRef, axSite_OCS_Release,
    axSite_OCS_OnControlInfoChanged, axSite_OCS_LockInPlaceActive,
    axSite_OCS_GetExtendedControl, axSite_OCS_TransformCoords,
    axSite_OCS_TranslateAccelerator, axSite_OCS_OnFocus, axSite_OCS_ShowPropertyFrame
};

/* --- IOleInPlaceSiteWindowless (继承 IOleInPlaceSite 的 10 个方法 + 11 个扩展) --- */
static HRESULT STDMETHODCALLTYPE axSite_IPSW_QueryInterface(IOleInPlaceSiteWindowless* This, REFIID riid, void** ppv) {
    return axSite_CS_QueryInterface((IOleClientSite*)This, riid, ppv);
}
static ULONG STDMETHODCALLTYPE axSite_IPSW_AddRef(IOleInPlaceSiteWindowless* This) { return axSite_CS_AddRef((IOleClientSite*)SITE_OF(This, lpVtblInPlaceSiteWindowless)); }
static ULONG STDMETHODCALLTYPE axSite_IPSW_Release(IOleInPlaceSiteWindowless* This) { return axSite_CS_Release((IOleClientSite*)SITE_OF(This, lpVtblInPlaceSiteWindowless)); }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_GetWindow(IOleInPlaceSiteWindowless* This, HWND* phwnd) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless); *phwnd = s->hwndForm; return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_ContextSensitiveHelp(IOleInPlaceSiteWindowless* This, BOOL f) { (void)This; (void)f; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_CanInPlaceActivate(IOleInPlaceSiteWindowless* This) { (void)This; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnInPlaceActivate(IOleInPlaceSiteWindowless* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless); s->inPlaceActive = 1; return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnUIActivate(IOleInPlaceSiteWindowless* This) { (void)This; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_GetWindowContext(IOleInPlaceSiteWindowless* This,
    IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc,
    LPRECT lprcPosRect, LPRECT lprcClipRect, LPOLEINPLACEFRAMEINFO lpFrameInfo) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    if (ppFrame) { *ppFrame = (IOleInPlaceFrame*)&s->lpVtblInPlaceFrame; s->ref++; }
    if (ppDoc) *ppDoc = NULL;
    RECT rb; rb.left=0; rb.top=0; rb.right=400; rb.bottom=300;
    const RECT* rc = (s->rcCtrl.right > s->rcCtrl.left) ? &s->rcCtrl : &rb;
    if (lprcPosRect)  memcpy(lprcPosRect, rc, sizeof(RECT));
    if (lprcClipRect) memcpy(lprcClipRect, rc, sizeof(RECT));
    if (lpFrameInfo) { lpFrameInfo->fMDIApp=FALSE; lpFrameInfo->hwndFrame=s->hwndForm;
                       lpFrameInfo->haccel=0; lpFrameInfo->cAccelEntries=0; }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_Scroll(IOleInPlaceSiteWindowless* This, SIZE sz) { (void)This; (void)sz; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnUIDeactivate(IOleInPlaceSiteWindowless* This, BOOL f) { (void)This; (void)f; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnInPlaceDeactivate(IOleInPlaceSiteWindowless* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless); s->inPlaceActive = 0; return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_DeactivateAndUndo(IOleInPlaceSiteWindowless* This) { (void)This; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnPosRectChange(IOleInPlaceSiteWindowless* This, LPCRECT r) { (void)This; (void)r; return S_OK; }

/* ---- windowless 扩展 ---- */
static HRESULT STDMETHODCALLTYPE axSite_IPSW_CanWindowlessActivate(IOleInPlaceSiteWindowless* This) {
    (void)This; return S_OK;   /* 允许无窗口激活: 控件直接画到窗体 DC */
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_GetCapture(IOleInPlaceSiteWindowless* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    return (s->captured && GetCapture() == s->hwndForm) ? S_OK : S_FALSE;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_SetCapture(IOleInPlaceSiteWindowless* This, BOOL fCapture) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    s->captured = fCapture ? 1 : 0;
    if (fCapture) SetCapture(s->hwndForm); else ReleaseCapture();
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_GetFocus(IOleInPlaceSiteWindowless* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    return (GetFocus() == s->hwndForm) ? S_OK : S_FALSE;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_SetFocus(IOleInPlaceSiteWindowless* This, BOOL fFocus) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    s->focused = fFocus ? 1 : 0;
    if (fFocus) SetFocus(s->hwndForm);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_GetDC(IOleInPlaceSiteWindowless* This, LPCRECT prc, DWORD flags, HDC* phdc) {
    (void)prc;
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    if (!phdc) return E_POINTER;
    *phdc = GetDC(s->hwndForm);
    if (!*phdc) return E_FAIL;
    if (flags & OLEDC_NODRAW) return S_OK;
    /* 裁剪到控件矩形, 保证控件不会画到窗体其它区域 */
    if (s->rcCtrl.right > s->rcCtrl.left) {
        IntersectClipRect(*phdc, s->rcCtrl.left, s->rcCtrl.top, s->rcCtrl.right, s->rcCtrl.bottom);
    }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_ReleaseDC(IOleInPlaceSiteWindowless* This, HDC hdc) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    if (hdc) ReleaseDC(s->hwndForm, hdc);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_InvalidateRect(IOleInPlaceSiteWindowless* This, LPCRECT prc, BOOL fErase) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    InvalidateRect(s->hwndForm, prc, fErase);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_InvalidateRgn(IOleInPlaceSiteWindowless* This, HRGN hRgn, BOOL fErase) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    InvalidateRgn(s->hwndForm, hRgn, fErase);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_ScrollRect(IOleInPlaceSiteWindowless* This, INT dx, INT dy, LPCRECT prc, LPCRECT prcScroll) {
    (void)This; (void)dx; (void)dy; (void)prc; (void)prcScroll;
    return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_AdjustRect(IOleInPlaceSiteWindowless* This, LPRECT prc) { (void)This; (void)prc; return S_FALSE; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnDefWindowMessage(IOleInPlaceSiteWindowless* This, UINT msg, WPARAM wp, LPARAM lp, LRESULT* pl) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    *pl = DefWindowProcW(s->hwndForm, msg, wp, lp);
    return S_OK;
}

const IOleInPlaceSiteWindowlessVtbl g_axSiteInPlaceSiteWindowlessVtbl = {
    axSite_IPSW_QueryInterface, axSite_IPSW_AddRef, axSite_IPSW_Release,
    axSite_IPSW_GetWindow, axSite_IPSW_ContextSensitiveHelp,
    axSite_IPSW_CanInPlaceActivate, axSite_IPSW_OnInPlaceActivate, axSite_IPSW_OnUIActivate,
    axSite_IPSW_GetWindowContext, axSite_IPSW_Scroll,
    axSite_IPSW_OnUIDeactivate, axSite_IPSW_OnInPlaceDeactivate,
    axSite_IPSW_DeactivateAndUndo, axSite_IPSW_OnPosRectChange,
    axSite_IPSW_CanWindowlessActivate, axSite_IPSW_GetCapture, axSite_IPSW_SetCapture,
    axSite_IPSW_GetFocus, axSite_IPSW_SetFocus, axSite_IPSW_GetDC, axSite_IPSW_ReleaseDC,
    axSite_IPSW_InvalidateRect, axSite_IPSW_InvalidateRgn, axSite_IPSW_ScrollRect,
    axSite_IPSW_AdjustRect, axSite_IPSW_OnDefWindowMessage
};


#ifdef __cplusplus
} // extern "C"
#endif
