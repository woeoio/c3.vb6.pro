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
        else if (_wcsicmp(n, L"AutoClip") == 0) d = DISPID_AMBIENT_AUTOCLIP;
        else if (_wcsicmp(n, L"Palette") == 0) d = DISPID_AMBIENT_PALETTE;
        dispids[i] = d;
        if (d == DISPID_UNKNOWN && GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
            Vb6AxSite* s0 = SITE_OF(This, lpVtblDispatch);
            fprintf(stderr, "[C3_OCX] site GetIDsOfNames ctrl='%ls' name='%ls' -> UNKNOWN\n",
                    s0->ctrlName, n);
            fflush(stderr);
        }
    }
    return S_OK;
}
/* ambient Font 用的真 IFont: 从宿主窗口当前字体 (LOGFONT) 造一个 oleaut32 的
 * StdFont。容器没写过字体时退回 VB6 默认 MS Sans Serif 8.25pt —— 与
 * vb6forms.c vb6_Vb6DefaultGuiFont 同一口径。控件要在 IPersistPropertyBag::Load
 * 之后才取 ambient Font, 那时窗体字体已经落账, 所以这里先问窗口。
 * 参数 plf 允许为 NULL (直接取默认)。 */
static void* axSiteMakeFont(const LOGFONTW* plf) {
    static const GUID kIID_IFont =
        {0xBEF6E003, 0xA874, 0x101A, {0x8B, 0xBA, 0x00, 0xAA, 0x00, 0x30, 0x0C, 0xAB}};
    FONTDESC fd;
    IFont* pf = NULL;
    double pt = 8.25;
    int dpiY;
    HDC hdc;
    if (plf && plf->lfHeight) {
        hdc = GetDC(NULL);
        dpiY = hdc ? GetDeviceCaps(hdc, LOGPIXELSY) : 96;
        if (hdc) ReleaseDC(NULL, hdc);
        pt = (double)(plf->lfHeight < 0 ? -plf->lfHeight : plf->lfHeight) * 72.0 / dpiY;
        if (pt <= 0.0) pt = 8.25;
    }
    memset(&fd, 0, sizeof(fd));
    fd.cbSizeofstruct = sizeof(FONTDESC);
    fd.lpstrName = (LPOLESTR)((plf && plf->lfFaceName[0]) ? plf->lfFaceName : L"MS Sans Serif");
    fd.cySize.int64 = (LONGLONG)(pt * 10000.0);
    fd.sWeight = (short)((plf && plf->lfWeight > 0) ? plf->lfWeight : 400);
    fd.sCharset = (short)(plf ? plf->lfCharSet : DEFAULT_CHARSET);
    fd.fItalic = (plf && plf->lfItalic) ? TRUE : FALSE;
    fd.fUnderline = (plf && plf->lfUnderline) ? TRUE : FALSE;
    fd.fStrikethrough = (plf && plf->lfStrikeOut) ? TRUE : FALSE;
    if (FAILED(OleCreateFontIndirect(&fd, &kIID_IFont, (void**)&pf))) return NULL;
    return pf;
}

/* 账 (Outlook2013 cmdSearch): 容器 ambient BackColor 是**窗体自己的 BackColor**
 * (VB6 语义), 不是恒定的 COLOR_BTNFACE。以前一律答 0x8000000F, 于是白底窗体上的
 * Codejock 扁平按钮内部被涂成 F0F0F0 —— 参考 (真 VB6) 里是白底 + ABABAB 边框。
 * 取色口径与 uc_host_create.inc 同一位主人: 窗体类注册时登记的那张表; 表里没有
 * (窗体没写 BackColor, 用默认色) 才退回 COLOR_BTNFACE。 */
static int axSiteFormBackColor(HWND hwndForm) {
    if (!hwndForm) return -1;
    wchar_t cls[128];
    if (GetClassNameW(hwndForm, cls, 127) <= 0) return -1;
    char u8[256] = {0};
    vb6_wideToU8Buf(cls, u8, (int)sizeof(u8));
    return vb6_Forms_QueryClassBg(u8);
}

static HRESULT STDMETHODCALLTYPE axSite_D_Invoke(IDispatch* This, DISPID dispid, REFIID riid, LCID lcid,
                                                 WORD flags, DISPPARAMS* dp, VARIANT* result,
                                                 EXCEPINFO* ei, UINT* ae) {
    (void)This; (void)riid; (void)lcid; (void)flags; (void)dp; (void)ei; (void)ae;
    if (!result) return S_OK;
    VariantInit(result);
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
        Vb6AxSite* s0 = SITE_OF(This, lpVtblDispatch);
        fprintf(stderr, "[C3_OCX] ambient invoke ctrl='%ls' dispid=%ld flags=0x%04X\n",
                s0->ctrlName, (long)dispid, (unsigned)flags);
    }
    switch (dispid) {
        case DISPID_AMBIENT_BACKCOLOR: {
            Vb6AxSite* s = SITE_OF(This, lpVtblDispatch);
            int bg = axSiteFormBackColor(s->hwndForm);
            result->vt = VT_I4;
            result->lVal = (LONG)(bg >= 0 ? bg : 0x8000000F);
            return S_OK;
        }
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
        case DISPID_AMBIENT_FONT: {
            /* 控件按 VB6 语义向容器要 ambient Font (如 frmInbox.txtSearchBar 这枚
             * Codejock FlatEdit 的设计期 .frm 里没有 Font 块)。以前落到 default
             * 回 DISP_E_MEMBERNOTFOUND, 控件于是退回自带的大号默认字体
             * (Segoe UI 9pt), 参考 VB6 里是 MS Sans Serif 8.25pt —— 文字明显偏大。
             * 容器字体以窗体窗口当前字体为准 (窗体没写过就是 VB6 默认)。 */
            Vb6AxSite* s = SITE_OF(This, lpVtblDispatch);
            if (!s->pAmbientFont) {
                LOGFONTW lf;
                /* 问窗口字体只许走 vb6_ControlFont 那一枚出口（哨兵 control_dc D11） */
                HFONT hf = vb6_ControlFont(s->hwndForm);
                s->pAmbientFont = (hf && GetObjectW(hf, sizeof(lf), &lf) == sizeof(lf))
                                      ? axSiteMakeFont(&lf) : axSiteMakeFont(NULL);
                if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
                    fprintf(stderr, "[C3_OCX] ambient Font ctrl='%ls' face='%ls' hf=%p obj=%p\n",
                            s->ctrlName, (hf && GetObjectW(hf, sizeof(lf), &lf) == sizeof(lf))
                                              ? lf.lfFaceName : L"MS Sans Serif",
                            (void*)hf, s->pAmbientFont);
            }
            if (!s->pAmbientFont) { result->vt = VT_EMPTY; return S_OK; }
            ((IDispatch*)s->pAmbientFont)->lpVtbl->AddRef((IDispatch*)s->pAmbientFont);
            result->vt = VT_DISPATCH;
            result->pdispVal = (IDispatch*)s->pAmbientFont;
            return S_OK;
        }
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
    if (lpFrameInfo) { /* 账 #305 第三百五十一刀: fMDIApp 按**宿主是否真有 MDIClient** 答 —— Codejock
     * 的 CommandBars/DockingPaneManager 靠它进 MDI 布局模式 (接管 MDI 客体,
     * 随窗体 resize 回流窗格); 恒 FALSE 便退化为普通窗体, maximize 后窗格
     * 不回流 (GUISample 实测)。 */
    lpFrameInfo->fMDIApp = (GetPropW(s->hwndForm, L"VB6_MDIClient") != NULL);
    lpFrameInfo->hwndFrame=s->hwndForm;
                       lpFrameInfo->haccel=0; lpFrameInfo->cAccelEntries=0; }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_Scroll(IOleInPlaceSiteWindowless* This, SIZE sz) { (void)This; (void)sz; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnUIDeactivate(IOleInPlaceSiteWindowless* This, BOOL f) { (void)This; (void)f; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnInPlaceDeactivate(IOleInPlaceSiteWindowless* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless); s->inPlaceActive = 0; return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_DiscardUndoState(IOleInPlaceSiteWindowless* This) { (void)This; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_DeactivateAndUndo(IOleInPlaceSiteWindowless* This) { (void)This; return S_OK; }
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnPosRectChange(IOleInPlaceSiteWindowless* This, LPCRECT r) { (void)This; (void)r; return S_OK; }
/* --- IOleInPlaceSiteEx 三格 (窗口无窗口站点继承它, 槽位必须补齐) --- */
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnInPlaceActivateEx(IOleInPlaceSiteWindowless* This, BOOL* pfNoRedraw, DWORD dwFlags) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    s->inPlaceActive = 1;
    if (pfNoRedraw) *pfNoRedraw = FALSE;   /* 要重画: 控件自己画, 宿主不代劳 */
    (void)dwFlags;
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_OnInPlaceDeactivateEx(IOleInPlaceSiteWindowless* This, BOOL fNoRedraw) {
    Vb6AxSite* s = SITE_OF(This, lpVtblInPlaceSiteWindowless);
    s->inPlaceActive = 0; (void)fNoRedraw; return S_OK;
}
static HRESULT STDMETHODCALLTYPE axSite_IPSW_RequestUIActivate(IOleInPlaceSiteWindowless* This) { (void)This; return S_OK; }

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
    return s->focused ? S_OK : S_FALSE;
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

/* 槽位与 oleidl.h 的 IOleInPlaceSiteWindowless 逐格对齐 (3 IUnknown + 12
 * IOleInPlaceSite + 3 IOleInPlaceSiteEx + 12 windowless 扩展 = 30 格)。
 *
 * 账 (Outlook2013 frmPaneCalendar.SetFlatStyle): 这张表以前只有 26 格 —— 漏了
 * DiscardUndoState / OnInPlaceActivateEx / OnInPlaceDeactivateEx / RequestUIActivate,
 * 于是从第 13 格起**整条尾巴错位四格**: 控件改完属性要 InvalidateRect (真槽 26)
 * 时落到的正是本表的第 26 项 OnDefWindowMessage —— 三个形参 (prc, fErase) 被当成
 * (msg, wp, lp, pl) 读, pl 拿到 0x13, 一句 `*pl = ...` 当场 AV (WRITE 0x13)。
 * 少一格不是"少一个功能", 是让后面每一格都去执行别人的函数。 */
const IOleInPlaceSiteWindowlessVtbl g_axSiteInPlaceSiteWindowlessVtbl = {
    axSite_IPSW_QueryInterface, axSite_IPSW_AddRef, axSite_IPSW_Release,
    axSite_IPSW_GetWindow, axSite_IPSW_ContextSensitiveHelp,
    axSite_IPSW_CanInPlaceActivate, axSite_IPSW_OnInPlaceActivate, axSite_IPSW_OnUIActivate,
    axSite_IPSW_GetWindowContext, axSite_IPSW_Scroll,
    axSite_IPSW_OnUIDeactivate, axSite_IPSW_OnInPlaceDeactivate,
    axSite_IPSW_DiscardUndoState,
    axSite_IPSW_DeactivateAndUndo, axSite_IPSW_OnPosRectChange,
    axSite_IPSW_OnInPlaceActivateEx, axSite_IPSW_OnInPlaceDeactivateEx, axSite_IPSW_RequestUIActivate,
    axSite_IPSW_CanWindowlessActivate, axSite_IPSW_GetCapture, axSite_IPSW_SetCapture,
    axSite_IPSW_GetFocus, axSite_IPSW_SetFocus, axSite_IPSW_GetDC, axSite_IPSW_ReleaseDC,
    axSite_IPSW_InvalidateRect, axSite_IPSW_InvalidateRgn, axSite_IPSW_ScrollRect,
    axSite_IPSW_AdjustRect, axSite_IPSW_OnDefWindowMessage
};


/* ===== IAdviseSink: VB6 容器的通知面 ===== */

static HRESULT STDMETHODCALLTYPE axSite_AS_QueryInterface(IAdviseSink* This, REFIID riid, void** ppv) {
    return axSite_CS_QueryInterface((IOleClientSite*)(void*)SITE_OF(This, lpVtblAdviseSink), riid, ppv);
}
static ULONG STDMETHODCALLTYPE axSite_AS_AddRef(IAdviseSink* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblAdviseSink); return InterlockedIncrement(&s->ref);
}
static ULONG STDMETHODCALLTYPE axSite_AS_Release(IAdviseSink* This) {
    Vb6AxSite* s = SITE_OF(This, lpVtblAdviseSink); LONG r = InterlockedDecrement(&s->ref);
    if (r <= 0) free(s);
    return (ULONG)r;
}
static void STDMETHODCALLTYPE axSite_AS_OnDataChange(IAdviseSink* This, FORMATETC* f, STGMEDIUM* st) {
    (void)This; (void)f; (void)st;
}
static void STDMETHODCALLTYPE axSite_AS_OnViewChange(IAdviseSink* This, DWORD aspects, LONG index) {
    Vb6AxSite* s = SITE_OF(This, lpVtblAdviseSink);
    (void)aspects; (void)index;
    if (s->hwndForm) InvalidateRect(s->hwndForm, &s->rcCtrl, FALSE);
}
static void STDMETHODCALLTYPE axSite_AS_OnRename(IAdviseSink* This, IMoniker* pmk) {
    (void)This; (void)pmk;
}
static void STDMETHODCALLTYPE axSite_AS_OnSave(IAdviseSink* This) { (void)This; }
static void STDMETHODCALLTYPE axSite_AS_OnClose(IAdviseSink* This) { (void)This; }

/* IAdviseSink 共 8 槽 (IUnknown 3 + OnDataChange/OnViewChange/OnRename/
 * OnSave/OnClose 5); OnShowWindow 属于 IOleClientSite, 不在此表. */
const IAdviseSinkVtbl g_axSiteAdviseSinkVtbl = {
    axSite_AS_QueryInterface, axSite_AS_AddRef, axSite_AS_Release,
    axSite_AS_OnDataChange, axSite_AS_OnViewChange, axSite_AS_OnRename,
    axSite_AS_OnSave, axSite_AS_OnClose
};


#ifdef __cplusplus
} // extern "C"
#endif
