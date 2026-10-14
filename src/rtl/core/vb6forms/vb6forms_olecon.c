// vb6forms_olecon.c - C29-OLE: VB.OLE 容器控件 (OLE Container) —— 自研复刻
//
// VB6 的 OLE 控件把**外部 OLE 对象**（Excel 工作表 / Word 文档 / 图片 / 任意已注册的
// OLE 服务器）嵌入或链接到窗体上, 支持就地激活编辑。它天生就是 OLE —— 没有"原生替身",
// 所以按 §五-3 的口径用 ole32 的容器侧 API 实现, 不加载任何 .ocx。
//
// 与 axsite/（第三方 OCX 真宿主）的区别, 想照抄前先读:
//   * axsite 的对象在设计期就写死在 .frm 里（CLSID 已知）, 定位是"把 OCX 当子窗口塞进来"
//   * 本控件的对象由属性决定（Class / SourceDoc / CreateEmbed / CreateLink, 运行期也能换）,
//     定位是"窗口本身就是一个 OLE 容器"
//   * 容器的 **SaveObject / GetContainer / IStorage 必须是真实现** —— 服务器要靠它们保存
//     自己、枚举兄弟对象、定位; axsite 那套是 E_NOTIMPL/E_NOINTERFACE 也够用的最小站点。
//     照抄 axsite 会把"能显示但存不了盘"当成功。
//
// 结构: `VB6_OLECONTAINER` 自注册窗口类; 每个窗口的 GWLP_USERDATA 挂一个 Vb6OleCon,
// 它同时是 IOleClientSite / IOleInPlaceSite / IOleInPlaceFrame / IOleContainer /
// IOleItemContainer 五套 vtable 的宿主（多接口共用同一对象, 用 SITE_OF 反推首址 ——
// QI 返回的是对象内某个 vtable 字段的地址, 直接 (Vb6OleCon*)This 只在首字段上恰好对）。
//
// 持久化: 用 ILockBytes 的**内存复合文档**当对象的 IStorage, SaveToFile 时把它落盘。
// 这是经典 OLE 容器的做法, 不依赖任何临时文件。
//
// 判据: tests/olecon（**只本地跑**, 按用户指示不进 CI —— 嵌入对象需要目标机器装了
// 对应的 OLE 服务器, GA 的 runner 不带 Office）。系统自带的 `Package`（packager.dll,
// amd64/i386 都在）是最合适的判据对象。

#define _CRT_SECURE_NO_WARNINGS
#include <windows.h>
#include <objbase.h>
#include <ole2.h>
#include <oleauto.h>
#include <olectl.h>
// oledlg.h 只取结构体与常量定义 —— 函数本体走 LoadLibrary 现取（见 §六-3: 不新增 import lib）
#include <oledlg.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#ifndef __cplusplus
#include <stdint.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6forms_prop_ctrl.h"
#include "vb6rtl.h"      /* vb6_BSTR_FromStr */

// ============================================================
// 数据
// ============================================================

#define VB6_OLECON_CLASS L"VB6_OLECONTAINER"

#define VB6_OLETYPE_LINKED   0
#define VB6_OLETYPE_EMBEDDED 1
#define VB6_OLETYPE_NONE     3

typedef struct Vb6OleCon Vb6OleCon;

struct Vb6OleCon {
    /* --- 五套接口的 vtable 字段 (顺序即 QI 的映射表) --- */
    const IOleClientSiteVtbl*        lpVtblClientSite;
    const IOleInPlaceSiteVtbl*       lpVtblInPlaceSite;
    const IOleInPlaceFrameVtbl*      lpVtblInPlaceFrame;
    const IOleContainerVtbl*         lpVtblContainer;
    const IOleItemContainerVtbl*     lpVtblItemContainer;

    LONG ref;
    HWND hwnd;

    /* --- 被嵌入的对象及其接口 --- */
    IOleObject*      pOleObj;
    IViewObject2*    pViewObj;
    IDataObject*     pDataObj;
    IPersistStorage* pPersist;
    IDispatch*       pDisp;          /* 供 .Object 属性 */
    IStorage*        pStorage;       /* 对象的复合文档存储 */
    ILockBytes*      pLockBytes;     /* 内存 backing */

    int inPlaceActive;
    int oletType;                    /* 0=链接 1=嵌入 3=无 */
    int oletTypeAllowed;             /* 0=仅链接 1=仅嵌入 2=皆可 (默认) */
    int displayAsIcon;
    int autoActivate;                /* 0=手动 1=获焦点 2=双击(默认) 3=单击 */
    int autoVerbMenu;                /* 默认 1 */
    int sizeMode;                    /* 0=Clip 1=Stretch 2=AutoSize 3=Auto */
    int borderStyle;                 /* 0=无 1=单线 */
    int updating;                    /* CreateEmbed/Link 期间抑制事件 */

    wchar_t className[128];
    wchar_t sourceDoc[MAX_PATH];
    wchar_t sourceItem[128];
};

/* COM 多重接口共用同一个对象, 但 QI 返回的是**对象内某个 vtable 字段的地址**,
 * 所以每个方法里必须减去该字段的 offset 才能拿到对象首地址。 */
#define SITE_OF(thisptr, field) ((Vb6OleCon*)((char*)(thisptr) - offsetof(Vb6OleCon, field)))

/* 前台 trace: C3_OLECON_TRACE=1 打印各步 HRESULT (诊断"对象建不出来"用) */
static int oleconTrace(void) {
    static int t = -1;
    if (t < 0) t = (GetEnvironmentVariableW(L"C3_OLECON_TRACE", NULL, 0) > 0) ? 1 : 0;
    return t;
}
#define OLC_LOG(...) do { if (oleconTrace()) fprintf(stderr, __VA_ARGS__); } while (0)

// ============================================================
// 前置声明
// ============================================================

static void  olcReleaseObject(Vb6OleCon* s);
static HRESULT olcEnsureStorage(Vb6OleCon* s);
static void  olcCloseStorage(Vb6OleCon* s);
static int   olcActivate(Vb6OleCon* s, LONG verb);
static void  olcRepaint(Vb6OleCon* s);
static void  olcSizeToObject(Vb6OleCon* s);
static LRESULT CALLBACK vb6_OleConWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp);

static Vb6OleCon* olcFromHwnd(HWND hwnd) {
    if (!hwnd) return NULL;
    return (Vb6OleCon*)(INT_PTR)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
}

// ============================================================
// IUnknown —— 五套接口都转发到这里
// ============================================================

static HRESULT STDMETHODCALLTYPE olc_QueryInterface(IUnknown* This, REFIID riid, void** ppv) {
    Vb6OleCon* s;
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (!This) return E_NOINTERFACE;
    s = (Vb6OleCon*)This;   /* IUnknown 是首字段, 直接就是对象首址 */

    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IOleClientSite)) {
        *ppv = (void*)&s->lpVtblClientSite;
    } else if (IsEqualIID(riid, &IID_IOleInPlaceSite)) {
        *ppv = (void*)&s->lpVtblInPlaceSite;
    } else if (IsEqualIID(riid, &IID_IOleInPlaceFrame)) {
        *ppv = (void*)&s->lpVtblInPlaceFrame;
    } else if (IsEqualIID(riid, &IID_IOleContainer)) {
        *ppv = (void*)&s->lpVtblContainer;
    } else if (IsEqualIID(riid, &IID_IOleItemContainer)) {
        *ppv = (void*)&s->lpVtblItemContainer;
    } else {
        return E_NOINTERFACE;
    }
    s->ref++;
    return S_OK;
}

static ULONG STDMETHODCALLTYPE olc_AddRef(IUnknown* This) {
    return (ULONG)InterlockedIncrement(&((Vb6OleCon*)This)->ref);
}

static ULONG STDMETHODCALLTYPE olc_Release(IUnknown* This) {
    /* 站点对象的所有权归窗口 (WM_DESTROY 时 free), 这里只减计数不销毁 ——
     * 服务器可能在窗口已死之后还持有我们的接口指针, 提前 free 会 UAF。 */
    return (ULONG)InterlockedDecrement(&((Vb6OleCon*)This)->ref);
}

/* 其余四套接口的 QI/AddRef/Release 一律转发 (用 SITE_OF 回到首址)。
 * 形参类型必须写成各自接口类型 —— 写 void* 虽能编过但赋值进 vtable 是
 * C4113 (函数指针类型不匹配), 而 vtable 字段类型是严格的。 */
#define OLC_FWD_IU(IFACE, PFX, FIELD)                                            \
    static HRESULT STDMETHODCALLTYPE PFX##_QueryInterface(IFACE* This, REFIID r, void** p) { \
        return olc_QueryInterface((IUnknown*)SITE_OF(This, FIELD), r, p); }              \
    static ULONG STDMETHODCALLTYPE PFX##_AddRef(IFACE* This) {                   \
        return olc_AddRef((IUnknown*)SITE_OF(This, FIELD)); }                    \
    static ULONG STDMETHODCALLTYPE PFX##_Release(IFACE* This) {                  \
        return olc_Release((IUnknown*)SITE_OF(This, FIELD)); }

OLC_FWD_IU(IOleInPlaceSite,    olc_IPS, lpVtblInPlaceSite)
OLC_FWD_IU(IOleInPlaceFrame,   olc_IPF, lpVtblInPlaceFrame)
OLC_FWD_IU(IOleContainer,      olc_IC,  lpVtblContainer)
OLC_FWD_IU(IOleItemContainer,  olc_IIC, lpVtblItemContainer)
OLC_FWD_IU(IOleClientSite,     olc_CS,  lpVtblClientSite)

// ============================================================
// IOleClientSite
// ============================================================

/* SaveObject: **真实现** —— 服务器请求容器保存它。
 * 把对象按 IPersistStorage::Save 写进我们那本复合文档; 拿不到 IPersistStorage
 * (有些服务器只支持 IPersistStream) 就退 OleSave。 */
static HRESULT STDMETHODCALLTYPE olc_CS_SaveObject(IOleClientSite* This) {
    Vb6OleCon* s = SITE_OF(This, lpVtblClientSite);
    HRESULT hr = S_OK;
    if (!s->pOleObj || !s->pStorage) return E_FAIL;
    if (s->pPersist) {
        hr = s->pPersist->lpVtbl->Save(s->pPersist, s->pStorage, TRUE);
    } else {
        hr = OleSave((IPersistStorage*)s->pOleObj, s->pStorage, TRUE);
    }
    OLC_LOG("[OLECON] SaveObject hr=0x%08lX\n", (unsigned long)hr);
    if (SUCCEEDED(hr)) hr = s->pStorage->lpVtbl->Commit(s->pStorage, STGC_DEFAULT);
    return hr;
}

/* GetMoniker: 链接对象用。嵌入对象不需要 —— 返回 E_NOTIMPL 是合法的
 * (VB6 的嵌入容器也这么答), 但**链接**要给出文件 moniker, 否则 SetMoniker 失败。 */
static HRESULT STDMETHODCALLTYPE olc_CS_GetMoniker(IOleClientSite* This, DWORD assign,
                                                   DWORD which, IMoniker** ppM) {
    Vb6OleCon* s = SITE_OF(This, lpVtblClientSite);
    if (!ppM) return E_POINTER;
    *ppM = NULL;
    if (which == OLEWHICHMK_CONTAINER && s->sourceDoc[0]) {
        /* 链接源所在文件的 moniker */
        return CreateFileMoniker(s->sourceDoc, ppM);
    }
    return E_NOTIMPL;
}

/* GetContainer: **真实现** —— 服务器靠它找到 IOleContainer 去枚举兄弟对象 /
 * 解析相对名字。返回 E_NOINTERFACE (axsite 的做法) 会让部分服务器视为"没有容器"。 */
static HRESULT STDMETHODCALLTYPE olc_CS_GetContainer(IOleClientSite* This, IOleContainer** ppC) {
    Vb6OleCon* s = SITE_OF(This, lpVtblClientSite);
    if (!ppC) return E_POINTER;
    *ppC = NULL;
    s->ref++;
    *ppC = (IOleContainer*)&s->lpVtblContainer;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_CS_ShowObject(IOleClientSite* This) {
    Vb6OleCon* s = SITE_OF(This, lpVtblClientSite);
    olcRepaint(s);
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_CS_OnShowWindow(IOleClientSite* This, BOOL fShow) {
    Vb6OleCon* s = SITE_OF(This, lpVtblClientSite);
    if (s->hwnd) ShowWindow(s->hwnd, fShow ? SW_SHOW : SW_HIDE);
    return S_OK;
}

/* 对象改了大小要新布局 —— 按 SizeMode 决定是"控件随对象"还是"对象随控件" */
static HRESULT STDMETHODCALLTYPE olc_CS_RequestNewObjectLayout(IOleClientSite* This) {
    Vb6OleCon* s = SITE_OF(This, lpVtblClientSite);
    if (s->sizeMode == 2) olcSizeToObject(s);   /* 2 = 控件随对象 */
    olcRepaint(s);
    return S_OK;
}

static const IOleClientSiteVtbl g_olcClientSiteVtbl = {
    olc_CS_QueryInterface, olc_CS_AddRef, olc_CS_Release,
    olc_CS_SaveObject, olc_CS_GetMoniker, olc_CS_GetContainer,
    olc_CS_ShowObject, olc_CS_OnShowWindow, olc_CS_RequestNewObjectLayout
};

// ============================================================
// IOleInPlaceSite (+ IOleInPlaceFrame) —— 就地激活
// ============================================================

static HRESULT STDMETHODCALLTYPE olc_IPS_GetWindow(IOleInPlaceSite* This, HWND* phwnd) {
    Vb6OleCon* s = SITE_OF(This, lpVtblInPlaceSite);
    if (!phwnd) return E_POINTER;
    *phwnd = s->hwnd;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_ContextSensitiveHelp(IOleInPlaceSite* This, BOOL f) {
    (void)This; (void)f; return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_CanInPlaceActivate(IOleInPlaceSite* This) {
    (void)This;
    return S_OK;   /* 单对象容器: 永远允许就地激活 */
}

static HRESULT STDMETHODCALLTYPE olc_IPS_OnInPlaceActivate(IOleInPlaceSite* This) {
    Vb6OleCon* s = SITE_OF(This, lpVtblInPlaceSite);
    s->inPlaceActive = 1;
    OLC_LOG("[OLECON] OnInPlaceActivate\n");
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_OnUIActivate(IOleInPlaceSite* This) {
    (void)This; return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_GetWindowContext(IOleInPlaceSite* This,
        IOleInPlaceFrame** ppFrame, IOleInPlaceUIWindow** ppDoc, LPRECT rcPos, LPRECT rcClip,
        LPOLEINPLACEFRAMEINFO info) {
    Vb6OleCon* s = SITE_OF(This, lpVtblInPlaceSite);
    RECT rc;
    if (!s->hwnd) return E_UNEXPECTED;
    GetClientRect(s->hwnd, &rc);
    if (ppFrame) {
        s->ref++;
        *ppFrame = (IOleInPlaceFrame*)&s->lpVtblInPlaceFrame;
    }
    if (ppDoc) *ppDoc = NULL;          /* 单对象容器没有文档级 UI 窗口 */
    if (rcPos) *rcPos = rc;
    if (rcClip) *rcClip = rc;
    if (info) {
        /* 就地激活期间要转发的加速键 —— 容器不吞任何键, 全给对象 */
        info->cb = sizeof(OLEINPLACEFRAMEINFO);
        info->fMDIApp = FALSE;
        info->hwndFrame = GetAncestor(s->hwnd, GA_ROOT);
        info->haccel = NULL;
        info->cAccelEntries = 0;
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_Scroll(IOleInPlaceSite* This, SIZE sz) {
    (void)This; (void)sz; return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_OnUIDeactivate(IOleInPlaceSite* This, BOOL fUndoable) {
    (void)This; (void)fUndoable; return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_OnInPlaceDeactivate(IOleInPlaceSite* This) {
    Vb6OleCon* s = SITE_OF(This, lpVtblInPlaceSite);
    s->inPlaceActive = 0;
    OLC_LOG("[OLECON] OnInPlaceDeactivate\n");
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_DeactivateAndUndo(IOleInPlaceSite* This) {
    (void)This; return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IPS_OnPosRectChange(IOleInPlaceSite* This, LPCRECT prc) {
    Vb6OleCon* s = SITE_OF(This, lpVtblInPlaceSite);
    IOleInPlaceObject* pIPO = NULL;
    if (!prc) return E_POINTER;
    /* 单对象容器: 接受对象要求的矩形, 不裁剪 */
    if (s->pOleObj &&
        SUCCEEDED(s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IOleInPlaceObject, (void**)&pIPO)) &&
        pIPO) {
        pIPO->lpVtbl->SetObjectRects(pIPO, prc, prc);
        pIPO->lpVtbl->Release(pIPO);
    }
    return S_OK;
}

static const IOleInPlaceSiteVtbl g_olcInPlaceSiteVtbl = {
    olc_IPS_QueryInterface, olc_IPS_AddRef, olc_IPS_Release,
    olc_IPS_GetWindow, olc_IPS_ContextSensitiveHelp,
    olc_IPS_CanInPlaceActivate, olc_IPS_OnInPlaceActivate, olc_IPS_OnUIActivate,
    olc_IPS_GetWindowContext, olc_IPS_Scroll, olc_IPS_OnUIDeactivate,
    olc_IPS_OnInPlaceDeactivate, olc_IPS_DeactivateAndUndo, olc_IPS_OnPosRectChange
};

/* --- IOleInPlaceFrame --- */
static HRESULT STDMETHODCALLTYPE olc_IPF_GetWindow(IOleInPlaceFrame* This, HWND* p) {
    Vb6OleCon* s = SITE_OF(This, lpVtblInPlaceFrame);
    if (!p) return E_POINTER;
    *p = GetAncestor(s->hwnd, GA_ROOT);
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_ContextSensitiveHelp(IOleInPlaceFrame* T, BOOL f) {
    (void)T; (void)f; return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_GetBorder(IOleInPlaceFrame* T, LPRECT r) {
    (void)T; if (r) SetRectEmpty(r); return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_RequestBorderSpace(IOleInPlaceFrame* T, LPCBORDERWIDTHS b) {
    (void)T; (void)b; return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_SetBorderSpace(IOleInPlaceFrame* T, LPCBORDERWIDTHS b) {
    (void)T; (void)b; return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_SetActiveObject(IOleInPlaceFrame* T,
        IOleInPlaceActiveObject* o, LPCOLESTR n) { (void)T; (void)o; (void)n; return S_OK; }
/* InsertMenus / SetMenu: 就地激活时服务器要把自己的菜单并进容器菜单。
 * 本实现不做菜单合并 (§五-5 的"设计期/菜单类不做"), 只要返回 S_OK —— 返回失败
 * 会让部分服务器直接放弃激活。SetMenu(NULL) 时表示"恢复容器原菜单", 也 S_OK。 */
static HRESULT STDMETHODCALLTYPE olc_IPF_InsertMenus(IOleInPlaceFrame* T, HMENU m,
        LPOLEMENUGROUPWIDTHS w) { (void)T; (void)m; (void)w; return S_OK; }
static HRESULT STDMETHODCALLTYPE olc_IPF_SetMenu(IOleInPlaceFrame* T, HMENU m, HOLEMENU hm,
        HWND h) { (void)T; (void)m; (void)hm; (void)h; return S_OK; }
static HRESULT STDMETHODCALLTYPE olc_IPF_RemoveMenus(IOleInPlaceFrame* T, HMENU m) {
    (void)T; (void)m; return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_SetStatusText(IOleInPlaceFrame* T, LPCOLESTR t) {
    (void)T; (void)t; return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_EnableModeless(IOleInPlaceFrame* T, BOOL f) {
    (void)T; (void)f; return S_OK;
}
static HRESULT STDMETHODCALLTYPE olc_IPF_TranslateAccelerator(IOleInPlaceFrame* T, LPMSG m, WORD id) {
    (void)T; (void)m; (void)id;
    return S_FALSE;   /* S_FALSE = "我没处理" —— 对象自己接着处理 */
}

static const IOleInPlaceFrameVtbl g_olcInPlaceFrameVtbl = {
    olc_IPF_QueryInterface, olc_IPF_AddRef, olc_IPF_Release,
    olc_IPF_GetWindow, olc_IPF_ContextSensitiveHelp,
    olc_IPF_GetBorder, olc_IPF_RequestBorderSpace, olc_IPF_SetBorderSpace,
    olc_IPF_SetActiveObject, olc_IPF_InsertMenus, olc_IPF_SetMenu, olc_IPF_RemoveMenus,
    olc_IPF_SetStatusText, olc_IPF_EnableModeless, olc_IPF_TranslateAccelerator
};

// ============================================================
// IOleContainer / IOleItemContainer
// ============================================================

static HRESULT STDMETHODCALLTYPE olc_IC_EnumObjects(IOleContainer* This, DWORD flags,
                                                    IEnumUnknown** ppEnum) {
    (void)This; (void)flags;
    if (ppEnum) *ppEnum = NULL;
    /* 容器里只有被嵌入的那一个对象, 且本控件不对外暴露枚举 —— 返回空枚举比
     * E_NOTIMPL 更友好 (部分服务器拿不到枚举就认为自己不在容器里)。 */
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE olc_IC_LockContainer(IOleContainer* This, BOOL lock) {
    (void)This; (void)lock; return S_OK;
}

static HRESULT STDMETHODCALLTYPE olc_IC_ParseDisplayName(IOleContainer* This, IBindCtx* pbc,
        OLECHAR* name, ULONG* eaten, IMoniker** ppmk) {
    (void)This; (void)pbc;
    if (ppmk) *ppmk = NULL;
    if (eaten) *eaten = 0;
    if (name) return CreateFileMoniker(name, ppmk);
    return E_INVALIDARG;
}

static const IOleContainerVtbl g_olcContainerVtbl = {
    olc_IC_QueryInterface, olc_IC_AddRef, olc_IC_Release,
    olc_IC_EnumObjects, olc_IC_LockContainer, olc_IC_ParseDisplayName
};

static HRESULT STDMETHODCALLTYPE olc_IIC_GetObject(IOleItemContainer* This, LPOLESTR name,
        DWORD speed, IBindCtx* pbc, REFIID riid, void** ppv) {
    (void)This; (void)name; (void)speed; (void)pbc; (void)riid;
    if (ppv) *ppv = NULL;
    return MK_E_NOOBJECT;
}
static HRESULT STDMETHODCALLTYPE olc_IIC_IsRunning(IOleItemContainer* This, LPOLESTR name) {
    (void)This; (void)name; return S_FALSE;
}

static const IOleItemContainerVtbl g_olcItemContainerVtbl = {
    olc_IIC_QueryInterface, olc_IIC_AddRef, olc_IIC_Release,
    olc_IC_EnumObjects, olc_IC_LockContainer, olc_IC_ParseDisplayName,
    olc_IIC_GetObject, olc_IIC_IsRunning
};

// ============================================================
// 存储: 用 ILockBytes 的内存复合文档当对象的 IStorage
// ============================================================

static void olcCloseStorage(Vb6OleCon* s) {
    if (s->pStorage) { s->pStorage->lpVtbl->Release(s->pStorage); s->pStorage = NULL; }
    if (s->pLockBytes) { s->pLockBytes->lpVtbl->Release(s->pLockBytes); s->pLockBytes = NULL; }
}

/* 建一本新的空复合文档 (内存 backing)。对象要用它保存自己。 */
static HRESULT olcEnsureStorage(Vb6OleCon* s) {
    HRESULT hr;
    if (s->pStorage) return S_OK;
    hr = CreateILockBytesOnHGlobal(NULL, TRUE, &s->pLockBytes);
    if (FAILED(hr)) return hr;
    hr = StgCreateDocfileOnILockBytes(s->pLockBytes,
             STGM_CREATE | STGM_READWRITE | STGM_SHARE_EXCLUSIVE, 0, &s->pStorage);
    OLC_LOG("[OLECON] ensureStorage hr=0x%08lX pStg=%p\n", (unsigned long)hr, (void*)s->pStorage);
    return hr;
}

/* 把内存复合文档落盘 —— 走 ILockBytes::Flush 之后读回全局内存 */
static int olcStorageToFile(Vb6OleCon* s, const wchar_t* path) {
    HGLOBAL hG = NULL;
    SIZE_T cb = 0;
    void* pv = NULL;
    FILE* fp;
    if (!s->pStorage || !s->pLockBytes) return 0;
    s->pStorage->lpVtbl->Commit(s->pStorage, STGC_DEFAULT);
    if (FAILED(s->pLockBytes->lpVtbl->Flush(s->pLockBytes))) return 0;
    if (FAILED(GetHGlobalFromILockBytes(s->pLockBytes, &hG)) || !hG) return 0;
    cb = GlobalSize(hG);
    pv = GlobalLock(hG);
    if (!pv) return 0;
    fp = _wfopen(path, L"wb");
    if (!fp) { GlobalUnlock(hG); return 0; }
    fwrite(pv, 1, cb, fp);
    fclose(fp);
    GlobalUnlock(hG);
    OLC_LOG("[OLECON] storageToFile %ls cb=%llu\n", path, (unsigned long long)cb);
    return 1;
}

/* 从文件装载成内存复合文档 */
static HRESULT olcStorageFromFile(Vb6OleCon* s, const wchar_t* path) {
    HANDLE hf;
    DWORD cb, got;
    HGLOBAL hG;
    void* pv;
    HRESULT hr;
    olcCloseStorage(s);
    hf = CreateFileW(path, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING,
                     FILE_ATTRIBUTE_NORMAL, NULL);
    if (hf == INVALID_HANDLE_VALUE) return STG_E_FILENOTFOUND;
    cb = GetFileSize(hf, NULL);
    hG = GlobalAlloc(GMEM_MOVEABLE, cb ? cb : 1);
    if (!hG) { CloseHandle(hf); return E_OUTOFMEMORY; }
    pv = GlobalLock(hG);
    ReadFile(hf, pv, cb, &got, NULL);
    GlobalUnlock(hG);
    CloseHandle(hf);
    hr = CreateILockBytesOnHGlobal(hG, TRUE, &s->pLockBytes);
    if (FAILED(hr)) { GlobalFree(hG); return hr; }
    hr = StgOpenStorageOnILockBytes(s->pLockBytes, NULL,
             STGM_READWRITE | STGM_SHARE_EXCLUSIVE, NULL, 0, &s->pStorage);
    OLC_LOG("[OLECON] storageFromFile hr=0x%08lX\n", (unsigned long)hr);
    return hr;
}

// ============================================================
// 对象装载 / 释放
// ============================================================

/* 从对象上把我们要用的接口一次取全 */
static void olcBindInterfaces(Vb6OleCon* s) {
    if (!s->pOleObj) return;
    s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IViewObject2, (void**)&s->pViewObj);
    if (!s->pViewObj)
        s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IViewObject, (void**)&s->pViewObj);
    s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IDataObject, (void**)&s->pDataObj);
    s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IPersistStorage, (void**)&s->pPersist);
    s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IDispatch, (void**)&s->pDisp);
    /* 关键: 告知对象"你是被包含的" —— 不调它, 对象会把自己当顶层文档,
     * 表现为菜单/焦点/关闭语义全错 (Excel 尤其明显)。 */
    OleSetContainedObject((IUnknown*)s->pOleObj, TRUE);
    OLC_LOG("[OLECON] bind: view=%p data=%p persist=%p disp=%p\n",
            (void*)s->pViewObj, (void*)s->pDataObj, (void*)s->pPersist, (void*)s->pDisp);
}

static void olcReleaseObject(Vb6OleCon* s) {
    if (s->inPlaceActive && s->pOleObj) {
        IOleInPlaceObject* pIPO = NULL;
        if (SUCCEEDED(s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IOleInPlaceObject, (void**)&pIPO)) && pIPO) {
            pIPO->lpVtbl->InPlaceDeactivate(pIPO);
            pIPO->lpVtbl->Release(pIPO);
        }
        s->inPlaceActive = 0;
    }
    if (s->pDisp)     { s->pDisp->lpVtbl->Release(s->pDisp); s->pDisp = NULL; }
    if (s->pPersist)  { s->pPersist->lpVtbl->Release(s->pPersist); s->pPersist = NULL; }
    if (s->pDataObj)  { s->pDataObj->lpVtbl->Release(s->pDataObj); s->pDataObj = NULL; }
    if (s->pViewObj)  { s->pViewObj->lpVtbl->Release(s->pViewObj); s->pViewObj = NULL; }
    if (s->pOleObj)   {
        s->pOleObj->lpVtbl->SetClientSite(s->pOleObj, NULL);
        s->pOleObj->lpVtbl->Close(s->pOleObj, OLECLOSE_NOSAVE);
        s->pOleObj->lpVtbl->Release(s->pOleObj);
        s->pOleObj = NULL;
    }
    s->oletType = VB6_OLETYPE_NONE;
}

/* 就地激活; 失败退 OLEIVERB_SHOW; 再失败就只靠 IViewObject::Draw 静态绘制 */
static int olcActivate(Vb6OleCon* s, LONG verb) {
    HRESULT hr;
    if (!s->pOleObj || !s->hwnd) return 0;
    hr = s->pOleObj->lpVtbl->DoVerb(s->pOleObj, verb, NULL,
                                    (IOleClientSite*)&s->lpVtblClientSite,
                                    0, s->hwnd, NULL);
    OLC_LOG("[OLECON] DoVerb(%ld) hr=0x%08lX\n", verb, (unsigned long)hr);
    /* ⚠ 创建路径**不做** SHOW fallback: packager 这类"编辑器型"服务器,
     * OLEIVERB_SHOW = 打开完整编辑窗口 (模态等待) —— 无头/后台环境直接挂死。
     * in-place 激活失败 (如 packager 不支持, hr=E_NOTIMPL) 就退回容器自己
     * IViewObject::Draw 画静态内容。用户显式 DoVerb(OLEIVERB_SHOW) 时才真开。 */
    olcRepaint(s);
    return SUCCEEDED(hr);
}

/* 按对象固有尺寸调整控件 (SizeMode=2 "控件随对象")。
 * ⚠ IViewObject2::GetExtent 是 **5 个参数** (比 IViewObject 多 lindex 与 ptd) —— 少传
 * 一个 ptd 会编成 Draw 的参数表, 报 "参数太少" 这种看不出真因的错。 */
static void olcSizeToObject(Vb6OleCon* s) {
    SIZEL sz;
    if (!s->pViewObj || !s->hwnd) return;
    if (FAILED(s->pViewObj->lpVtbl->GetExtent(s->pViewObj, DVASPECT_CONTENT, -1, NULL, &sz))) return;
    if (sz.cx && sz.cy) {
        /* 账 #298 (§B129): 以前这两行走 HIMETRIC -> 缇 -> 像素 两步，第二步是写死 96 DPI 的
         * `/15`，而且两次截断。现在一次折到位（vb6_HimetricToPxX/Y 按真实 DPI）：实测差值
         * 只有"同值"或"+1 像素"两种，从不 >=2。 */
        SetWindowPos(s->hwnd, NULL, 0, 0,
                     vb6_HimetricToPxX(sz.cx), vb6_HimetricToPxY(sz.cy),
                     SWP_NOMOVE | SWP_NOZORDER | SWP_NOACTIVATE);
    }
}

static void olcRepaint(Vb6OleCon* s) {
    if (s && s->hwnd) InvalidateRect(s->hwnd, NULL, TRUE);
}

// ============================================================
// 装载入口（供属性/方法共用）
// ============================================================

/* 嵌入/链接的公共后半段: 建存储 → 绑接口 → 设 extent → 就地激活。
 * viaFile 为真走 OleCreateFromFile (源文件内容); 否则只按 CLSID 建新对象。 */
static int olcFinishCreate(Vb6OleCon* s, int isLink) {
    SIZEL sz;
    RECT rc;
    if (!s->hwnd || !s->pOleObj) return 0;
    s->oletType = isLink ? VB6_OLETYPE_LINKED : VB6_OLETYPE_EMBEDDED;

    olcEnsureStorage(s);
    olcBindInterfaces(s);

    /* 容器像素 -> HIMETRIC 走权威那两枚（px -> 缇 -> hm）。以前这里自己乘 15 再折算一次：
     * 那个 15 把 DPI 写死成 96，而 32 位在 px > 56364 就溢出（账 #298 §B129 并表）。
     * 注：这段注释里刻意不写那个常数 —— 哨兵 R2 是按行扫 `2540` 的。 */
    GetClientRect(s->hwnd, &rc);
    sz.cx = vb6_TwipsToHimetric(vb6_XToTwipX(rc.right - rc.left));
    sz.cy = vb6_TwipsToHimetric(vb6_YToTwipY(rc.bottom - rc.top));
    s->pOleObj->lpVtbl->SetExtent(s->pOleObj, DVASPECT_CONTENT, &sz);

    olcActivate(s, OLEIVERB_INPLACEACTIVATE);
    olcRepaint(s);
    return 1;
}

int vb6_OleCon_CreateEmbed(void* hwnd, const wchar_t* sourceDoc) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    HRESULT hr;
    if (!s) return 0;
    if (s->oletTypeAllowed == 0) return 0;      /* 0 = 仅链接 */
    olcReleaseObject(s);
    olcCloseStorage(s);
    hr = olcEnsureStorage(s);
    if (FAILED(hr)) return 0;

    if (sourceDoc && *sourceDoc) {
        /* 有源文件: 嵌入该文件的内容 (类由文件推断)。
         * ⚠ OleCreateFromFile 第一参是 REFCLSID (指针), CLSID_NULL 是聚合常量取不了址
         * —— 用一枚全零的局部 CLSID。 */
        CLSID nullCls;
        memset(&nullCls, 0, sizeof(nullCls));
        hr = OleCreateFromFile(&nullCls, sourceDoc, &IID_IOleObject,
                               OLERENDER_DRAW, NULL, (IOleClientSite*)&s->lpVtblClientSite,
                               s->pStorage, (void**)&s->pOleObj);
        wcsncpy(s->sourceDoc, sourceDoc, MAX_PATH - 1);
    } else if (s->className[0]) {
        CLSID clsid;
        hr = CLSIDFromProgID(s->className, &clsid);
        if (FAILED(hr)) { OLC_LOG("[OLECON] CLSIDFromProgID(%ls) failed\n", s->className); return 0; }
        hr = OleCreate(&clsid, &IID_IOleObject, OLERENDER_DRAW, NULL,
                       (IOleClientSite*)&s->lpVtblClientSite, s->pStorage, (void**)&s->pOleObj);
    } else {
        return 0;
    }
    OLC_LOG("[OLECON] CreateEmbed(%ls) hr=0x%08lX p=%p\n",
            sourceDoc ? sourceDoc : L"<class>", (unsigned long)hr, (void*)s->pOleObj);
    if (FAILED(hr) || !s->pOleObj) { s->pOleObj = NULL; return 0; }
    if (!s->pOleObj->lpVtbl->SetClientSite) return 0;
    s->pOleObj->lpVtbl->SetClientSite(s->pOleObj, (IOleClientSite*)&s->lpVtblClientSite);
    return olcFinishCreate(s, 0);
}

int vb6_OleCon_CreateLink(void* hwnd, const wchar_t* sourceDoc, const wchar_t* sourceItem) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    HRESULT hr;
    if (!s || !sourceDoc || !*sourceDoc) return 0;
    if (s->oletTypeAllowed == 1) return 0;      /* 1 = 仅嵌入 */
    olcReleaseObject(s);
    olcCloseStorage(s);
    hr = olcEnsureStorage(s);
    if (FAILED(hr)) return 0;

    hr = OleCreateLinkToFile(sourceDoc, &IID_IOleObject, OLERENDER_DRAW, NULL,
                             (IOleClientSite*)&s->lpVtblClientSite, s->pStorage,
                             (void**)&s->pOleObj);
    OLC_LOG("[OLECON] CreateLink(%ls) hr=0x%08lX p=%p\n",
            sourceDoc, (unsigned long)hr, (void*)s->pOleObj);
    if (FAILED(hr) || !s->pOleObj) { s->pOleObj = NULL; return 0; }
    wcsncpy(s->sourceDoc, sourceDoc, MAX_PATH - 1);
    if (sourceItem) wcsncpy(s->sourceItem, sourceItem, 127);
    s->pOleObj->lpVtbl->SetClientSite(s->pOleObj, (IOleClientSite*)&s->lpVtblClientSite);
    return olcFinishCreate(s, 1);
}

/* 从复合文档装载 (ReadFromFile / 设计期 Class 已存在时) */
int vb6_OleCon_ReadFromFile(void* hwnd, const wchar_t* path) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    HRESULT hr;
    if (!s || !path || !*path) return 0;
    olcReleaseObject(s);
    hr = olcStorageFromFile(s, path);
    if (FAILED(hr)) return 0;
    hr = OleLoad(s->pStorage, &IID_IOleObject, (IOleClientSite*)&s->lpVtblClientSite,
                 (void**)&s->pOleObj);
    OLC_LOG("[OLECON] OleLoad hr=0x%08lX p=%p\n", (unsigned long)hr, (void*)s->pOleObj);
    if (FAILED(hr) || !s->pOleObj) { s->pOleObj = NULL; return 0; }
    s->pOleObj->lpVtbl->SetClientSite(s->pOleObj, (IOleClientSite*)&s->lpVtblClientSite);
    return olcFinishCreate(s, 0);
}

int vb6_OleCon_SaveToFile(void* hwnd, const wchar_t* path) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    if (!s || !path || !*path) return 0;
    if (!s->pOleObj) return 0;
    /* 先让对象把自己写进我们的存储, 再落盘 */
    olc_CS_SaveObject((IOleClientSite*)&s->lpVtblClientSite);
    return olcStorageToFile(s, path);
}

int vb6_OleCon_DoVerb(void* hwnd, int verb) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    if (!s) return 0;
    return olcActivate(s, (LONG)verb);
}

int vb6_OleCon_Close(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    if (!s) return 0;
    olcReleaseObject(s);
    olcCloseStorage(s);
    olcRepaint(s);
    return 1;
}

void* vb6_OleCon_GetObject(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    if (!s || !s->pDisp) return NULL;
    s->pDisp->lpVtbl->AddRef(s->pDisp);
    return (void*)s->pDisp;
}

int vb6_OleCon_GetOleType(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    return s ? s->oletType : VB6_OLETYPE_NONE;
}

// ============================================================
// 剪贴板 (Copy / Paste) —— 复用对象自己的 IDataObject
// ============================================================

void vb6_OleCon_Copy(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    if (!s || !s->pDataObj) return;
    if (OleSetClipboard(s->pDataObj) == S_OK) OleFlushClipboard();
}

/* 从剪贴板粘一个 OLE 对象。
 * 第一刀只做 Copy —— Paste 要区分 CF_EMBEDDEDOBJECT / CF_LINKSOURCE 并走
 * OleCreateFromData, 与"普通数据变成静态图片"(OleCreateStaticFromData) 是两条语义,
 * 混在一起会静默把链接粘成图片。留待第二批。 */
int vb6_OleCon_Paste(void* hwnd) {
    (void)hwnd;
    return 0;
}

// ============================================================
// 系统对话框: InsertObjDlg
// 口径 (§六-3): 不新增 import lib —— oledlg 的函数用 LoadLibrary 现取。
// ============================================================

static void* oleconLoadOleDlg(const char* proc) {
    static HMODULE h = NULL;
    if (!h) h = LoadLibraryW(L"oledlg.dll");
    if (!h) return NULL;
    return (void*)GetProcAddress(h, proc);
}

/* 弹系统"插入对象"对话框, 用户选定后按选中的类建一个新嵌入对象。
 * 找不到 oledlg / 用户取消 → 返回 0 (VB6 的 InsertObjDlg 没有返回值, 0 只是内部信号)。 */
int vb6_OleCon_InsertObjDlg(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    typedef BOOL (WINAPI *PFN_Insert)(LPOLEUIINSERTOBJECTW);
    PFN_Insert fn;
    OLEUIINSERTOBJECTW d;
    wchar_t* buf;
    int ok = 0;

    if (!s) return 0;
    fn = (PFN_Insert)oleconLoadOleDlg("OleUIInsertObjectW");
    if (!fn) { OLC_LOG("[OLECON] oledlg.dll 不可用\n"); return 0; }

    memset(&d, 0, sizeof(d));
    d.cbStruct     = sizeof(d);
    d.hWndOwner    = (HWND)hwnd;
    d.lpszCaption  = L"插入对象";
    d.dwFlags      = IOF_SELECTCREATENEW | IOF_SELECTCREATEFROMFILE | IOF_SHOWHELP;
    buf = (wchar_t*)calloc(MAX_PATH, sizeof(wchar_t));
    d.lpszFile     = buf;
    d.cchFile      = MAX_PATH;

    if (!fn(&d)) { free(buf); return 0; }

    /* ① 用户选了"由文件创建" → 用文件路径嵌入 */
    if (buf && buf[0]) {
        ok = vb6_OleCon_CreateEmbed(hwnd, buf);
    } else if (d.clsid.Data1 || d.clsid.Data2 || d.clsid.Data3 || d.clsid.Data4[0]) {
        /* ② 选了"新建" → 记下类名再按类创建 (d.clsid 是 out 字段, 非指针) */
        LPOLESTR prog = NULL;
        if (SUCCEEDED(ProgIDFromCLSID(&d.clsid, &prog)) && prog) {
            wcsncpy(s->className, prog, 127);
            CoTaskMemFree(prog);
        }
        ok = vb6_OleCon_CreateEmbed(hwnd, NULL);
    }
    OLC_LOG("[OLECON] InsertObjDlg ok=%d file=%ls\n", ok, (buf && buf[0]) ? buf : L"");
    free(buf);
    return ok;
}

// ============================================================
// 窗口类
// ============================================================

/* 画"对象未建/无对象"时的样子 —— 空白带边框 / 斜纹, 与 VB6 一致 */
static void olcPaintEmpty(HWND hwnd, HDC hdc, Vb6OleCon* s) {
    RECT rc;
    HBRUSH br;
    GetClientRect(hwnd, &rc);
    br = CreateSolidBrush(s->borderStyle ? RGB(255, 255, 255) : GetSysColor(COLOR_BTNFACE));
    FillRect(hdc, &rc, br);
    DeleteObject(br);
    if (s->borderStyle) {
        HPEN pen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
        HPEN old = (HPEN)SelectObject(hdc, pen);
        HBRUSH ob = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));
        Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
        SelectObject(hdc, ob);
        SelectObject(hdc, old);
        DeleteObject(pen);
    }
}

/* 让对象把自己画到窗口 DC 上 —— 非就地激活时的显示路径 (SizeMode 决定目标矩形) */
static void olcPaintObject(HWND hwnd, HDC hdc, Vb6OleCon* s) {
    RECTL rc;
    GetClientRect(hwnd, &rc);
    if (!s->pViewObj) return;
    /* DVASPECT_ICON 时对象自己画图标; 否则画内容 */
    s->pViewObj->lpVtbl->Draw(s->pViewObj,
        (DWORD)(s->displayAsIcon ? DVASPECT_ICON : DVASPECT_CONTENT), -1,
        NULL, NULL, NULL, hdc, &rc, NULL, NULL, 0);
}

static LRESULT CALLBACK vb6_OleConWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    Vb6OleCon* s = olcFromHwnd(hwnd);

    switch (msg) {
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        if (!s) { EndPaint(hwnd, &ps); return 0; }
        if (s->borderStyle) {
            RECT rc; GetClientRect(hwnd, &rc);
            HBRUSH br = CreateSolidBrush(RGB(255, 255, 255));
            FillRect(hdc, &rc, br);
            DeleteObject(br);
        }
        if (s->pViewObj && !s->inPlaceActive) olcPaintObject(hwnd, hdc, s);
        else if (!s->pViewObj) olcPaintEmpty(hwnd, hdc, s);
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;    /* 全部在 WM_PAINT 里画, 避免闪烁 */
    case WM_SIZE: {
        SIZEL sz;
        if (!s || !s->pOleObj) break;
        sz.cx = vb6_TwipsToHimetric(vb6_XToTwipX((int)LOWORD(lp)));
        sz.cy = vb6_TwipsToHimetric(vb6_YToTwipY((int)HIWORD(lp)));
        if (sz.cx > 0 && sz.cy > 0) s->pOleObj->lpVtbl->SetExtent(s->pOleObj, DVASPECT_CONTENT, &sz);
        if (s->inPlaceActive) {
            IOleInPlaceObject* pIPO = NULL;
            RECT rc; GetClientRect(hwnd, &rc);
            if (SUCCEEDED(s->pOleObj->lpVtbl->QueryInterface(s->pOleObj, &IID_IOleInPlaceObject, (void**)&pIPO)) && pIPO) {
                pIPO->lpVtbl->SetObjectRects(pIPO, &rc, &rc);
                pIPO->lpVtbl->Release(pIPO);
            }
        }
        return 0;
    }
    case WM_SETFOCUS:
        if (s && s->autoActivate == 1) olcActivate(s, OLEIVERB_UIACTIVATE);   /* 1 = 获焦点 */
        break;
    case WM_LBUTTONDOWN:
        if (s && (s->autoActivate == 3)) olcActivate(s, OLEIVERB_PRIMARY);    /* 3 = 单击 */
        break;
    case WM_LBUTTONDBLCLK: {
        /* 2 = 双击 (默认): 执行主动词 —— 就是"进对象自己的编辑界面" */
        if (s && s->autoActivate == 2) olcActivate(s, OLEIVERB_PRIMARY);
        return 0;
    }
    case WM_SETCURSOR:
        if (s && s->pOleObj) { SetCursor(LoadCursorW(NULL, IDC_ARROW)); return TRUE; }
        break;
    case WM_DESTROY:
        if (s) {
            olcReleaseObject(s);
            olcCloseStorage(s);
            free(s);
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        }
        return 0;
    default:
        break;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

int vb6_RegisterOleConClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return 1;
    done = 1;
    memset(&wc, 0, sizeof(wc));
    wc.lpfnWndProc   = vb6_OleConWndProc;
    wc.hInstance     = (HINSTANCE)hInstance;
    wc.hCursor       = LoadCursorW(NULL, IDC_ARROW);
    wc.hbrBackground = NULL;
    wc.lpszClassName = VB6_OLECON_CLASS;
    if (!RegisterClassW(&wc)) return 0;
    return 1;
}

// ============================================================
// 初始化（cgen 在设计期属性之后调一次）
// ============================================================

void vb6_OleCon_Init(void* hwnd, const wchar_t* cls, int oletTypeAllowed, int sizeMode,
                     int displayAsIcon, int autoActivate, int autoVerbMenu,
                     int borderStyle, const wchar_t* sourceDoc, const wchar_t* sourceItem) {
    Vb6OleCon* s;
    HWND h = (HWND)hwnd;
    if (!h) return;

    s = (Vb6OleCon*)calloc(1, sizeof(Vb6OleCon));
    if (!s) return;
    s->lpVtblClientSite        = &g_olcClientSiteVtbl;
    s->lpVtblInPlaceSite       = &g_olcInPlaceSiteVtbl;
    s->lpVtblInPlaceFrame      = &g_olcInPlaceFrameVtbl;
    s->lpVtblContainer         = &g_olcContainerVtbl;
    s->lpVtblItemContainer     = &g_olcItemContainerVtbl;
    s->ref = 1;
    s->hwnd = h;
    s->oletType = VB6_OLETYPE_NONE;
    s->oletTypeAllowed = (oletTypeAllowed >= 0 && oletTypeAllowed <= 2) ? oletTypeAllowed : 2;
    s->sizeMode = sizeMode;
    s->displayAsIcon = displayAsIcon;
    s->autoActivate = autoActivate;
    s->autoVerbMenu = autoVerbMenu;
    s->borderStyle = borderStyle;
    if (cls) wcsncpy(s->className, cls, 127);
    if (sourceDoc) wcsncpy(s->sourceDoc, sourceDoc, MAX_PATH - 1);
    if (sourceItem) wcsncpy(s->sourceItem, sourceItem, 127);
    SetWindowLongPtrW(h, GWLP_USERDATA, (LONG_PTR)s);

    OLC_LOG("[OLECON] Init hwnd=%p cls=%ls allowed=%d sizeMode=%d docs=%ls\n",
            (void*)h, s->className, s->oletTypeAllowed, s->sizeMode, s->sourceDoc);
}

// ============================================================
// 属性 get/set —— 供 cgen 的属性表挂接
// ============================================================

#define OLC_GETINT(field, dflt)  do { Vb6OleCon* s = olcFromHwnd((HWND)hwnd); return s ? s->field : (dflt); } while (0)
#define OLC_SETINT(field)        do { Vb6OleCon* s = olcFromHwnd((HWND)hwnd); if (s) { s->field = v; olcRepaint(s); } } while (0)

int  vb6_OleCon_GetOLETypeAllowed(void* hwnd) { OLC_GETINT(oletTypeAllowed, 2); }
void vb6_OleCon_SetOLETypeAllowed(void* hwnd, int v) { OLC_SETINT(oletTypeAllowed); }
int  vb6_OleCon_GetSizeMode(void* hwnd) { OLC_GETINT(sizeMode, 0); }
void vb6_OleCon_SetSizeMode(void* hwnd, int v) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    if (!s) return;
    s->sizeMode = v;
    if (v == 2) olcSizeToObject(s);      /* 控件随对象 */
    olcRepaint(s);
}
int  vb6_OleCon_GetDisplayAsIcon(void* hwnd) { OLC_GETINT(displayAsIcon, 0); }
void vb6_OleCon_SetDisplayAsIcon(void* hwnd, int v) { OLC_SETINT(displayAsIcon); }
int  vb6_OleCon_GetAutoActivate(void* hwnd) { OLC_GETINT(autoActivate, 2); }
void vb6_OleCon_SetAutoActivate(void* hwnd, int v) { OLC_SETINT(autoActivate); }
int  vb6_OleCon_GetAutoVerbMenu(void* hwnd) { OLC_GETINT(autoVerbMenu, 1); }
void vb6_OleCon_SetAutoVerbMenu(void* hwnd, int v) { OLC_SETINT(autoVerbMenu); }
int  vb6_OleCon_GetBorderStyle(void* hwnd) { OLC_GETINT(borderStyle, 0); }
void vb6_OleCon_SetBorderStyle(void* hwnd, int v) { OLC_SETINT(borderStyle); }

wchar_t* vb6_OleCon_GetClass(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    return vb6_BSTR_FromStr(s ? s->className : L"");
}

wchar_t* vb6_OleCon_GetSourceDoc(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    return vb6_BSTR_FromStr(s ? s->sourceDoc : L"");
}

wchar_t* vb6_OleCon_GetSourceItem(void* hwnd) {
    Vb6OleCon* s = olcFromHwnd((HWND)hwnd);
    return vb6_BSTR_FromStr(s ? s->sourceItem : L"");
}
