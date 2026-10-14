// vb6forms_oledd.c — VB6 OLE 拖放 (P20-44)
//
// 三套 COM 实现:
//   IDataObject   拖动数据的载体 (VB6 的 DataObject)
//   IDropSource   拖动源头 (vb6_OLEDrag_Start → DoDragDrop)
//   IDropTarget   放置目标 (vb6_OLEDrop_Register → RegisterDragDrop)
//
// 事件回 VB: RTL 不认识窗体名, 用回调表注册。
//   vb6_OLEDrop_SetHandler(hwnd, kind, cb)   kind: 0=OLEDragOver 1=OLEDragDrop
//   回调签名 void cb(void* dataObj, int32_t* effect, int16_t* button,
//                    int16_t* shift, float* x, float* y)
//   dataObj 是 Vb6DataObject* (不透明句柄), 供 Data.GetText 等取用。
//
// DataObject 支持的格式: CF_TEXT / CF_UNICODETEXT / CF_HDROP (文件清单)。
//
// 自测: 环境变量 C3_OLEDDB_TEST=1 时, vb6_Init 之后由夹具调
//   vb6_oleDD_SelfTest(const wchar_t* outPath)   <- 路径走宽 (多语言路径), 见下
// 它建一个哑窗口、注册目标、直接调 IDropTarget 的 DragEnter/Drop (绕开
// DoDragDrop 的模态循环 —— 无头环境没法真拖), 把结果写进 outPath。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
// **DROPFILES / CF_HDROP 在 shlobj.h, 不在 shellapi.h**(SDK 10.0.19041.0 实测:
// 只 include shellapi.h 会报 DROPFILES 未声明)。shlobj.h 拖的东西多, 但 RTL 只编一次。
#include <shlobj.h>
#include <ole2.h>
#include <stdio.h>
#endif

#include "vb6forms.h"
#include "vb6rtl_bstr.h"   // BSTR 副本用 (static inline; RTL 全部平铺解包, 同 imagelist/sstab)
#include "vb6forms_internal.h"

#ifdef _WIN32

// ===================== DataObject 数据 =====================

typedef struct {
    wchar_t* text;          // CF_UNICODETEXT
    wchar_t** files;        // CF_HDROP 的文件清单
    int fileCount;
    DWORD allowedEffects;   // 1=Copy 2=Move 4=Link
    DWORD finalEffect;      // DoDragDrop 结束后的实际效果
} Vb6DataObj;

static void dataObjFree(Vb6DataObj* d) {
    if (!d) return;
    if (d->text) HeapFree(GetProcessHeap(), 0, d->text);
    for (int i = 0; i < d->fileCount; i++) HeapFree(GetProcessHeap(), 0, d->files[i]);
    if (d->files) HeapFree(GetProcessHeap(), 0, d->files);
    HeapFree(GetProcessHeap(), 0, d);
}

// ===================== IDataObject =====================

// FormatEtc 枚举器: 我们支持的格式是固定的三个, 枚举器只需要顺序吐出来。
typedef struct {
    IEnumFORMATETC vt;
    LONG refs;
    int idx;
} Vb6FmtEnum;

static const FORMATETC kFormats43[] = {
    { CF_TEXT,        NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL },
    { CF_UNICODETEXT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL },
    { CF_HDROP,       NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL },
};
#define VB6_NFMT 3

static HRESULT WINAPI fmtEnum_QueryInterface(IEnumFORMATETC* self, REFIID riid, void** out) {
    if (!out) return E_POINTER;
    *out = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IEnumFORMATETC)) {
        *out = self; self->lpVtbl->AddRef(self); return S_OK;
    }
    return E_NOINTERFACE;
}
static ULONG WINAPI fmtEnum_AddRef(IEnumFORMATETC* self) {
    Vb6FmtEnum* e = (Vb6FmtEnum*)self; return (ULONG)InterlockedIncrement(&e->refs);
}
static ULONG WINAPI fmtEnum_Release(IEnumFORMATETC* self) {
    Vb6FmtEnum* e = (Vb6FmtEnum*)self;
    LONG r = InterlockedDecrement(&e->refs);
    if (r == 0) HeapFree(GetProcessHeap(), 0, e);
    return (ULONG)r;
}
static HRESULT WINAPI fmtEnum_Next(IEnumFORMATETC* self, ULONG n, FORMATETC* out, ULONG* got) {
    Vb6FmtEnum* e = (Vb6FmtEnum*)self;
    if (got) *got = 0;
    ULONG c = 0;
    while (c < n && e->idx < VB6_NFMT) { out[c++] = kFormats43[e->idx++]; }
    if (got) *got = c;
    return (c == n) ? S_OK : S_FALSE;
}
static HRESULT WINAPI fmtEnum_Skip(IEnumFORMATETC* self, ULONG n) {
    Vb6FmtEnum* e = (Vb6FmtEnum*)self; e->idx += (int)n; return S_OK;
}
static HRESULT WINAPI fmtEnum_Reset(IEnumFORMATETC* self) {
    Vb6FmtEnum* e = (Vb6FmtEnum*)self; e->idx = 0; return S_OK;
}
static HRESULT WINAPI fmtEnum_Clone(IEnumFORMATETC* self, IEnumFORMATETC** out) {
    if (!out) return E_POINTER;
    Vb6FmtEnum* e = (Vb6FmtEnum*)self;
    Vb6FmtEnum* n = (Vb6FmtEnum*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6FmtEnum));
    if (!n) return E_OUTOFMEMORY;
    n->vt = e->vt; n->idx = e->idx;
    *out = &n->vt; return S_OK;
}

static IEnumFORMATETCVtbl kFmtEnumVtbl = {
    fmtEnum_QueryInterface, fmtEnum_AddRef, fmtEnum_Release,
    fmtEnum_Next, fmtEnum_Skip, fmtEnum_Reset, fmtEnum_Clone
};

typedef struct {
    IDataObject vt;
    LONG refs;
    Vb6DataObj* d;
} Vb6IDataObj;

static HRESULT WINAPI do_QueryInterface(IDataObject* self, REFIID riid, void** out) {
    if (!out) return E_POINTER;
    *out = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDataObject)) {
        *out = self; self->lpVtbl->AddRef(self); return S_OK;
    }
    return E_NOINTERFACE;
}
static ULONG WINAPI do_AddRef(IDataObject* self) {
    Vb6IDataObj* o = (Vb6IDataObj*)self; return (ULONG)InterlockedIncrement(&o->refs);
}
static ULONG WINAPI do_Release(IDataObject* self) {
    Vb6IDataObj* o = (Vb6IDataObj*)self;
    LONG r = InterlockedDecrement(&o->refs);
    if (r == 0) { dataObjFree(o->d); HeapFree(GetProcessHeap(), 0, o); }
    return (ULONG)r;
}
static HRESULT WINAPI do_GetData(IDataObject* self, FORMATETC* fmt, STGMEDIUM* med) {
    Vb6IDataObj* o = (Vb6IDataObj*)self;
    if (!fmt || !med) return E_POINTER;
    ZeroMemory(med, sizeof(*med));
    med->tymed = TYMED_HGLOBAL;

    if (fmt->cfFormat == CF_TEXT || fmt->cfFormat == CF_UNICODETEXT) {
        if (!o->d->text) return DV_E_FORMATETC;
        int chars = (int)lstrlenW(o->d->text);
        // CF_TEXT 要求 ANSI: 统一 WideCharToMultiByte(ACP)
        if (fmt->cfFormat == CF_TEXT) {
            int need = WideCharToMultiByte(CP_ACP, 0, o->d->text, chars, NULL, 0, NULL, NULL);
            HGLOBAL g = GlobalAlloc(GMEM_MOVEABLE, need + 1);
            if (!g) return STG_E_MEDIUMFULL;
            char* p = (char*)GlobalLock(g);
            if (!p) { GlobalFree(g); return STG_E_MEDIUMFULL; }
            WideCharToMultiByte(CP_ACP, 0, o->d->text, chars, p, need, NULL, NULL);
            p[need] = 0; GlobalUnlock(g);
            med->hGlobal = g; return S_OK;
        }
        HGLOBAL g = GlobalAlloc(GMEM_MOVEABLE, sizeof(wchar_t) * (chars + 1));
        if (!g) return STG_E_MEDIUMFULL;
        wchar_t* p = (wchar_t*)GlobalLock(g);
        if (!p) { GlobalFree(g); return STG_E_MEDIUMFULL; }
        memcpy(p, o->d->text, sizeof(wchar_t) * chars);
        p[chars] = 0; GlobalUnlock(g);
        med->hGlobal = g; return S_OK;
    }
    if (fmt->cfFormat == CF_HDROP) {
        if (!o->d->files || o->d->fileCount == 0) return DV_E_FORMATETC;
        int need = sizeof(DROPFILES);
        for (int i = 0; i < o->d->fileCount; i++) {
            if (!o->d->files[i]) continue;
            need += (int)sizeof(wchar_t) * (lstrlenW(o->d->files[i]) + 1);
        }
        need += sizeof(wchar_t);   // 结尾双 0
        HGLOBAL g = GlobalAlloc(GMEM_MOVEABLE, need);
        if (!g) return STG_E_MEDIUMFULL;
        DROPFILES* df = (DROPFILES*)GlobalLock(g);
        if (!df) { GlobalFree(g); return STG_E_MEDIUMFULL; }
        ZeroMemory(df, sizeof(*df));
        df->pFiles = sizeof(DROPFILES);
        df->fWide = TRUE;
        wchar_t* w = (wchar_t*)((char*)df + sizeof(DROPFILES));
        for (int i = 0; i < o->d->fileCount; i++) {
            if (!o->d->files[i]) continue;
            int n = lstrlenW(o->d->files[i]) + 1;
            memcpy(w, o->d->files[i], sizeof(wchar_t) * n);
            w += n;
        }
        *w = 0;
        GlobalUnlock(g);
        med->hGlobal = g; return S_OK;
    }
    return DV_E_FORMATETC;
}
static HRESULT WINAPI do_GetDataHere(IDataObject* self, FORMATETC* f, STGMEDIUM* m) { (void)self; (void)f; (void)m; return E_NOTIMPL; }
static HRESULT WINAPI do_QueryGetData(IDataObject* self, FORMATETC* fmt) {
    Vb6IDataObj* o = (Vb6IDataObj*)self;
    if (!fmt) return E_POINTER;
    if (fmt->cfFormat == CF_TEXT || fmt->cfFormat == CF_UNICODETEXT)
        return o->d->text ? S_OK : S_FALSE;
    if (fmt->cfFormat == CF_HDROP)
        return (o->d->files && o->d->fileCount) ? S_OK : S_FALSE;
    return S_FALSE;
}
static HRESULT WINAPI do_GetCanonicalFormatEtc(IDataObject* self, FORMATETC* in, FORMATETC* out) {
    if (!out) return E_POINTER; *out = *in; return DATA_S_SAMEFORMATETC; (void)self; (void)in;
}
static HRESULT WINAPI do_SetData(IDataObject* self, FORMATETC* f, STGMEDIUM* m, BOOL release) {
    (void)self; (void)f; (void)m; (void)release; return E_NOTIMPL;
}
static HRESULT WINAPI do_EnumFormatEtc(IDataObject* self, DWORD dir, IEnumFORMATETC** out) {
    if (!out) return E_POINTER;
    *out = NULL;
    if (dir != DATADIR_GET) return E_NOTIMPL;
    Vb6FmtEnum* e = (Vb6FmtEnum*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6FmtEnum));
    if (!e) return E_OUTOFMEMORY;
    e->vt.lpVtbl = &kFmtEnumVtbl;
    *out = &e->vt;
    fmtEnum_AddRef(&e->vt);
    return S_OK;
}
static HRESULT WINAPI do_DAdvise(IDataObject* self, FORMATETC* f, DWORD flags, IAdviseSink* sink, DWORD* conn) {
    (void)self; (void)f; (void)flags; (void)sink; (void)conn; return OLE_E_ADVISENOTSUPPORTED;
}
static HRESULT WINAPI do_DUnadvise(IDataObject* self, DWORD conn) { (void)self; (void)conn; return OLE_E_ADVISENOTSUPPORTED; }
static HRESULT WINAPI do_EnumDAdvise(IDataObject* self, IEnumSTATDATA** out) { (void)self; (void)out; return OLE_E_ADVISENOTSUPPORTED; }

static IDataObjectVtbl kDataObjVtbl = {
    do_QueryInterface, do_AddRef, do_Release,
    do_GetData, do_GetDataHere, do_QueryGetData, do_GetCanonicalFormatEtc,
    do_SetData, do_EnumFormatEtc, do_DAdvise, do_DUnadvise, do_EnumDAdvise
};

// 造一个填好文本/文件的 IDataObject。失败返回 NULL。
// 返回 **void\*** (接口指针的裸表示): 统一 RTL 口径, 头文件不必 include ole2.h。
// ⚠ 必须在 vb6forms_prop_ctrl.h 里有原型 —— 漏了就 C 隐式 int, x64 把指针截成
// 32 位 (实测 dataObj=0xFFFFFFFF90FE65A0 → SetText 里解引用崩)。
void* vb6_oleDD_MakeDataObject(const wchar_t* text, wchar_t** files, int fileCount) {
    Vb6DataObj* d = (Vb6DataObj*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6DataObj));
    if (!d) return NULL;
    if (text && *text) {
        int n = lstrlenW(text);
        d->text = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, sizeof(wchar_t) * (n + 1));
        if (d->text) { memcpy(d->text, text, sizeof(wchar_t) * n); d->text[n] = 0; }
    }
    if (files && fileCount > 0) {
        d->files = (wchar_t**)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(wchar_t*) * fileCount);
        if (d->files) {
            for (int i = 0; i < fileCount; i++) {
                int n = lstrlenW(files[i]);
                d->files[i] = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, sizeof(wchar_t) * (n + 1));
                if (d->files[i]) { memcpy(d->files[i], files[i], sizeof(wchar_t) * n); d->files[i][n] = 0; }
                d->fileCount++;
            }
        }
    }
    d->allowedEffects = 1 /*Copy*/ | 2 /*Move*/;
    Vb6IDataObj* o = (Vb6IDataObj*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6IDataObj));
    if (!o) { dataObjFree(d); return NULL; }
    o->vt.lpVtbl = &kDataObjVtbl;
    o->d = d;
    return (void*)&o->vt;
}

// 取出 DataObject 里的文本 (给 VB 的 Data.GetText 用)
// 返回 **BSTR 副本**: VB 侧拿到的是独立串, 不会因为 DataObject 释放而悬垂。
// (签名是 void*: 与 RTL 其它取值函数同口径, 记进 kBstrReturningCalls。)
void* vb6_oleDD_GetText(IDataObject* self) {
    Vb6IDataObj* o = (Vb6IDataObj*)self;
    if (!o || !o->d->text) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(o->d->text);
}

int32_t vb6_oleDD_GetFileCount(IDataObject* self) {
    Vb6IDataObj* o = (Vb6IDataObj*)self;
    return o ? (int32_t)o->d->fileCount : 0;
}

const wchar_t* vb6_oleDD_GetFile(IDataObject* self, int32_t i) {
    Vb6IDataObj* o = (Vb6IDataObj*)self;
    if (!o || i < 0 || i >= o->d->fileCount) return L"";
    return o->d->files[i];
}

// ===================== IDropSource =====================

typedef struct {
    IDropSource vt;
    LONG refs;
} Vb6DropSrc;

static HRESULT WINAPI ds_QueryInterface(IDropSource* self, REFIID riid, void** out) {
    if (!out) return E_POINTER; *out = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDropSource)) {
        *out = self; self->lpVtbl->AddRef(self); return S_OK;
    }
    return E_NOINTERFACE;
}
static ULONG WINAPI ds_AddRef(IDropSource* self) {
    Vb6DropSrc* s = (Vb6DropSrc*)self; return (ULONG)InterlockedIncrement(&s->refs);
}
static ULONG WINAPI ds_Release(IDropSource* self) {
    Vb6DropSrc* s = (Vb6DropSrc*)self;
    LONG r = InterlockedDecrement(&s->refs);
    if (r == 0) HeapFree(GetProcessHeap(), 0, s);
    return (ULONG)r;
}
static HRESULT WINAPI ds_QueryContinueDrag(IDropSource* self, BOOL esc, DWORD state) {
    (void)self;
    if (esc) return DRAGDROP_S_CANCEL;
    if (!(state & (MK_LBUTTON | MK_RBUTTON))) return DRAGDROP_S_DROP;
    return S_OK;
}
static HRESULT WINAPI ds_GiveFeedback(IDropSource* self, DWORD eff) {
    (void)self; (void)eff; return DRAGDROP_S_USEDEFAULTCURSORS;
}

static HRESULT WINAPI src_GiveFeedbackCb(IDropSource* self, DWORD eff);  // 定义在目标回调表之后

static IDropSourceVtbl kDropSrcVtbl = {
    ds_QueryInterface, ds_AddRef, ds_Release, ds_QueryContinueDrag, src_GiveFeedbackCb
};

// ===================== 目标: 事件回调表 =====================

// 首参是 **void\*** (指向 DataObject 指针的指针): C3 给
// `Sub X_OLEDragDrop(Data As DataObject, ...)` 生成的 C 签名首参就是 `void** Data`,
// 这样 RTL 可以把处理器**直接注册成回调**, 不需要 thunk。
typedef void (*Vb6OLEDropCb)(void** dataObj, int32_t* effect, int16_t* button,
                             int16_t* shift, float* x, float* y);
#define VB6_OLECB_OVER 0
#define VB6_OLECB_DROP 1
// P20-44 源侧: 处理器签名各不相同, 所以存 void* 按各自形状 cast 调用。
#define VB6_OLESRC_STARTDRAG  2   // void X(void** Data, int32_t* AllowedEffects)
#define VB6_OLESRC_SETDATA    3   // void X(void** Data, int16_t* DataFormat)
#define VB6_OLESRC_GIVEFEEDBK 4   // void X(int32_t* Effect, int16_t* Button, int16_t* Shift, float* X, float* Y)
#define VB6_OLESRC_COMPLETED  5   // void X(int32_t* Effect)

#define VB6_OLETARGET_MAX 32
typedef struct {
    HWND hwnd;
    Vb6OLEDropCb cb[2];
    void* srcCb[4];        // 源侧: [0]=StartDrag [1]=SetData [2]=GiveFeedback [3]=CompleteDrag
} Vb6OLETarget;

static Vb6OLETarget g_targets[VB6_OLETARGET_MAX];
static int g_targetCount = 0;

static Vb6OLETarget* targetFind(HWND h) {
    for (int i = 0; i < g_targetCount; i++) if (g_targets[i].hwnd == h) return &g_targets[i];
    return NULL;
}

void vb6_OLEDrop_SetHandler(void* hwnd, int32_t kind, void* cb) {
    if (!hwnd || kind < 0 || kind > 5) return;
    Vb6OLETarget* t = targetFind((HWND)hwnd);
    if (!t) {
        if (g_targetCount >= VB6_OLETARGET_MAX) return;
        t = &g_targets[g_targetCount++];
        t->hwnd = (HWND)hwnd;
    }
    if (kind <= 1) t->cb[kind] = (Vb6OLEDropCb)cb;
    else t->srcCb[kind - 2] = cb;   // 源侧四事件 (签名各异, 调用点按形状 cast)
}

// ===================== 源侧: OLEDrag / DoDragDrop (P20-44) =====================
// VB6 用法: 通常在 MouseMove 里手动调 `X.OLEDrag` (或 OLEDragMode=1 自动)。
// 这里封装 DoDragDrop 的模态循环; 源事件在循环前后/中回调:
//   OLEStartDrag(Data, AllowedEffects)  循环前 (可往 Data 里塞数据)
//   OLEGiveFeedback(Effect, ...)        循环中每次 GiveFeedback
//   OLECompleteDrag(Effect)             循环后 (Effect = 最终效果, 0=未落下)
static HRESULT g_srcPendingGive = S_OK;
static HWND g_srcActiveHwnd = NULL;          // 当前拖动的源控件
static IDataObject* g_srcDataObj = NULL;     // 当前拖动的数据对象 (引用在调用方)

static HRESULT WINAPI src_GiveFeedbackCb(IDropSource* self, DWORD eff) {
    (void)self;
    if (g_srcActiveHwnd && g_srcDataObj) {
        Vb6OLETarget* t = targetFind(g_srcActiveHwnd);
        if (t && t->srcCb[2]) {
            int32_t e = (int32_t)eff; int16_t b2 = 0, s2 = 0; float x = 0, y = 0;
            void* objPtr = (void*)g_srcDataObj;
            ((void(*)(int32_t*, int16_t*, int16_t*, float*, float*))t->srcCb[2])
                (&e, &b2, &s2, &x, &y);
        }
    }
    return DRAGDROP_S_USEDEFAULTCURSORS;
}


// ===================== IDropTarget =====================

typedef struct {
    IDropTarget vt;
    LONG refs;
    HWND hwnd;
    DWORD lastEffect;
    Vb6DataObj* curData;   // DragEnter 进来的对象 (引用计数在 IDataObject 上)
    IDataObject* curObj;
} Vb6DropTarget;

static void targetFire(HWND hwnd, int kind, IDataObject* obj, DWORD* effect, POINTL pt) {
    Vb6OLETarget* t = targetFind(hwnd);
    if (!t || !t->cb[kind]) return;
    int32_t e = (int32_t)(effect ? *effect : 0);
    int16_t btn = 0, shf = 0;
    // 键修饰: 拖放时的 Shift/Alt/Ctrl 由 DoDragDrop 的键盘状态决定
    if (GetKeyState(VK_SHIFT) & 0x8000) shf |= 1;
    if (GetKeyState(VK_CONTROL) & 0x8000) shf |= 2;
    if (GetKeyState(VK_MENU) & 0x8000) shf |= 4;
    // 坐标: 屏幕坐标 → 客户区 (VB6 事件给的是目标控件的客户区坐标)
    POINT p = { (LONG)pt.x, (LONG)pt.y };
    ScreenToClient(hwnd, &p);
    float fx = (float)p.x, fy = (float)p.y;
    void* objPtr43 = (void*)obj;
    t->cb[kind](&objPtr43, &e, &btn, &shf, &fx, &fy);
    if (effect) *effect = (DWORD)e;
}

static HRESULT WINAPI dt_QueryInterface(IDropTarget* self, REFIID riid, void** out) {
    if (!out) return E_POINTER; *out = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDropTarget)) {
        *out = self; self->lpVtbl->AddRef(self); return S_OK;
    }
    return E_NOINTERFACE;
}
static ULONG WINAPI dt_AddRef(IDropTarget* self) {
    Vb6DropTarget* t = (Vb6DropTarget*)self; return (ULONG)InterlockedIncrement(&t->refs);
}
static ULONG WINAPI dt_Release(IDropTarget* self) {
    Vb6DropTarget* t = (Vb6DropTarget*)self;
    LONG r = InterlockedDecrement(&t->refs);
    if (r == 0) HeapFree(GetProcessHeap(), 0, t);
    return (ULONG)r;
}
static HRESULT WINAPI dt_DragEnter(IDropTarget* self, IDataObject* obj, DWORD keys, POINTL pt, DWORD* effect) {
    Vb6DropTarget* t = (Vb6DropTarget*)self;
    t->curObj = obj; if (obj) obj->lpVtbl->AddRef(obj);
    if (effect) *effect = DROPEFFECT_COPY;
    targetFire(t->hwnd, VB6_OLECB_OVER, obj, effect, pt);
    return S_OK;
}
static HRESULT WINAPI dt_DragOver(IDropTarget* self, DWORD keys, POINTL pt, DWORD* effect) {
    Vb6DropTarget* t = (Vb6DropTarget*)self;
    if (effect) *effect = DROPEFFECT_COPY;
    targetFire(t->hwnd, VB6_OLECB_OVER, t->curObj, effect, pt);
    return S_OK;
}
static HRESULT WINAPI dt_DragLeave(IDropTarget* self) {
    Vb6DropTarget* t = (Vb6DropTarget*)self;
    if (t->curObj) { t->curObj->lpVtbl->Release(t->curObj); t->curObj = NULL; }
    return S_OK;
}
static HRESULT WINAPI dt_Drop(IDropTarget* self, IDataObject* obj, DWORD keys, POINTL pt, DWORD* effect) {
    Vb6DropTarget* t = (Vb6DropTarget*)self;
    if (effect) *effect = DROPEFFECT_COPY;
    targetFire(t->hwnd, VB6_OLECB_DROP, obj, effect, pt);
    if (t->curObj) { t->curObj->lpVtbl->Release(t->curObj); t->curObj = NULL; }
    return S_OK;
}

static IDropTargetVtbl kDropTargetVtbl = {
    dt_QueryInterface, dt_AddRef, dt_Release,
    dt_DragEnter, dt_DragOver, dt_DragLeave, dt_Drop
};

int32_t vb6_OLEDrop_Register(void* hwnd) {
    if (!hwnd) return 0;
    if (targetFind((HWND)hwnd)) return 1;   // 已注册
    Vb6DropTarget* t = (Vb6DropTarget*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6DropTarget));
    if (!t) return 0;
    t->vt.lpVtbl = &kDropTargetVtbl;
    t->hwnd = (HWND)hwnd;
    if (RegisterDragDrop((HWND)hwnd, &t->vt) != S_OK) { HeapFree(GetProcessHeap(), 0, t); return 0; }
    if (g_targetCount < VB6_OLETARGET_MAX) {
        g_targets[g_targetCount].hwnd = (HWND)hwnd;
        g_targetCount++;
    }
    return 1;
}

// ---------------- 源侧入口 ----------------
// X.OLEDrag 的落地: 造好 DataObject 后调这里。allowedEffects 用 VB6 常量
// (1=Copy 2=Move, 位组合)。返回最终效果 (DoDragDrop 的 dwOutEffect)。
int32_t vb6_OLEDrag_Start(void* hwnd, void* dataObj, int32_t allowedEffects) {
    if (!hwnd || !dataObj) return 0;
    // DoDragDrop 要求当前线程已 OleInitialize —— 没初始化返回 CO_E_NOTINITIALIZED
    // (0x800401F0), 拖动静默不发生。LoadPicture 那条路会初始化, 纯拖放场景不会。
    // (vb6_OleEnsureInit 是 vb6rtl_com.c 里的 static, 跨不了编译单元 —— 自己来一次。
    //  重复调用无害: S_FALSE / RPC_E_CHANGED_MODE 都算"已初始化"。)
    { static int oleInit44 = 0; if (!oleInit44) { oleInit44 = 1; OleInitialize(NULL); } }
    // 无处理器时也要能拖 (VB6 的 OLEDrag 不要求写 OLEStartDrag)
    Vb6OLETarget* t = targetFind((HWND)hwnd);
    if (!t) {
        if (g_targetCount >= VB6_OLETARGET_MAX) return 0;
        t = &g_targets[g_targetCount++];
        t->hwnd = (HWND)hwnd;
    }
    if (t->srcCb[0]) {   // OLEStartDrag(Data, AllowedEffects)
        void* objPtr = dataObj;
        int32_t allowed = allowedEffects;
        ((void(*)(void**, int32_t*))t->srcCb[0])(&objPtr, &allowed);
        allowedEffects = allowed;
    }
    // ⚠ DoDragDrop 是**模态循环**, 只有真实按住鼠标键时才会自行结束
    // (QueryContinueDrag 见键松开才返回 DRAGDROP_S_DROP)。所以两种情况下**不进**它:
    //   ① 没按住左键 —— 程序化调用 / 无交互会话 (CI runner 就是这种): 实测 CI 上
    //      空转满 60s 超时、CPU 59s (纯自旋)。VB6 里 OLEDrag 由 MouseMove 且按键触发,
    //      真实拖动必然带左键, 所以这条不影响真机。
    //   ② C3_OLEDDB_TEST=1 (无头联测) —— 只想验**源事件链**。
    // 两者都只跑 StartDrag (上面已调) 再以 allowedEffects 当"落下"结果调 CompleteDrag。
    {
        static int headless44 = -1;
        if (headless44 < 0) {
            wchar_t f44[8] = { 0 };
            headless44 = (GetEnvironmentVariableW(L"C3_OLEDDB_TEST", f44, 8) > 0) ? 1 : 0;
        }
        int noButton44 = (GetKeyState(VK_LBUTTON) & 0x8000) ? 0 : 1;
        if (headless44 || noButton44) {
            if (t->srcCb[3]) { int32_t fe44 = allowedEffects; ((void(*)(int32_t*))t->srcCb[3])(&fe44); }
            return allowedEffects;
        }
    }
    g_srcActiveHwnd = (HWND)hwnd;
    g_srcDataObj = (IDataObject*)dataObj;
    // DropSource: 每次 DoDragDrop 造一个 (引用计数归 DoDragDrop/我们各管一段)
    Vb6DropSrc* src = (Vb6DropSrc*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6DropSrc));
    if (!src) { g_srcActiveHwnd = NULL; g_srcDataObj = NULL; return 0; }
    src->vt.lpVtbl = &kDropSrcVtbl;
    src->refs = 1;
    DWORD outEffect = (DWORD)allowedEffects;
    HRESULT hr = DoDragDrop((IDataObject*)dataObj, &src->vt,
                            (DWORD)allowedEffects, &outEffect);
    src->vt.lpVtbl->Release(&src->vt);
    if (t->srcCb[3]) {   // OLECompleteDrag(Effect)
        int32_t finalE = (hr == DRAGDROP_S_DROP) ? (int32_t)outEffect : 0;
        ((void(*)(int32_t*))t->srcCb[3])(&finalE);
    }
    g_srcActiveHwnd = NULL;
    g_srcDataObj = NULL;
    return (hr == DRAGDROP_S_DROP) ? (int32_t)outEffect : 0;
}

// Data.SetData: 直接换文本 (CF_UNICODETEXT)。文件/其它格式后续按需扩。
void vb6_oleDD_SetText(void* dataObj, void* bstrText) {
    // ⚠ 入参是 **IDataObject 接口指针** (&o->vt), 不是 Vb6DataObj —— 数据结构在 o->d。
    // 直接当 Vb6DataObj* 用会把 vt.lpVtbl 当 text 指针 free → 0xC0000005 (实测踩过)。
    Vb6IDataObj* o = (Vb6IDataObj*)dataObj;
    if (!o || !o->d) return;
    Vb6DataObj* d = o->d;
    if (d->text) HeapFree(GetProcessHeap(), 0, d->text);
    d->text = NULL;
    UINT len = SysStringLen((BSTR)bstrText);
    if (len) {
        d->text = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, (len + 1) * sizeof(wchar_t));
        memcpy(d->text, (const wchar_t*)bstrText, len * sizeof(wchar_t));
        d->text[len] = 0;
    }
}

// Data.Clear
void vb6_oleDD_Clear(void* dataObj) {
    Vb6IDataObj* o = (Vb6IDataObj*)dataObj;   // 同上: 接口指针 -> o->d
    if (!o || !o->d) return;
    Vb6DataObj* d = o->d;
    if (d->text) { HeapFree(GetProcessHeap(), 0, d->text); d->text = NULL; }
    if (d->files) {
        for (int i = 0; i < d->fileCount; i++)
            if (d->files[i]) HeapFree(GetProcessHeap(), 0, d->files[i]);
        HeapFree(GetProcessHeap(), 0, d->files);
        d->files = NULL; d->fileCount = 0;
    }
}

// Data.Files(i): 返回 BSTR 副本 (同 GetText 的口径)
void* vb6_oleDD_GetFileBstr(void* dataObj, int32_t idx) {
    const wchar_t* f = vb6_oleDD_GetFile((IDataObject*)dataObj, idx);
    if (!f) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(f);
}

// 窗体销毁时统一撤销所有已注册目标 —— 逐个 Revoke 比记"哪个还活着"简单可靠。
void vb6_OLEDrop_RevokeAll(void) {
    for (int i = 0; i < g_targetCount; i++) {
        if (g_targets[i].hwnd) RevokeDragDrop(g_targets[i].hwnd);
    }
    g_targetCount = 0;
}

void vb6_OLEDrop_Revoke(void* hwnd) {
    if (!hwnd) return;
    RevokeDragDrop((HWND)hwnd);
    for (int i = 0; i < g_targetCount; i++) {
        if (g_targets[i].hwnd == (HWND)hwnd) {
            g_targets[i] = g_targets[g_targetCount - 1];
            g_targetCount--;
            return;
        }
    }
}

// ===================== 无头联测: 给已注册目标发一次 Drop =====================
// C3_OLEDDB_TEST=1 时由主消息循环入口调一次 (那时窗体已显示、目标已注册)。
// 事件链: 造 DataObject("OLE-TEST-DROP") → 直接调该目标的 DragEnter/Drop
// (与真实拖拽走同一条 WM_NOTIFY 之外的 IDropTarget 路径)。
static void testFireCb(void** dataObj, int32_t* effect, int16_t* button,
                       int16_t* shift, float* x, float* y) {
    // 无输出文件也能验: 事件链通了, VB 处理器自己的 Debug.Print 会说话
    if (*effect == 0) *effect = 1;
}

void vb6_oleDD_FireTestDropAtRegistered(void) {
    for (int i = 0; i < g_targetCount; i++) {
        HWND h = g_targets[i].hwnd;
        if (!h) continue;
        IDataObject* obj = (IDataObject*)vb6_oleDD_MakeDataObject(L"OLE-TEST-DROP", NULL, 0);
        if (!obj) continue;
        Vb6DropTarget* t = (Vb6DropTarget*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6DropTarget));
        if (!t) { obj->lpVtbl->Release(obj); continue; }
        t->vt.lpVtbl = &kDropTargetVtbl;
        t->hwnd = h;
        POINTL pt = { 10, 10 };
        DWORD eff = DROPEFFECT_COPY;
        t->vt.lpVtbl->DragEnter(&t->vt, obj, MK_LBUTTON, pt, &eff);
        eff = DROPEFFECT_COPY;
        t->vt.lpVtbl->Drop(&t->vt, obj, MK_LBUTTON, pt, &eff);
        obj->lpVtbl->Release(obj);
        HeapFree(GetProcessHeap(), 0, t);
    }
}

// ===================== 自测 =====================
// 建一个哑窗口当目标, 注册, 然后直接调 IDropTarget 的 DragEnter/Drop
// (绕开 DoDragDrop 的模态循环 —— 无头环境没法真拖), 把 VB 回调收到的内容写进文件。

static FILE* g_testOut = NULL;
static void testCb(void** dataObj, int32_t* effect, int16_t* button,
                   int16_t* shift, float* x, float* y) {
    if (!g_testOut) return;
    const wchar_t* txt = vb6_oleDD_GetText((IDataObject*)*dataObj);
    fprintf(g_testOut, "cb textlen=%d effect=%d btn=%d shift=%d x=%.1f y=%.1f\n",
            (int)lstrlenW(txt), (int)*effect, (int)*button, (int)*shift, (double)*x, (double)*y);
    // 顺手验证 effect 回写: VB 侧把 effect 改成 Move(2) 应能传回 OLE
    if (*effect == 1) *effect = 2;
}

int32_t vb6_oleDD_SelfTest(const wchar_t* outPath) {
    WNDCLASSW wc;
    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = DefWindowProcW;
    wc.lpszClassName = L"C3_OLEDDB_TEST";
    wc.hInstance = GetModuleHandleW(NULL);
    RegisterClassW(&wc);
    HWND w = CreateWindowExW(0, L"C3_OLEDDB_TEST", L"t", WS_OVERLAPPEDWINDOW,
                             0, 0, 200, 100, NULL, NULL, wc.hInstance, NULL);
    if (!w) return 0;

    g_testOut = _wfopen(outPath, L"w");
    if (!g_testOut) return 0;
    fprintf(g_testOut, "selftest begin\n");

    // 造 DataObject (文本 + 两个文件)
    wchar_t* files[2] = { (wchar_t*)L"C:\\a.txt", (wchar_t*)L"C:\\b.txt" };
    IDataObject* obj = (IDataObject*)vb6_oleDD_MakeDataObject(L"拖放文本", files, 2);
    if (!obj) { fprintf(g_testOut, "FAIL make dataobject\n"); fclose(g_testOut); return 0; }

    // 拿到我们注册的 IDropTarget (RegisterDragDrop 内部存的, 这里直接构造一份调用)
    Vb6DropTarget* t = (Vb6DropTarget*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6DropTarget));
    t->vt.lpVtbl = &kDropTargetVtbl;
    t->hwnd = w;
    // 回调表与 RegisterDragDrop 共用同一张 g_targets —— vb6_OLEDrop_SetHandler 先登记
    vb6_OLEDrop_SetHandler(w, VB6_OLECB_OVER, (void*)testCb);
    vb6_OLEDrop_SetHandler(w, VB6_OLECB_DROP, (void*)testCb);

    POINTL pt = { 30, 40 };
    DWORD eff = DROPEFFECT_COPY;
    // 走完整的三段: DragEnter → DragOver → Drop
    t->vt.lpVtbl->DragEnter(&t->vt, obj, MK_LBUTTON, pt, &eff);
    fprintf(g_testOut, "after enter effect=%d\n", (int)eff);
    eff = DROPEFFECT_COPY;
    t->vt.lpVtbl->DragOver(&t->vt, MK_LBUTTON, pt, &eff);
    fprintf(g_testOut, "after over effect=%d\n", (int)eff);
    eff = DROPEFFECT_COPY;
    HRESULT hr = t->vt.lpVtbl->Drop(&t->vt, obj, MK_LBUTTON, pt, &eff);
    fprintf(g_testOut, "after drop hr=0x%lx effect=%d\n", (unsigned long)hr, (int)eff);

    // 验证 QueryGetData / EnumFormatEtc
    FORMATETC f; ZeroMemory(&f, sizeof(f)); f.cfFormat = CF_UNICODETEXT; f.tymed = TYMED_HGLOBAL;
    fprintf(g_testOut, "queryText=0x%lx\n", (unsigned long)obj->lpVtbl->QueryGetData(obj, &f));
    IEnumFORMATETC* en = NULL;
    obj->lpVtbl->EnumFormatEtc(obj, DATADIR_GET, &en);
    fprintf(g_testOut, "enum=%s\n", en ? "ok" : "NULL");
    if (en) { en->lpVtbl->Release(en); }

    // 验证 CF_HDROP 取出
    STGMEDIUM med;
    ZeroMemory(&med, sizeof(med));
    f.cfFormat = CF_HDROP;
    HRESULT hd = obj->lpVtbl->GetData(obj, &f, &med);
    if (hd == S_OK && med.hGlobal) {
        DROPFILES* df = (DROPFILES*)GlobalLock(med.hGlobal);
        wchar_t* list = (wchar_t*)((char*)df + df->pFiles);
        fprintf(g_testOut, "hdrop first=%ls\n", list);
        GlobalUnlock(med.hGlobal);
        ReleaseStgMedium(&med);
    } else {
        fprintf(g_testOut, "hdrop FAIL hr=0x%lx\n", (unsigned long)hd);
    }

    obj->lpVtbl->Release(obj);
    HeapFree(GetProcessHeap(), 0, t);
    fprintf(g_testOut, "selftest end\n");
    fclose(g_testOut);
    g_testOut = NULL;
    DestroyWindow(w);
    return 1;
}

#endif  // _WIN32
