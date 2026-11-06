// vb6comserver_cp.c - vb6comserver 模块拆分: 连接点 (IConnectionPoint / IConnectionPointContainer)
// 由 vb6comserver.c 按 COM 接口家族拆分而来 (纯搬移, 零行为改动)

#include <string.h>
#include <stdio.h>   /* ucTreeView: vb6_UCEventMiss 的 stderr 行 */
#define COBJMACROS  /* Enable C COM macros (ITypeLib_Release etc.) */
#include "vb6comserver.h"
#include "vb6comserver_internal.h"

#ifndef CONNECT_E_NOCONNECTION
#define CONNECT_E_NOCONNECTION 0x80040200
#endif

#ifndef GUIDKIND_DEFAULT_SOURCE_DISP_IID
#define GUIDKIND_DEFAULT_SOURCE_DISP_IID 1
#endif


// ============================================================
// IConnectionPointContainer + IConnectionPoint implementation
// ============================================================

// --- IConnectionPoint methods ---

static HRESULT STDMETHODCALLTYPE CP_QueryInterface(vb6_ConnectionPoint* self, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown_) || IsEqualIID(riid, &IID_IConnectionPoint_)) {
        *ppv = self;
        self->vtable->AddRef(self);
        return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE CP_AddRef(vb6_ConnectionPoint* self) {
    return InterlockedIncrement(&self->refCount);
}

static ULONG STDMETHODCALLTYPE CP_Release(vb6_ConnectionPoint* self) {
    ULONG c = InterlockedDecrement(&self->refCount);
    if (c == 0) {
        int i;
        for (i = 0; i < self->connCount; i++) {
            if (self->sinks[i]) self->sinks[i]->lpVtbl->Release(self->sinks[i]);
        }
        CoTaskMemFree(self->cookies);
        CoTaskMemFree(self->sinks);
        CoTaskMemFree(self);
    }
    return c;
}

static HRESULT STDMETHODCALLTYPE CP_GetConnectionInterface(vb6_ConnectionPoint* self, IID* pIID) {
    if (!pIID) return E_POINTER;
    *pIID = self->sourceIid;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE CP_GetConnectionPointContainer(vb6_ConnectionPoint* self, void** ppCPC) {
    if (!ppCPC) return E_POINTER;
    if (!self->container) return E_FAIL;
    *ppCPC = self->container;
    self->container->vtable->AddRef(self->container);
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE CP_Advise(vb6_ConnectionPoint* self, IUnknown* pUnkSink, DWORD* pdwCookie) {
    int i;
    if (!pUnkSink || !pdwCookie) return E_POINTER;
    *pdwCookie = 0;
    // Grow arrays if needed
    if (self->connCount >= self->connCapacity) {
        int newCap = self->connCapacity ? self->connCapacity * 2 : 4;
        DWORD* newCookies = (DWORD*)CoTaskMemRealloc(self->cookies, newCap * sizeof(DWORD));
        IUnknown** newSinks = (IUnknown**)CoTaskMemRealloc(self->sinks, newCap * sizeof(IUnknown*));
        if (!newCookies || !newSinks) return E_OUTOFMEMORY;
        self->cookies = newCookies;
        self->sinks = newSinks;
        self->connCapacity = newCap;
    }
    i = self->connCount++;
    self->cookies[i] = ++self->nextCookie;
    pUnkSink->lpVtbl->AddRef(pUnkSink);
    self->sinks[i] = pUnkSink;
    *pdwCookie = self->cookies[i];
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE CP_Unadvise(vb6_ConnectionPoint* self, DWORD dwCookie) {
    int i;
    if (dwCookie == 0) return E_POINTER;
    for (i = 0; i < self->connCount; i++) {
        if (self->cookies[i] == dwCookie) {
            if (self->sinks[i]) {
                self->sinks[i]->lpVtbl->Release(self->sinks[i]);
                self->sinks[i] = NULL;
            }
            self->cookies[i] = 0;
            return S_OK;
        }
    }
    return CONNECT_E_NOCONNECTION;
}

static HRESULT STDMETHODCALLTYPE CP_EnumConnections(vb6_ConnectionPoint* self, void** ppEnum) {
    if (!ppEnum) return E_POINTER;
    *ppEnum = NULL;
    return E_NOTIMPL;
}

static const vb6_IConnectionPointVtable g_CPVtable = {
    CP_QueryInterface,
    CP_AddRef,
    CP_Release,
    CP_GetConnectionInterface,
    CP_GetConnectionPointContainer,
    CP_Advise,
    CP_Unadvise,
    CP_EnumConnections,
};


// --- IConnectionPointContainer methods ---

static HRESULT STDMETHODCALLTYPE CPC_QueryInterface(vb6_ConnectionPointContainer* self, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown_) || IsEqualIID(riid, &IID_IConnectionPointContainer_)) {
        *ppv = self;
        self->vtable->AddRef(self);
        return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE CPC_AddRef(vb6_ConnectionPointContainer* self) {
    return InterlockedIncrement(&self->refCount);
}

static ULONG STDMETHODCALLTYPE CPC_Release(vb6_ConnectionPointContainer* self) {
    ULONG c = InterlockedDecrement(&self->refCount);
    if (c == 0) {
        if (self->connPoint) self->connPoint->vtable->Release(self->connPoint);
        CoTaskMemFree(self);
    }
    return c;
}

static HRESULT STDMETHODCALLTYPE CPC_EnumConnectionPoints(vb6_ConnectionPointContainer* self, void** ppEnum) {
    if (!ppEnum) return E_POINTER;
    *ppEnum = NULL;
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE CPC_FindConnectionPoint(vb6_ConnectionPointContainer* self, REFIID riid, void** ppCP) {
    if (!ppCP) return E_POINTER;
    *ppCP = NULL;
    if (!self->connPoint) return CONNECT_E_NOCONNECTION;
    // Check if requested IID matches our source interface
    if (IsEqualIID(riid, &self->connPoint->sourceIid)) {
        *ppCP = self->connPoint;
        self->connPoint->vtable->AddRef(self->connPoint);
        return S_OK;
    }
    return CONNECT_E_NOCONNECTION;
}

static const vb6_IConnectionPointContainerVtable g_CPCVtable = {
    CPC_QueryInterface,
    CPC_AddRef,
    CPC_Release,
    CPC_EnumConnectionPoints,
    CPC_FindConnectionPoint,
};


// Create ConnectionPointContainer for a COM object
vb6_ConnectionPointContainer* vb6_CPC_Create(vb6_ComObject* comObj) {
    if (!comObj || !comObj->desc || !comObj->desc->sourceIfaceIid) return NULL;
    
    vb6_ConnectionPointContainer* cpc = (vb6_ConnectionPointContainer*)CoTaskMemAlloc(sizeof(vb6_ConnectionPointContainer));
    if (!cpc) return NULL;
    
    cpc->vtable = &g_CPCVtable;
    cpc->refCount = 1;
    cpc->comObj = comObj;
    cpc->connPoint = NULL;
    
    // Create the single ConnectionPoint
    vb6_ConnectionPoint* cp = (vb6_ConnectionPoint*)CoTaskMemAlloc(sizeof(vb6_ConnectionPoint));
    if (!cp) {
        CoTaskMemFree(cpc);
        return NULL;
    }
    cp->vtable = &g_CPVtable;
    cp->refCount = 1;
    // Parse source IID from string
    {
        wchar_t wbuf[64];
        MultiByteToWideChar(CP_ACP, 0, comObj->desc->sourceIfaceIid, -1, wbuf, 64);
        CLSIDFromString(wbuf, &cp->sourceIid);
    }
    cp->container = cpc;
    cp->cookies = NULL;
    cp->sinks = NULL;
    cp->connCount = 0;
    cp->connCapacity = 0;
    cp->nextCookie = 0;
    
    cpc->connPoint = cp;
    return cpc;
}

// Fire an event to all connected sinks
void vb6_FireEvent(vb6_ComObject* comObj, int32_t dispid, VARIANT* args, int argc) {
    int i, j;
    if (!comObj || !comObj->cpc || !comObj->cpc->connPoint) return;
    vb6_ConnectionPoint* cp = comObj->cpc->connPoint;
    
    for (i = 0; i < cp->connCount; i++) {
        if (!cp->sinks[i]) continue;
        // QI for IDispatch
        IDispatch* pDisp = NULL;
        HRESULT hr = cp->sinks[i]->lpVtbl->QueryInterface(cp->sinks[i], &IID_IDispatch_, (void**)&pDisp);
        if (SUCCEEDED(hr) && pDisp) {
            DISPPARAMS dp = {0};
            VARIANT* reversedArgs = NULL;
            if (argc > 0) {
                reversedArgs = (VARIANT*)CoTaskMemAlloc(argc * sizeof(VARIANT));
                for (j = 0; j < argc; j++) {
                    VariantInit(&reversedArgs[j]);
                    VariantCopy(&reversedArgs[j], &args[argc - 1 - j]);
                }
                dp.rgvarg = reversedArgs;
                dp.cArgs = (UINT)argc;
            }
            pDisp->lpVtbl->Invoke(pDisp, dispid, &IID_NULL, LOCALE_USER_DEFAULT,
                DISPATCH_METHOD, &dp, NULL, NULL, NULL);
            if (reversedArgs) {
                for (j = 0; j < argc; j++) VariantClear(&reversedArgs[j]);
                CoTaskMemFree(reversedArgs);
            }
            pDisp->lpVtbl->Release(pDisp);
        }
    }
}

/* ucTreeView (2026-11-06): RaiseEvent 的可见 miss —— 见 vb6comserver.h 说明。
 * 唯一出口: 计数 + 可选 trace, 不改分发语义。 */
static int vb6_ucEventMiss = 0;

void vb6_UCEventMiss(const wchar_t* evtName) {
    static int vb6_traceOn = -1;
    vb6_ucEventMiss++;
    if (vb6_traceOn < 0) {
        vb6_traceOn = (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) ? 1 : 0;
    }
    if (vb6_traceOn) {
        fprintf(stderr, "[C3_EVT_MISS] #%d event='%ls' (events sink or handler absent)\n",
                vb6_ucEventMiss, evtName ? evtName : L"?");
    }
}

int vb6_UCEventMissCount(void) { return vb6_ucEventMiss; }
