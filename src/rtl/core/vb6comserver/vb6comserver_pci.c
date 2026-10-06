// vb6comserver_pci.c - vb6comserver 模块拆分: 类型信息 (IProvideClassInfo2)
// 由 vb6comserver.c 按 COM 接口家族拆分而来 (纯搬移, 零行为改动)

#include <string.h>
#include <wchar.h>   /* 账 #173 census: swprintf 的原型 (此前走 C4013) */
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
// IProvideClassInfo2 helper
// ============================================================

// Load ITypeInfo for this coclass from the registered TypeLib
static HRESULT vb6_LoadCoClassTypeInfo(const vb6_CoClassDesc* desc, ITypeInfo** ppTypeInfo) {
    HKEY hKey;
    LONG ret;
    wchar_t dllPath[MAX_PATH];
    DWORD sz;
    ITypeLib* pTypeLib;
    CLSID clsid;
    HRESULT hr;
    UINT count, i;
    
    if (!desc || !ppTypeInfo) return E_POINTER;
    *ppTypeInfo = NULL;
    
    // Parse CLSID
    hr = vb6_CLSIDFromStrA(desc->clsidStr, &clsid);
    if (FAILED(hr)) return hr;
    
    // Get DLL path from CLSID\InprocServer32
    {
        wchar_t clsidStr[64];
        wchar_t keyPath[256];
        StringFromGUID2(&clsid, clsidStr, 64);
        swprintf(keyPath, 256, L"CLSID\\%s\\InprocServer32", clsidStr);
        hKey = NULL;
        ret = RegOpenKeyExW(HKEY_CLASSES_ROOT, keyPath, 0, KEY_READ, &hKey);
        if (ret != ERROR_SUCCESS) return TYPE_E_REGISTRYACCESS;
        sz = sizeof(dllPath);
        ret = RegQueryValueExW(hKey, NULL, NULL, NULL, (LPBYTE)dllPath, &sz);
        RegCloseKey(hKey);
        if (ret != ERROR_SUCCESS) return TYPE_E_REGISTRYACCESS;
    }
    
    // Load TypeLib directly from the DLL file
    hr = LoadTypeLib(dllPath, &pTypeLib);
    if (FAILED(hr)) return hr;
    
    // Find the coclass ITypeInfo by CLSID
    count = pTypeLib->lpVtbl->GetTypeInfoCount(pTypeLib);
    for (i = 0; i < count; i++) {
        ITypeInfo* pInfo = NULL;
        hr = pTypeLib->lpVtbl->GetTypeInfo(pTypeLib, i, &pInfo);
        if (FAILED(hr)) continue;
        
        TYPEATTR* pAttr = NULL;
        hr = pInfo->lpVtbl->GetTypeAttr(pInfo, &pAttr);
        if (SUCCEEDED(hr) && pAttr) {
            if (pAttr->typekind == TKIND_COCLASS && IsEqualIID(&pAttr->guid, &clsid)) {
                *ppTypeInfo = pInfo;
                pInfo->lpVtbl->ReleaseTypeAttr(pInfo, pAttr);
                ITypeLib_Release(pTypeLib);
                return S_OK;
            }
            pInfo->lpVtbl->ReleaseTypeAttr(pInfo, pAttr);
        }
        pInfo->lpVtbl->Release(pInfo);
    }
    
    ITypeLib_Release(pTypeLib);
    return TYPE_E_ELEMENTNOTFOUND;
}


// ============================================================
// IProvideClassInfo2 implementation (minimal - GetGUID only)
// ============================================================

static HRESULT STDMETHODCALLTYPE PCI_QueryInterface(vb6_ProvideClassInfo2* self, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown_) || IsEqualIID(riid, &IID_IProvideClassInfo2_)) {
        *ppv = self;
        self->vtable->AddRef(self);
        return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE PCI_AddRef(vb6_ProvideClassInfo2* self) {
    return InterlockedIncrement(&self->refCount);
}

static ULONG STDMETHODCALLTYPE PCI_Release(vb6_ProvideClassInfo2* self) {
    ULONG c = InterlockedDecrement(&self->refCount);
    if (c == 0) {
        CoTaskMemFree(self);
    }
    return c;
}

static HRESULT STDMETHODCALLTYPE PCI_GetClassInfo(vb6_ProvideClassInfo2* self, ITypeInfo** ppTI) {
    if (!ppTI) return E_POINTER;
    *ppTI = NULL;
    if (!self->comObj || !self->comObj->desc) return E_FAIL;
    return vb6_LoadCoClassTypeInfo(self->comObj->desc, ppTI);
}

static HRESULT STDMETHODCALLTYPE PCI_GetGUID(vb6_ProvideClassInfo2* self, DWORD dwGuidKind, GUID* pGUID) {
    if (!pGUID) return E_POINTER;
    if (dwGuidKind != GUIDKIND_DEFAULT_SOURCE_DISP_IID) return E_FAIL;
    if (!self->comObj || !self->comObj->desc || !self->comObj->desc->sourceIfaceIid) return E_FAIL;
    // Parse source IID from string
    wchar_t wbuf[64];
    MultiByteToWideChar(CP_ACP, 0, self->comObj->desc->sourceIfaceIid, -1, wbuf, 64);
    return CLSIDFromString(wbuf, pGUID);
}

static const vb6_IProvideClassInfo2Vtable g_PCIVtable = {
    PCI_QueryInterface,
    PCI_AddRef,
    PCI_Release,
    PCI_GetClassInfo,
    PCI_GetGUID,
};

vb6_ProvideClassInfo2* vb6_PCI_Create(vb6_ComObject* comObj) {
    if (!comObj) return NULL;
    vb6_ProvideClassInfo2* pci = (vb6_ProvideClassInfo2*)CoTaskMemAlloc(sizeof(vb6_ProvideClassInfo2));
    if (!pci) return NULL;
    pci->vtable = &g_PCIVtable;
    pci->refCount = 1;
    pci->comObj = comObj;
    return pci;
}
