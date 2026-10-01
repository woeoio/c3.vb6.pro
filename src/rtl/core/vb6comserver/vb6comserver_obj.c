// vb6comserver_obj.c - vb6comserver 模块拆分: IDispatch 包装层 (vb6_ComObject)
// 由 vb6comserver.c 按 COM 接口家族拆分而来 (纯搬移, 零行为改动)

#include <string.h>
#include <stdlib.h>
#define COBJMACROS  /* Enable C COM macros (ITypeLib_Release etc.) */
#include "vb6comserver.h"
#include "vb6comserver_internal.h"

/* [PROBE189] 临时探针: 包装器引用计数轨迹 (C3_COM_RC_TRACE=1 时输出到 stderr) */
#include <stdio.h>
static int probe189_on(void) {
    static int v = -1;
    if (v < 0) v = (getenv("C3_COM_RC_TRACE") != NULL);
    return v;
}
#define P189(...) do { if (probe189_on()) { fprintf(stderr, __VA_ARGS__); fflush(stderr); } } while (0)
/* [/PROBE189] */

#ifndef CONNECT_E_NOCONNECTION
#define CONNECT_E_NOCONNECTION 0x80040200
#endif

#ifndef GUIDKIND_DEFAULT_SOURCE_DISP_IID
#define GUIDKIND_DEFAULT_SOURCE_DISP_IID 1
#endif

// ============================================================
// vb6_ComObject - IDispatch wrapper implementation
// ============================================================

// --- IUnknown ---

static HRESULT STDMETHODCALLTYPE ComObj_QueryInterface(vb6_ComObject* self, REFIID riid, void** ppv) {
    int i;
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown_) || IsEqualIID(riid, &IID_IDispatch_)) {
        *ppv = self;
        self->vtable->AddRef(self);
        return S_OK;
    }
    // P12.1: Check Implements interface IIDs
    if (self->desc && self->desc->ifaceCount > 0 && self->desc->ifaceIids) {
        for (i = 0; i < self->desc->ifaceCount; i++) {
            if (IsEqualIID(riid, self->desc->ifaceIids[i])) {
                // ai/022 B16: 新式接口 (Interface 块) 交回**薄指针** —— 首字段是
                // vb6_ivtbl_<I>*, 槽 3 起是契约成员, 与类型库 TKIND_INTERFACE 那一条
                // (cFuncs / oVft / CC_STDCALL) 同形. 这正是 B15 起对外广告的那一档,
                // 所以"广告 == 应答"在接口这一档上成立 (D60-4 的"胖应答瘦"到此收口).
                // 生命周期: 薄指针的引用记在实例自己的 __refcount 上 (B05), 包装器
                // 最后一次 Release 会把底座引用交还 (见 ComObj_Release) → QI 之后立刻
                // Release 包装器 (IClassFactory::CreateInstance 的规范姿势) 也安全.
                if (self->desc->ifaceThinPtr && self->vb6Instance) {
                    void* thin = self->desc->ifaceThinPtr(self->vb6Instance,
                                                           self->desc->ifaceIids[i]);
                    if (thin) {
                        vb6_ivtbl_prefix* vt = *(vb6_ivtbl_prefix**)thin;
                        vt->AddRef(thin);
                        *ppv = thin;
                        return S_OK;
                    }
                    /* 本类没实现这个新式接口 (provider 认 IID): 落到胖路 */
                }
                *ppv = self;  // dispinterface: same IDispatch pointer
                self->vtable->AddRef(self);
                return S_OK;
            }
        }
    }
    // Default dispinterface IID
    if (self->desc && self->desc->defaultIfaceIid && IsEqualIID(riid, self->desc->defaultIfaceIid)) {
        *ppv = self;  // dispinterface = same IDispatch pointer
        self->vtable->AddRef(self);
        return S_OK;
    }
    // IConnectionPointContainer (only if coclass has events)
    if (self->desc && self->desc->sourceIfaceIid && IsEqualIID(riid, &IID_IConnectionPointContainer_)) {
        if (!self->cpc) {
            self->cpc = vb6_CPC_Create(self);
        }
        if (self->cpc) {
            *ppv = self->cpc;
            self->cpc->vtable->AddRef(self->cpc);
            return S_OK;
        }
    }
    // IProvideClassInfo2
    if (self->desc && IsEqualIID(riid, &IID_IProvideClassInfo2_)) {
        if (!self->pci) {
            self->pci = vb6_PCI_Create(self);
        }
        if (self->pci) {
            *ppv = self->pci;
            self->pci->vtable->AddRef(self->pci);
            return S_OK;
        }
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE ComObj_AddRef(vb6_ComObject* self) {
    ULONG count = InterlockedIncrement(&self->refCount);
    P189("[189] ADDREF W=%p cnt=%u\n", (void*)self, (unsigned)count);
    InterlockedIncrement(&g_vb6_cRef);
    return count;
}

static ULONG STDMETHODCALLTYPE ComObj_Release(vb6_ComObject* self) {
    ULONG count = InterlockedDecrement(&self->refCount);
    P189("[189] RELEASE W=%p cnt=%u owns=%d inst=%p\n", (void*)self, (unsigned)count,
         self->ownsInstance, (void*)self->vb6Instance);
    InterlockedDecrement(&g_vb6_cRef);
    if (count == 0) {
        // Fix 188 + ai/022 B16 的并集: 是否销毁实例先看**拥有关系** —— 只有拥有型包装
        // (CoCreateInstance / Set x = New cY) 才动实例; 借用型 (Public 对象字段 getter /
        // 方法返回的工程类实例) 的生命周期归宿主, 只清实例上的 __comObj 回填。
        // 拥有型里, 实现了新式接口的类走 claim 那条: 包装器退掉自己的"底座引用",
        // 实例若有薄引用在世则不销毁 (由最后一个薄引用的 Release 收尾), 否则当场销毁;
        // 没有这道 claim 的类 (存量/无接口) 结果与原来直接 destroyFunc 完全一致。
        if (self->ownsInstance) {
            if (self->desc && self->desc->instanceClaimRelease && self->vb6Instance) {
                self->desc->instanceClaimRelease(self->vb6Instance);
            } else if (self->desc && self->desc->destroyFunc && self->vb6Instance) {
                self->desc->destroyFunc(self->vb6Instance);
            }
        } else if (self->vb6Instance) {
            *(void**)self->vb6Instance = NULL;
        }
        self->vb6Instance = NULL;
        // Release PCI
        if (self->pci) {
            self->pci->vtable->Release(self->pci);
            self->pci = NULL;
        }
        // Release CPC
        if (self->cpc) {
            self->cpc->vtable->Release(self->cpc);
            self->cpc = NULL;
        }
        CoTaskMemFree(self);
    }
    return count;
}

// --- IDispatch ---

static HRESULT STDMETHODCALLTYPE ComObj_GetTypeInfoCount(vb6_ComObject* self, UINT* pctinfo) {
    if (!pctinfo) return E_POINTER;
    *pctinfo = 1;  // Provide TypeLib info (needed for early binding)
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE ComObj_GetTypeInfo(vb6_ComObject* self, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) {
    if (!ppTInfo) return E_POINTER;
    if (iTInfo != 0) return DISP_E_BADINDEX;
    *ppTInfo = NULL;
    if (!self->desc) return E_FAIL;
    // Get DLL path, load TypeLib from embedded resource
    wchar_t dllPath[MAX_PATH];
    HMODULE hMod = NULL;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCWSTR)&g_vb6_cRef, &hMod)) {
        return E_FAIL;
    }
    GetModuleFileNameW(hMod, dllPath, MAX_PATH);
    ITypeLib* pTypeLib = NULL;
    HRESULT hr = LoadTypeLib(dllPath, &pTypeLib);
    if (FAILED(hr) || !pTypeLib) {
        return E_FAIL;
    }
    // Find coclass ITypeInfo by CLSID
    CLSID clsid;
    hr = vb6_CLSIDFromStrA(self->desc->clsidStr, &clsid);
    if (FAILED(hr)) {
        ITypeLib_Release(pTypeLib);
        return E_FAIL;
    }
    hr = ITypeLib_GetTypeInfoOfGuid(pTypeLib, &clsid, ppTInfo);
    ITypeLib_Release(pTypeLib);
    return hr;
}

static HRESULT STDMETHODCALLTYPE ComObj_GetIDsOfNames(vb6_ComObject* self, REFIID riid, 
    LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) 
{
    if (!self->desc || !self->desc->methods) return DISP_E_UNKNOWNNAME;
    
    for (UINT n = 0; n < cNames; n++) {
        rgDispId[n] = DISPID_UNKNOWN;
        for (int i = 0; i < self->desc->methodCount; i++) {
            if (wcscmp(rgszNames[n], self->desc->methods[i].name) == 0) {
                rgDispId[n] = self->desc->methods[i].dispid;
                break;
            }
        }
    }
    
    return (cNames > 0 && rgDispId[0] != DISPID_UNKNOWN) ? S_OK : DISP_E_UNKNOWNNAME;
}

static HRESULT STDMETHODCALLTYPE ComObj_Invoke(vb6_ComObject* self, DISPID dispIdMember, 
    REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, 
    VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    if (!self->desc || !self->desc->methods) return DISP_E_MEMBERNOTFOUND;
    
    // Find method matching dispid AND invkind (same dispid may have Get/Let variants)
    const vb6_DispMethodDesc* method = NULL;
    for (int i = 0; i < self->desc->methodCount; i++) {
        if (self->desc->methods[i].dispid == dispIdMember) {
            int ik = self->desc->methods[i].invkind;
            int ok = 0;
            if (ik == 1 && (wFlags & DISPATCH_METHOD)) ok = 1;
            else if (ik == 2 && (wFlags & DISPATCH_PROPERTYGET)) ok = 1;
            else if ((ik == 4 || ik == 8) && (wFlags & DISPATCH_PROPERTYPUT)) ok = 1;
            if (ok) { method = &self->desc->methods[i]; break; }
        }
    }
    // Fallback: if no invkind match, try first dispid match
    if (!method) {
        for (int i = 0; i < self->desc->methodCount; i++) {
            if (self->desc->methods[i].dispid == dispIdMember) {
                method = &self->desc->methods[i];
                break;
            }
        }
    }
    if (!method) return DISP_E_MEMBERNOTFOUND;
    P189("[189] INVOKE W=%p cls=%s mem=%ls argc=%d\n", (void*)self,
         (self->desc && self->desc->classVariable) ? self->desc->classVariable : "?",
         method->name ? method->name : L"?", (int)(pDispParams ? pDispParams->cArgs : 0));

    // Collect arguments
    // Note: Script engines like VBScript may pass VT_I2 etc.,
    // while bridge functions expect VT_I4 (via lVal). Coerce each param to VT_I4.
    // invokeFunc 的实参/返回槽在适配器里按 vb6_VARIANT (x86 24B: vt 4B +
    // pad 4B + union 16B) 布局访问; 前 16 字节与 OLE VARIANT 内容一致。
    typedef struct { uint64_t v[3]; } vb6_VarSlot;
    int argc = pDispParams ? (int)pDispParams->cArgs : 0;
    void** args = NULL;
    void* coercedArgs = NULL;   /* 槽阵列, 每槽 vb6_VarSlot */

    if (argc > 0) {
        args = (void**)CoTaskMemAlloc(argc * sizeof(void*));
        coercedArgs = (void*)CoTaskMemAlloc(argc * sizeof(vb6_VarSlot));
        if (!args || !coercedArgs) {
            if (args) CoTaskMemFree(args);
            if (coercedArgs) CoTaskMemFree(coercedArgs);
            return E_OUTOFMEMORY;
        }
        // 适配器按 vb6_VARIANT (24B) 布局访问实参与返回槽, 而 COM 边界的
        // OLE VARIANT 是 16B。每个实参先整复制进 24B 槽, 再按需数值强转。
        // (ai/022 B17 那条"未初始化就被 VariantClear"的堆损坏在这里由下面的
        //  `memset(slot, 0, sizeof(*slot))` 一并解决 —— 旧写法按 VARIANT 尺寸清零,
        //  在这套 24B 槽设计下 x86 会算少, 故合并时只保留这一份。)
        for (int i = 0; i < argc; i++) {
            VARIANT* src = &pDispParams->rgvarg[argc - 1 - i];
            vb6_VarSlot* slot = &((vb6_VarSlot*)coercedArgs)[i];
            memset(slot, 0, sizeof(*slot));
            memcpy(slot, src, sizeof(VARIANT));
            args[i] = slot;
        }
        for (int i = 0; i < argc; i++) {
            VARIANT* src = &pDispParams->rgvarg[argc - 1 - i];
            VARIANT* dst = (VARIANT*)&((vb6_VarSlot*)coercedArgs)[i];
            // Coerce numeric types to VT_I4 (bridge functions use lVal)
            if (src->vt == VT_I2 || src->vt == VT_I1 || src->vt == VT_UI1 ||
                src->vt == VT_UI2 || src->vt == VT_BOOL || src->vt == VT_EMPTY) {
                VariantChangeType(dst, src, 0, VT_I4);
            } else if (src->vt == VT_R4) {
                // Float -> Double for dblVal access
                VariantChangeType(dst, src, 0, VT_R8);
            }
        }
    }

    // Call VB6 method
    if (method->invokeFunc) {
        vb6_VarSlot tmpRet;
        memset(&tmpRet, 0, sizeof(tmpRet));
        method->invokeFunc(self->vb6Instance, args, argc, &tmpRet);
        if (pVarResult)
            memcpy(pVarResult, &tmpRet, sizeof(VARIANT));
    }

    // 槽内是调用方 VARIANT 的副本, 资源仍归调用方, 不得 VariantClear
    if (coercedArgs) CoTaskMemFree(coercedArgs);
    if (args) CoTaskMemFree(args);
    return S_OK;
}

// IDispatch vtable instance
static const vb6_IDispatchVtable g_ComObjectVtable = {
    ComObj_QueryInterface,
    ComObj_AddRef,
    ComObj_Release,
    ComObj_GetTypeInfoCount,
    ComObj_GetTypeInfo,
    ComObj_GetIDsOfNames,
    ComObj_Invoke,
};

// Create VB6 COM object
vb6_ComObject* vb6_ComObject_Create(const vb6_CoClassDesc* desc) {
    if (!desc || !desc->factoryFunc) return NULL;
    
    vb6_ComObject* obj = (vb6_ComObject*)CoTaskMemAlloc(sizeof(vb6_ComObject));
    if (!obj) return NULL;
    
    obj->vtable = &g_ComObjectVtable;
    obj->refCount = 1;
    obj->desc = desc;
    obj->vb6Instance = desc->factoryFunc();  // Call vb6_cls_<Name>_New()
    P189("[189] CREATE W=%p inst=%p cls=%s\n", (void*)obj, (void*)obj->vb6Instance,
         desc->classVariable ? desc->classVariable : "?");
    obj->ownsInstance = 1;  // Fix 188: 类工厂创建的实例归客户端所有
    obj->cpc = NULL;  // Lazy init CPC
    obj->pci = NULL;  // Lazy init PCI
    // Set back-pointer for event support (first field of VB6 class struct = __comObj)
    if (obj->vb6Instance) {
        void** ppComObj = (void**)obj->vb6Instance;
        *ppComObj = obj;
    }
    
    if (!obj->vb6Instance) {
        CoTaskMemFree(obj);
        return NULL;
    }
    
    InterlockedIncrement(&g_vb6_cRef);
    return obj;
}

// ============================================================
// Fix 099: 包装已存在的 VB6 类实例 (Public 对象字段的 COM 暴露)
// ============================================================
// ActiveX DLL 里 `Public Router As cHttpServerRouter` 这类字段的实例, 是在类
// 内部 (Class_Initialize 的 `Set Router = New ...`) 创建的裸结构体指针, 其
// __comObj 仍为 NULL. COM 客户端读该字段 (Property Get) 时必须拿到一个能继续
// 调用方法的 IDispatch —— 直接返回裸指针冒充 VT_DISPATCH, 客户端首次 Invoke
// 就会把结构体当 IDispatch 解引用 vtable → 0xC0000005.
//
// 复用规则: 实例首个字段 __comObj 已指向包装对象时直接 AddRef 复用 (同一实例
// 每次读出返回同一 IDispatch, 引用计数正确), 否则新建包装并回填 __comObj.
// 前置条件: 实例所在类的结构体首字段必须是 __comObj —— cgen 仅在 ActiveX DLL
// 工程 (isDll_) 的类结构体里生成该字段, 故本函数只能用于 DLL 侧的字段 getter.
vb6_ComObject* vb6_ComObject_FromInstance(const vb6_CoClassDesc* desc, void* instance) {
    void** ppComObj;
    vb6_ComObject* existing;
    vb6_ComObject* obj;
    if (!desc || !instance) return NULL;

    ppComObj = (void**)instance;
    existing = (vb6_ComObject*)*ppComObj;
    if (existing) {
        existing->vtable->AddRef(existing);
        return existing;
    }

    obj = (vb6_ComObject*)CoTaskMemAlloc(sizeof(vb6_ComObject));
    if (!obj) return NULL;

    obj->vtable = &g_ComObjectVtable;
    obj->refCount = 1;
    obj->desc = desc;
    obj->vb6Instance = instance;
    obj->ownsInstance = 1;  /* Fix 188: 调用方把新实例的所有权交给包装器 */
    P189("[189] FROM-INST W=%p inst=%p cls=%s\n", (void*)obj, (void*)instance,
         desc->classVariable ? desc->classVariable : "?");
    obj->cpc = NULL;
    obj->pci = NULL;
    *ppComObj = obj;  /* 回填 __comObj, 后续复用 */

    InterlockedIncrement(&g_vb6_cRef);
    return obj;
}

// Fix 188: 包装"宿主已拥有"的实例 —— 见 vb6comserver.h 声明. 释放到 0 时只回收
// 包装器并清空 __comObj 回填, 不调用 destroyFunc.
vb6_ComObject* vb6_ComObject_FromBorrowedInstance(const vb6_CoClassDesc* desc, void* instance) {
    void** ppComObj;
    vb6_ComObject* existing;
    vb6_ComObject* obj;
    if (!desc || !instance) return NULL;

    ppComObj = (void**)instance;
    existing = (vb6_ComObject*)*ppComObj;
    if (existing) {
        existing->vtable->AddRef(existing);
        return existing;
    }

    obj = (vb6_ComObject*)CoTaskMemAlloc(sizeof(vb6_ComObject));
    if (!obj) return NULL;

    obj->vtable = &g_ComObjectVtable;
    obj->refCount = 1;
    obj->desc = desc;
    obj->vb6Instance = instance;
    obj->ownsInstance = 0;
    P189("[189] FROM-BORROW W=%p inst=%p cls=%s\n", (void*)obj, (void*)instance,
         desc->classVariable ? desc->classVariable : "?");
    obj->cpc = NULL;
    obj->pci = NULL;
    *ppComObj = obj;

    InterlockedIncrement(&g_vb6_cRef);
    return obj;
}

// Fix 099: 按类变量名 (VB6 模块名, 如 "cHttpServerRouter") 查 coclass 描述.
// 供 dll_entry.c 的 Public 对象字段 getter 使用 (包装实例需要 desc).
// 只命中 coclass 表内的类 (cgen 仅收 MultiUse/SingleUse); 未命中返回 NULL,
// 此时 getter 返回 NULL dispatch — 客户端拿到 Nothing, 不会崩.
const vb6_CoClassDesc* vb6_FindCoClassDesc(const char* classVariable) {
    int i;
    if (!classVariable) return NULL;
    for (i = 0; i < g_vb6_coclassCount; i++) {
        if (g_vb6_coclasses[i].classVariable &&
            _stricmp(g_vb6_coclasses[i].classVariable, classVariable) == 0) {
            return &g_vb6_coclasses[i];
        }
    }
    return NULL;
}

// Fix 099: 从 IDispatch 取回其 VB6 类实例裸指针 (Public 对象字段的 Property Let/Set
// 桥接用: 把客户端传来的对象写进字段). 只认本 RTL 产出的 vb6_ComObject (用自身
// vtable 指针判定), 外部 COM 对象/非对象返回 NULL (字段置空, 与 VB6 传 Nothing 等效).
void* vb6_ComObject_GetInstance(void* pdisp) {
    vb6_ComObject* obj;
    if (!pdisp) return NULL;
    obj = (vb6_ComObject*)pdisp;
    if (obj->vtable != &g_ComObjectVtable) {
        /* Fix <vbeclipse> rev14: 宿主对象回落。
         * cgen 对 `Set x = <Variant 值>` 一律生成
         *   (vb6_cls_X*)vb6_ComObject_GetInstance(vb6_VariantToObjectVal(v))
         * 而 Controls.Item/Add 给出的是 UserControl **宿主 HWND** (不是 vb6_ComObject)。
         * vb6_UC_InstanceOf 只遍历 g_uc_recs 比较 hwnd, 对任意指针安全 (含 NULL /
         * 裸 HWND / 非宿主指针), 命中即返回该宿主的 vb6_cls_<UC>* 实例。 */
        { extern void* vb6_UC_InstanceOf(void* hwnd);
          void* inst = vb6_UC_InstanceOf(pdisp); if (inst) return inst; }
        if (getenv("C3_IV_TRACE")) {
            fprintf(stderr, "[IV] REJECT pdisp=%p vtable=%p (expect %p)\n",
                    pdisp, (void*)obj->vtable, (void*)&g_ComObjectVtable);
            fflush(stderr);
        }
        return NULL;
    }
    if (getenv("C3_IV_TRACE")) {
        fprintf(stderr, "[IV] OK pdisp=%p inst=%p\n", pdisp, obj->vb6Instance);
        fflush(stderr);
    }
    return obj->vb6Instance;
}

// ============================================================
// ExeComBridge 03: 工程类实例 → COM 调用实参 (VT_DISPATCH VARIANT*)
// ============================================================
// 触发场景: EXE/DLL 里 `New <本工程类>` 直接作为 COM 调用的实参
//   (VBMAN_DEMO: .Router.Reg "Demo", New bHello)
// cgen 原先对它生成通用打包 vb6_ComPackValue(vb6_cls_bHello_New()) —— 裸结构体
// 指针被当成 VT_DISPATCH 传出, 对端 (Dictionary 赋值 / Property Set) 取值或释放
// 时 AddRef, 把结构体首字段当 vtable 解引用 → 0xC0000005.
// 这里先经 vb6_ComObject_FromInstance 包装成真 IDispatch, 再转成 cgen 侧约定的
// VARIANT*. 引用计数: FromInstance 交给我们的 1 个引用直接转移给 VARIANT, 不额外
// AddRef —— VariantClear/Release 时正好归零, 包装器与实例随之释放.
void* vb6_ComPackVB6InstanceRaw(const char* classVariable, void* instance) {
    const vb6_CoClassDesc* desc;
    vb6_ComObject* obj;
    VARIANT* pv;

    if (!instance) return NULL;
    desc = vb6_FindCoClassDesc(classVariable);
    P189("[189] PACKRAW cls=%s inst=%p desc=%p\n", classVariable, instance, (void*)desc);
    // 未进 coclass 表 (Private / 非 MultiUse|SingleUse) 的类没有 IDispatch 方法表,
    // 包装也无从应答 GetIDsOfNames → 返回 NULL (等效 VB6 传 Nothing), 不冒充对象.
    if (!desc) return NULL;

    pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    if (!pv) return NULL;
    obj = vb6_ComObject_FromInstance(desc, instance);
    if (!obj) {
        free(pv);
        return NULL;
    }
    VariantInit(pv);
    pv->vt = VT_DISPATCH;
    pv->pdispVal = (IDispatch*)obj;  /* 引用已由 FromInstance 提供, 所有权转移给 VARIANT */
    return (void*)pv;
}
