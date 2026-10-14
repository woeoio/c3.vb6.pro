// vb6com_foreach.c - vb6com 模块拆分: For Each COM 集合枚举 (IEnumVARIANT)
// 由 vb6com.c 按 COM 调用层次拆分而来 (纯搬移, 零行为改动)

#include "vb6com.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "vb6com_internal.h"
#include "vb6forms.h"   /* Fix 112: 宿主 Controls 集合枚举 */


// ============================================================
// P22-11: For Each COM collection (IEnumVARIANT)
// ============================================================

// P25-fix4: For Each enumeration with safety limit
// Some COM collections (e.g. MSComctlLib.ImageList) have buggy IEnumVARIANT
// that never terminates (Next() always returns S_OK+fetched=1, cursor never
// advances). VB6 solves this by using Count from ICollection as the upper
// bound. We do the same: store Count and never iterate more than Count items.

typedef struct vb6_ForEachState {
    IEnumVARIANT* pEnum;
    int32_t totalCount;    // from collection.Count, -1 if unknown
    int32_t fetchedCount;  // how many items we've returned to caller (non-empty)
} vb6_ForEachState;

// 账 #257 第十五层 / #269 第十七层: `For Each x In <项目类集合>` —— 集合表达式的 C 面是**裸
// 工程类实例** (结构体首字段是 `__comObj` back-ptr 且常为 NULL, 没有真 vtable), 不是 IDispatch。
// vb6_ForEach_Init 会按 IDispatch 走 GetIDsOfNames(L"Count") ⇒ 读 `[NULL+0x14]` 当场 AV。
// 实测 ComCtlsDemo ListView.ctl:7223 RebuildListItems
//   `For Each ListItem In Me.ListItems`
//   → vb6_ForEach_Init(vb6_ListView_prop_get_ListItems((void*)me))
//   → av read target=0x14 @ vb6_ComGetIdsOfNames+0x87 (函数尾那一族同名类集合 6 处同形)。
// VB6 对工程类的 For Each 要的是它的 `NewEnum` 成员 (VB_UserMemId = -4)。发码侧
// (cgen_control.cpp L424) 直接生成 `vb6_ForEach_InitFromEnumIface((void*)vb6_<Cls>_NewEnum(<coll>))`,
// 把 NewEnum 返回的 **已合法构造的 IEnumVARIANT 接口指针** 喂进来 —— 它本来就是 IEnumVARIANT, 不必再 QI。
//
// 第十七层坑: 工程类 NewEnum 是 VTableHandle 版 IEnumVARIANT, 其 QueryInterface 体期望 `IID`
// 参数是「带 Data1 属性的 VB IID 对象」(生成码 vb6_ComGetIntProp((*IID), L"Data1")==132100),
// 而**不是**裸 GUID 指针 &IID_IEnumVARIANT。若拿 &IID_IEnumVARIANT 去 QI, 它会把 GUID 当 COM
// 对象解引用 → 读 0x20404 (IID_IEnumVARIANT.Data1=132100) → vb6_ComIsDispatchable 撞 NULL+0x404 崩
// (崩溃栈: ForEach_InitFromEnumIface+0x36 → VTableHandle_IEnumVARIANT_QueryInterface →
//  ComGetIntProp → ComGetProp → getDispid → ComIsDispatchable)。
// 既然 NewEnum 已经把合法 IEnumVARIANT 交出来了, 这里只 AddRef 接管即可, 绝不二次 QI。
void* vb6_ForEach_InitFromEnumIface(void* pEnumRaw) {
    if (!pEnumRaw) return NULL;
    IEnumVARIANT* pEnum = (IEnumVARIANT*)pEnumRaw;
    pEnum->lpVtbl->AddRef(pEnum);
    vb6_ForEachState* state = (vb6_ForEachState*)calloc(1, sizeof(vb6_ForEachState));
    if (!state) {
        pEnum->lpVtbl->Release(pEnum);
        return NULL;
    }
    state->pEnum = pEnum;
    state->totalCount = -1;   /* 计数由枚举器自身负责 (NewEnum 已把长度带进去了) */
    state->fetchedCount = 0;
    return (void*)state;
}

// For Each Init: Call _NewEnum (DISPID -4) on collection object to get IEnumVARIANT
// Also reads Count property for safety limit (works around buggy IEnumVARIANT)
// Returns vb6_ForEachState* (or NULL on failure)
void* vb6_ForEach_Init(void* disp) {
    if (!disp) return NULL;
    /* Fix 112: 宿主 Controls 集合 (vb6_ComGetObjectProp(Me,"Controls") 的合成对象)
     * 不是 IDispatch, 用原生枚举器迭代窗体子控件. */
    if (vb6_UC_ControlsIsCollection(disp)) {
        return vb6_UC_ControlsEnumInit(disp);
    }
    /* Fix 112c: RTL 内建 Collection (New Collection 的产物) */
    if (vb6_Collection_IsCollection(disp)) {
        return vb6_Collection_EnumInit(disp);
    }
    IDispatch* pDisp = (IDispatch*)disp;

    // Read Count property for safety limit
    int32_t count = -1;  // unknown
    {
        DISPID dispidCount = 0;
        OLECHAR* countName = L"Count";
        HRESULT hr2 = vb6_ComGetIdsOfNames(pDisp, countName, &dispidCount);
        if (SUCCEEDED(hr2) && dispidCount > 0) {
            DISPPARAMS dpC = { NULL, NULL, 0, 0 };
            VARIANT vCount;
            VariantInit(&vCount);
            hr2 = vb6_ComInvoke(pDisp, dispidCount,
                DISPATCH_PROPERTYGET | DISPATCH_METHOD,
                &dpC, &vCount, NULL, NULL);
            if (SUCCEEDED(hr2) && (V_VT(&vCount) == VT_I4 || V_VT(&vCount) == VT_I2)) {
                count = (V_VT(&vCount) == VT_I4) ? V_I4(&vCount) : (int32_t)V_I2(&vCount);
            }
            VariantClear(&vCount);
        }
    }

    // DISPID -4 is the standard _NewEnum dispid in VB6/Automation
    DISPPARAMS dp = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = vb6_ComInvoke(pDisp, (DISPID)-4,
        DISPATCH_METHOD | DISPATCH_PROPERTYGET,
        &dp, &result, NULL, NULL);

    if (FAILED(hr) || (V_VT(&result) != VT_UNKNOWN && V_VT(&result) != VT_DISPATCH)) {
        // Try named invocation as fallback
        OLECHAR* names[] = { L"_NewEnum" };
        DISPID dispid;
        hr = vb6_ComGetIdsOfNames(pDisp, names[0], &dispid);
        if (SUCCEEDED(hr)) {
            hr = vb6_ComInvoke(pDisp, dispid,
                DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                &dp, &result, NULL, NULL);
        }
    }

    if (FAILED(hr)) {
        VariantClear(&result);
        return NULL;
    }

    // Get IEnumVARIANT from the returned IUnknown/IDispatch
    IEnumVARIANT* pEnum = NULL;
    if (V_VT(&result) == VT_UNKNOWN && V_UNKNOWN(&result)) {
        hr = V_UNKNOWN(&result)->lpVtbl->QueryInterface(
            V_UNKNOWN(&result), &IID_IEnumVARIANT, (void**)&pEnum);
    } else if (V_VT(&result) == VT_DISPATCH && V_DISPATCH(&result)) {
        hr = V_DISPATCH(&result)->lpVtbl->QueryInterface(
            V_DISPATCH(&result), &IID_IEnumVARIANT, (void**)&pEnum);
    }

    VariantClear(&result);

    if (!pEnum) return NULL;

    // Allocate wrapper struct
    vb6_ForEachState* state = (vb6_ForEachState*)calloc(1, sizeof(vb6_ForEachState));
    if (!state) {
        pEnum->lpVtbl->Release(pEnum);
        return NULL;
    }
    state->pEnum = pEnum;
    state->totalCount = count;
    state->fetchedCount = 0;
    return (void*)state;
}

// For Each Next: Fetch one element from IEnumVARIANT
// Returns 1 if element fetched, 0 if enumeration complete
// P25: VB6 compat - skip VT_EMPTY/VT_NULL/VT_ERROR items
// P25-fix4: enforce Count limit (buggy IEnumVARIANT may never terminate)
int32_t vb6_ForEach_Next(void* enumPtr, VARIANT* outVar) {
    if (!enumPtr || !outVar) return 0;
    /* Fix 112: 宿主 Controls 集合的原生枚举器 (vb6_ForEach_Init 的宿主分支) */
    if (vb6_UC_ControlsIsCollection(*((void**)enumPtr))) {
        char hv[64];
        int32_t got = vb6_UC_ControlsEnumNext(enumPtr, hv);
        VariantInit(outVar);
        if (!got) return 0;
        vb6_Host_ToWinVariant(&hv, outVar);
        vb6_Host_ClearVariant(hv);
        return 1;
    }
    /* Fix 112c: RTL 内建 Collection 的原生枚举器 */
    if (vb6_Collection_IsCollection(*((void**)enumPtr))) {
        char hv[64];
        int32_t got = vb6_Collection_EnumNext(enumPtr, hv);
        VariantInit(outVar);
        if (!got) return 0;
        vb6_Host_ToWinVariant(&hv, outVar);
        vb6_Host_ClearVariant(hv);
        return 1;
    }
    vb6_ForEachState* state = (vb6_ForEachState*)enumPtr;
    IEnumVARIANT* pEnum = state->pEnum;

    for (;;) {
        // Safety: never iterate more than Count items (if known)
        if (state->totalCount >= 0 && state->fetchedCount >= state->totalCount) {
            VariantInit(outVar);
            return 0;
        }
        VariantInit(outVar);
        ULONG fetched = 0;
        HRESULT hr = pEnum->lpVtbl->Next(pEnum, 1, outVar, &fetched);
        if (FAILED(hr) || fetched == 0) {
            VariantClear(outVar);
            return 0;
        }
        // VB6: skip items that are Empty, Null, or Error (e.g. Nothing)
        VARTYPE vt = V_VT(outVar);
        if (vt == VT_EMPTY || vt == VT_NULL || vt == VT_ERROR) {
            VariantClear(outVar);
            continue;
        }
        state->fetchedCount++;
        return 1;
    }
}

// For Each Release: Release IEnumVARIANT and free wrapper
void vb6_ForEach_Release(void* enumPtr) {
    if (!enumPtr) return;
    /* Fix 112/112c: 宿主枚举句柄 (Controls 集合 / 内建 Collection) 不是 vb6_ForEachState
     * —— 首元素是集合指针而非 IEnumVARIANT*, 当 state->pEnum 解引用会取到结构体自身
     * 的字节当 vtable → 0xC0000005. 宿主句柄没有需要 Release 的 COM 接口, 直接释放. */
    if (vb6_UC_ControlsIsCollection(*((void**)enumPtr)) ||
        vb6_Collection_IsCollection(*((void**)enumPtr))) {
        free(enumPtr);
        return;
    }
    vb6_ForEachState* state = (vb6_ForEachState*)enumPtr;
    state->pEnum->lpVtbl->Release(state->pEnum);
    free(state);
}
