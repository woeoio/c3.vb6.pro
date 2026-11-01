// vb6com_wrap.c - vb6com 模块拆分: 类型化调用包装 + 属性便捷读取 + QI/vtable 直调
// 由 vb6com.c 按 COM 调用层次拆分而来 (纯搬移, 零行为改动)

#include "vb6com.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "vb6com_internal.h"


// ============================================================
// 一体化COM辅助函数 (P6.2 cgen直接使用, 内部处理临时VARIANT清理)
// ============================================================

// COM方法调用, 返回对象 (IDispatch*) — 内部完成Unpack+VarFree
// methodName == NULL: **默认成员**调用 (`coll(idx)`)。VB6 把它编到 DISPID_VALUE(0),
// 从不按名字 "Item" 找; 详见 vb6com_invoke.c 的 vb6_ComCallObjectDefault。
void* vb6_ComCallObject(void* disp, const wchar_t* methodName,
                        void* args_void, int32_t argc) {
    if (!methodName) return vb6_ComCallObjectDefault(disp, args_void, argc);
    VARIANT* pv = (VARIANT*)vb6_ComCall(disp, methodName, args_void, argc);
    if (!pv) return NULL;
    void* obj = NULL;
    if (pv->vt == VT_DISPATCH) {
        // czUI fix: Pack 侧可能包了宿主包装器 (字体/集合等 RTL 伪对象),
        // 解包还原原始对象 — ctl 代码按结构体字段直接访问
        extern void* vb6_UC_UnwrapHost(void* obj);
        obj = vb6_UC_UnwrapHost((void*)pv->pdispVal);  // 转移引用所有权
    } else if (pv->vt == VT_UNKNOWN) {
        pv->punkVal->lpVtbl->QueryInterface(pv->punkVal, &IID_IDispatch, &obj);
        // QueryInterface增加了refcount, 需要Release原引用
        // 但原引用在VarFree中不会Release, 所以不需要额外操作
    } else if (pv->vt != VT_EMPTY && pv->vt != VT_NULL) {
        // 尝试转换为IDispatch
        VARIANT vObj;
        VariantInit(&vObj);
        if (SUCCEEDED(VariantChangeType(&vObj, pv, 0, VT_DISPATCH))) {
            obj = (void*)vObj.pdispVal;
        }
        VariantClear(&vObj);
    }
    if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_COM] CallObject: disp=%p method=%ls vt=%d -> obj=%p (pre-unwrap=%p)\n",
                disp, methodName, pv ? pv->vt : -1, obj,
                (pv && pv->vt == VT_DISPATCH) ? (void*)pv->pdispVal : NULL);
        fflush(stderr);
    }
    vb6_ComVarFree(pv);  // 仅释放结构体, 不Release pdispVal
    return obj;
}

// COM方法调用, 返回BSTR — 内部完成Unpack+VarClear
wchar_t* vb6_ComCallBSTR(void* disp, const wchar_t* methodName,
                         void* args_void, int32_t argc) {
    VARIANT* pv = (VARIANT*)vb6_ComCall(disp, methodName, args_void, argc);
    if (!pv) return NULL;
    BSTR result = vb6_ComUnpackBSTR(pv);
    vb6_ComVarClear(pv);  // 安全: UnpackBSTR已复制字符串内容
    return result;
}

// COM方法调用, 返回int32_t — 内部完成Unpack+VarClear
int32_t vb6_ComCallInt(void* disp, const wchar_t* methodName,
                       void* args_void, int32_t argc) {
    VARIANT* pv = (VARIANT*)vb6_ComCall(disp, methodName, args_void, argc);
    if (!pv) return 0;
    int32_t result = vb6_ComUnpackInt(pv);
    vb6_ComVarClear(pv);
    return result;
}

// COM方法调用, 返回double — 内部完成Unpack+VarClear
double vb6_ComCallDouble(void* disp, const wchar_t* methodName,
                         void* args_void, int32_t argc) {
    VARIANT* pv = (VARIANT*)vb6_ComCall(disp, methodName, args_void, argc);
    if (!pv) return 0.0;
    double result = vb6_ComUnpackDouble(pv);
    vb6_ComVarClear(pv);
    return result;
}

// COM属性Get, 返回BSTR — 内部完成Unpack+VarClear
wchar_t* vb6_ComGetStringProp(void* disp, const wchar_t* propName) {
    VARIANT* pv = (VARIANT*)vb6_ComGetProp(disp, propName);
    if (!pv) {
        if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
            fprintf(stderr, "[C3_GSPROP] '%ls' -> NULL (调用失败)\n", propName ? propName : L"?");
        return NULL;
    }
    BSTR result = vb6_ComUnpackBSTR(pv);
    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_GSPROP] '%ls' srcVT=%d -> '%ls'\n", propName ? propName : L"?",
                (int)pv->vt, result ? result : L"(null)");
    vb6_ComVarClear(pv);  // 安全: UnpackBSTR已复制字符串内容
    return result;
}

// COM属性Get, 返回int32_t — 内部完成Unpack+VarClear
int32_t vb6_ComGetIntProp(void* disp, const wchar_t* propName) {
    VARIANT* pv = (VARIANT*)vb6_ComGetProp(disp, propName);
    if (!pv) return 0;
    int32_t result = vb6_ComUnpackInt(pv);
    vb6_ComVarClear(pv);
    return result;
}

// COM属性Get, 返回intptr_t (LongPtr: 句柄/指针) — 兼容 32/64 位整数变体
intptr_t vb6_ComGetLongPtrProp(void* disp, const wchar_t* propName) {
    VARIANT* pv = (VARIANT*)vb6_ComGetProp(disp, propName);
    if (!pv) return 0;
    intptr_t result;
    if (pv->vt == VT_I8) result = (intptr_t)pv->llVal;
    else if (pv->vt == VT_UI8) result = (intptr_t)pv->ullVal;
    else result = (intptr_t)vb6_ComUnpackInt(pv);
    vb6_ComVarClear(pv);
    return result;
}

// COM属性Get, 返回double — 内部完成Unpack+VarClear
double vb6_ComGetDoubleProp(void* disp, const wchar_t* propName) {
    VARIANT* pv = (VARIANT*)vb6_ComGetProp(disp, propName);
    if (!pv) return 0.0;
    double result = vb6_ComUnpackDouble(pv);
    vb6_ComVarClear(pv);
    return result;
}

// COM属性Get, 返回对象 — 内部完成Unpack+VarFree
void* vb6_ComGetObjectProp(void* disp, const wchar_t* propName) {
    VARIANT* pv = (VARIANT*)vb6_ComGetProp(disp, propName);
    if (!pv) return NULL;
    void* obj = NULL;
    if (pv->vt == VT_DISPATCH) {
        obj = (void*)pv->pdispVal;
    } else if (pv->vt == VT_UNKNOWN) {
        pv->punkVal->lpVtbl->QueryInterface(pv->punkVal, &IID_IDispatch, &obj);
    } else if (pv->vt != VT_EMPTY && pv->vt != VT_NULL) {
        VARIANT vObj;
        VariantInit(&vObj);
        if (SUCCEEDED(VariantChangeType(&vObj, pv, 0, VT_DISPATCH))) {
            obj = (void*)vObj.pdispVal;
        }
        VariantClear(&vObj);
    }
    if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_COM] GetObjectProp obj: disp=%p prop=%ls vt=%d -> obj=%p\n",
                disp, propName, pv ? pv->vt : -1, obj); fflush(stderr);
    }
    vb6_ComVarFree(pv);  // 仅释放结构体, 不Release pdispVal
    return obj;
}

// ============================================================
// 参数化属性Get (带参数, 如Dictionary.Item(key))
// ============================================================

// 参数化属性Get→BSTR
wchar_t* vb6_ComGetPropertyString(void* disp, const wchar_t* propName,
                                   void* args, int32_t argc) {
    VARIANT* pv = (VARIANT*)vb6_ComGetPropArg(disp, propName, args, argc);
    if (!pv) return NULL;
    BSTR result = vb6_ComUnpackBSTR(pv);
    vb6_ComVarClear(pv);
    return result;
}

// 参数化属性Get→int32_t
int32_t vb6_ComGetPropertyInt(void* disp, const wchar_t* propName,
                               void* args, int32_t argc) {
    VARIANT* pv = (VARIANT*)vb6_ComGetPropArg(disp, propName, args, argc);
    if (!pv) return 0;
    int32_t result = vb6_ComUnpackInt(pv);
    vb6_ComVarClear(pv);
    return result;
}

// 参数化属性Get→double
double vb6_ComGetPropertyDouble(void* disp, const wchar_t* propName,
                                 void* args, int32_t argc) {
    VARIANT* pv = (VARIANT*)vb6_ComGetPropArg(disp, propName, args, argc);
    if (!pv) return 0.0;
    double result = vb6_ComUnpackDouble(pv);
    vb6_ComVarClear(pv);
    return result;
}

// 参数化属性Get→对象
void* vb6_ComGetPropertyObject(void* disp, const wchar_t* propName,
                                void* args, int32_t argc) {
    VARIANT* pv = (VARIANT*)vb6_ComGetPropArg(disp, propName, args, argc);
    if (!pv) return NULL;
    void* obj = NULL;
    if (pv->vt == VT_DISPATCH) {
        obj = (void*)pv->pdispVal;
    } else if (pv->vt == VT_UNKNOWN) {
        pv->punkVal->lpVtbl->QueryInterface(pv->punkVal, &IID_IDispatch, &obj);
    } else if (pv->vt != VT_EMPTY && pv->vt != VT_NULL) {
        VARIANT vObj;
        VariantInit(&vObj);
        if (SUCCEEDED(VariantChangeType(&vObj, pv, 0, VT_DISPATCH))) {
            obj = (void*)vObj.pdispVal;
        }
        VariantClear(&vObj);
    }
    vb6_ComVarFree(pv);
    return obj;
}

// 参数化属性Get→VARIANT (保留原样, 用于Variant变量赋值)
// 注意: 返回的VARIANT*需要调用方用vb6_ComVarClear释放
void* vb6_ComGetPropertyVariant(void* disp, const wchar_t* propName,
                                 void* args, int32_t argc) {
    return vb6_ComGetPropArg(disp, propName, args, argc);
}

// ============================================================
// P6.3: COM前期绑定运行时 (vtable直接调用)

// ============================================================
// P6.3: COM前期绑定运行时 (vtable直接调用)
// ============================================================

void* vb6_ComQI(void* obj, const char* iidStr) {
    if (!obj || !iidStr) return NULL;
    IID iid;
    wchar_t iidWide[64];
    int len = 0;
    for (; iidStr[len] && len < 63; len++) iidWide[len] = (wchar_t)iidStr[len];
    iidWide[len] = 0;
    HRESULT hr = IIDFromString(iidWide, &iid);
    if (FAILED(hr)) return NULL;
    void* pv = NULL;
    hr = vb6_UnknownQI(obj, &iid, &pv);
    if (FAILED(hr)) return NULL;
    return pv;
}

void* vb6_ComCreateTyped(const wchar_t* progId, const char* iidStr) {
    void* disp = vb6_CreateObject(progId);
    if (!disp) return NULL;
    void* iface = vb6_ComQI(disp, iidStr);
    vb6_UnknownRelease(disp);
    return iface;
}

void vb6_ComReleaseTyped(void** objPtr) {
    if (objPtr && *objPtr) {
        vb6_UnknownRelease(*objPtr);
        *objPtr = NULL;
    }
}

void vb6_ComVtableCallVoid(void* obj, int32_t vtIndex, ...) {
    if (!obj) return;
    void** vtable = *(void***)obj;
    if (!vtable || vtIndex < 0) return;
}

wchar_t* vb6_ComVtableGetBSTR(void* obj, int32_t vtIndex, ...) {
    if (!obj) return NULL;
    void** vtable = *(void***)obj;
    if (!vtable || vtIndex < 0) return NULL;
    typedef HRESULT(__stdcall* GetBSTRMethod)(void*, BSTR*);
    GetBSTRMethod fn = (GetBSTRMethod)vtable[vtIndex];
    BSTR result = NULL;
    fn(obj, &result);
    if (!result) return NULL;
    return result;  // Already OLE BSTR, caller uses vb6_BSTR_Free (SysFreeString)
}

int32_t vb6_ComVtableGetInt(void* obj, int32_t vtIndex, ...) {
    if (!obj) return 0;
    void** vtable = *(void***)obj;
    if (!vtable || vtIndex < 0) return 0;
    typedef HRESULT(__stdcall* GetVarMethod)(void*, VARIANT*);
    GetVarMethod fn = (GetVarMethod)vtable[vtIndex];
    VARIANT v; VariantInit(&v); fn(obj, &v);
    int32_t result = 0;
    if (v.vt == VT_I4) result = v.lVal;
    else if (v.vt == VT_I2) result = v.iVal;
    else if (v.vt == VT_BOOL) result = (v.boolVal != 0) ? -1 : 0;
    else if (v.vt == VT_R8) result = (int32_t)v.dblVal;
    else if (v.vt == VT_EMPTY || v.vt == VT_NULL) result = 0;
    else { VARIANT vO; VariantInit(&vO); if(SUCCEEDED(VariantChangeType(&vO,&v,0,VT_I4))) result=vO.lVal; VariantClear(&vO); }
    VariantClear(&v);
    return result;
}

double vb6_ComVtableGetDouble(void* obj, int32_t vtIndex, ...) {
    if (!obj) return 0.0;
    void** vtable = *(void***)obj;
    if (!vtable || vtIndex < 0) return 0.0;
    typedef HRESULT(__stdcall* GetVarMethod)(void*, VARIANT*);
    GetVarMethod fn = (GetVarMethod)vtable[vtIndex];
    VARIANT v; VariantInit(&v); fn(obj, &v);
    double result = 0.0;
    if (v.vt == VT_R8) result = v.dblVal;
    else if (v.vt == VT_R4) result = v.fltVal;
    else if (v.vt == VT_I4) result = (double)v.lVal;
    else if (v.vt == VT_I2) result = (double)v.iVal;
    else { VARIANT vO; VariantInit(&vO); if(SUCCEEDED(VariantChangeType(&vO,&v,0,VT_R8))) result=vO.dblVal; VariantClear(&vO); }
    VariantClear(&v);
    return result;
}

void* vb6_ComVtableGetObject(void* obj, int32_t vtIndex, ...) {
    if (!obj) return NULL;
    void** vtable = *(void***)obj;
    if (!vtable || vtIndex < 0) return NULL;
    typedef HRESULT(__stdcall* GetVarMethod)(void*, VARIANT*);
    GetVarMethod fn = (GetVarMethod)vtable[vtIndex];
    VARIANT v; VariantInit(&v); fn(obj, &v);
    void* result = NULL;
    if (v.vt == VT_DISPATCH && v.pdispVal) result = v.pdispVal;
    else if (v.vt == VT_UNKNOWN && v.punkVal) result = v.punkVal;
    if (v.vt == VT_DISPATCH || v.vt == VT_UNKNOWN) v.pdispVal = NULL;
    VariantClear(&v);
    return result;
}

void* vb6_ComVtableGetVoid(void* obj, int32_t vtIndex, ...) {
    if (!obj) return NULL;
    void** vtable = *(void***)obj;
    if (!vtable || vtIndex < 0) return NULL;
    typedef HRESULT(__stdcall* GetVarMethod)(void*, VARIANT*);
    GetVarMethod fn = (GetVarMethod)vtable[vtIndex];
    VARIANT* pv = (VARIANT*)malloc(sizeof(VARIANT));
    VariantInit(pv); fn(obj, pv);
    return (void*)pv;
}
