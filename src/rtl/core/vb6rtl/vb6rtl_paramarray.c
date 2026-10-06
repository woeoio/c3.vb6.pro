// vb6rtl_paramarray.c - VB6 运行时库: ParamArray 家族：PA_* 变参支持 + CallByName
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 4182~4430 行

#include "vb6rtl.h"
#include "vb6forms.h"   /* Fix 112: 宿主对象模型 (窗体/控件/集合/字体) */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <math.h>
#include <stdarg.h>
#include <wchar.h>
#include <wctype.h>
#include <time.h>
#ifdef _WIN32
#include <direct.h>
#include <io.h>
#include <oleauto.h>
#include <olectl.h>
#include <windows.h>
#endif

// P24-08: MessageBoxW (user32) + GetConsoleWindow (kernel32)
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "kernel32.lib")

// ============================================================
// COM互操作前向声明 (实现在vb6com.c中，避免vb6_VARIANT类型冲突)
// ============================================================
extern void* vb6_CreateObject(const wchar_t* progId);
extern void* vb6_GetObject(const wchar_t* pathName, const wchar_t* progId);
extern int32_t vb6_IsNothing(void* obj);
extern void vb6_ReleaseObject(void** objPtr);
extern void* vb6_ComCall(void* disp, const wchar_t* methodName, void* args, int32_t argc);
extern void* vb6_ComCallByDispid(void* disp, int32_t dispid, void* args, int32_t argc);
extern void* vb6_ComGetProp(void* disp, const wchar_t* propName);
extern void vb6_ComSetProp(void* disp, const wchar_t* propName, void* value);
extern void vb6_ComSetRef(void* disp, const wchar_t* propName, void* objRef);
extern void vb6_ComInit(void);
extern void vb6_ComExit(void);

// ============================================================
// ParamArray runtime support (P14.1.5)
// ============================================================

SAFEARRAY* vb6_PA_Create(int32_t count) {
    if (count <= 0) return NULL;
    SAFEARRAYBOUND bound;
    bound.lLbound = 0;
    bound.cElements = (ULONG)count;
    SAFEARRAY* psa = SafeArrayCreate(VT_VARIANT, 1, &bound);
    return psa;
}

void vb6_PA_Destroy(SAFEARRAY* psa) {
    if (psa) SafeArrayDestroy(psa);
}

void vb6_PA_SetVariant(SAFEARRAY* psa, int32_t index, VARIANT* pv) {
    if (!psa || !pv) return;
    long idx = (long)index;
    SafeArrayPutElement(psa, &idx, pv);
}

void vb6_PA_SetLong(SAFEARRAY* psa, int32_t index, int32_t val) {
    if (!psa) return;
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_I4;
    v.lVal = val;
    long idx = (long)index;
    SafeArrayPutElement(psa, &idx, &v);
}

/* Fix 082: x64-safe VarPtr parameter - store as VT_I8 (LongPtr) */
void vb6_PA_SetLongPtr(SAFEARRAY* psa, int32_t index, intptr_t val) {
    if (!psa) return;
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_I8;
    v.llVal = (LONGLONG)val;
    long idx = (long)index;
    SafeArrayPutElement(psa, &idx, &v);
}

void vb6_PA_SetDouble(SAFEARRAY* psa, int32_t index, double val) {
    if (!psa) return;
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_R8;
    v.dblVal = val;
    long idx = (long)index;
    SafeArrayPutElement(psa, &idx, &v);
}

void vb6_PA_SetBSTR(SAFEARRAY* psa, int32_t index, BSTR val) {
    if (!psa) return;
    VARIANT v;
    VariantInit(&v);
    v.vt = VT_BSTR;
    v.bstrVal = SysAllocString(val);  // SafeArrayPutElement doesn't copy BSTR
    long idx = (long)index;
    SafeArrayPutElement(psa, &idx, &v);
    // v.bstrVal is now owned by the array element
}

// Fix <vbeclipse> rev37: 返回 **vb6_VARIANT** 而非 Windows VARIANT。
//   生成码把 ParamArray 的 Variant 元素直接喂给 vb6_VariantFromValue, 而那个
//   _Generic 只认 vb6_VARIANT (C3 自己的同布局结构, vb6rtl_variant.h:35) ——
//   喂真正的 VARIANT 会落到 `default: vb6_VariantObject` (要 void*), 报
//   C2172 "不能从类型常数转换" (实测 paratest 夹具 4 处)。两者布局逐平台一致
//   (同文件已注明), 直接转换零成本。
vb6_VARIANT vb6_PA_GetVariant(SAFEARRAY* psa, int32_t index) {
    VARIANT v;
    VariantInit(&v);
    if (!psa) { VARIANT e; VariantInit(&e); return *(vb6_VARIANT*)&e; }
    long idx = (long)index;
    SafeArrayGetElement(psa, &idx, &v);
    // 转交所有权: 栈上 v 已无栈外引用, 交出其内部缓冲 (BSTR/数组) 由调用方 Clear。
    return *(vb6_VARIANT*)&v;
}

int32_t vb6_PA_GetLong(SAFEARRAY* psa, int32_t index) {
    if (!psa) return 0;
    VARIANT v;
    VariantInit(&v);
    long idx = (long)index;
    SafeArrayGetElement(psa, &idx, &v);
    int32_t result = 0;
    if (v.vt == VT_I4) result = v.lVal;
    else if (v.vt == VT_I2) result = (int32_t)v.iVal;
    else if (v.vt == VT_R8) result = (int32_t)v.dblVal;
    VariantClear(&v);
    return result;
}

double vb6_PA_GetDouble(SAFEARRAY* psa, int32_t index) {
    if (!psa) return 0.0;
    VARIANT v;
    VariantInit(&v);
    long idx = (long)index;
    SafeArrayGetElement(psa, &idx, &v);
    double result = 0.0;
    if (v.vt == VT_R8) result = v.dblVal;
    else if (v.vt == VT_I4) result = (double)v.lVal;
    else if (v.vt == VT_BSTR && v.bstrVal) result = wcstod(v.bstrVal, NULL);
    VariantClear(&v);
    return result;
}

BSTR vb6_PA_GetBSTR(SAFEARRAY* psa, int32_t index) {
    if (!psa) return NULL;
    VARIANT v;
    VariantInit(&v);
    long idx = (long)index;
    SafeArrayGetElement(psa, &idx, &v);
    BSTR result = NULL;
    if (v.vt == VT_BSTR && v.bstrVal) {
        result = SysAllocString(v.bstrVal);
    } else if (v.vt == VT_I4) {
        wchar_t buf[32];
        _ltow_s(v.lVal, buf, 32, 10);
        result = SysAllocString(buf);
    } else if (v.vt == VT_R8) {
        wchar_t buf[64];
        swprintf_s(buf, 64, L"%g", v.dblVal);
        result = SysAllocString(buf);
    }
    VariantClear(&v);
    return result;
}

int32_t vb6_IsMissing(SAFEARRAY* psa) {
    // IsMissing returns True if no arguments were passed to ParamArray
    return (psa == NULL) ? -1 : 0;  // VB6 True = -1, False = 0
}

int32_t vb6_PA_UBound(SAFEARRAY* psa) {
    if (!psa) return -1;  // empty ParamArray: UBound returns -1 (VB6 behavior)
    long ubound = 0;
    SafeArrayGetUBound(psa, 1, &ubound);
    return (int32_t)ubound;
}

int32_t vb6_PA_LBound(SAFEARRAY* psa) {
    if (!psa) return 0;
    long lbound = 0;
    SafeArrayGetLBound(psa, 1, &lbound);
    return (int32_t)lbound;
}

// P14.3.5: CallByName - 按名称动态调用方法/属性
// calltype: 1=VbLet(set property), 2=VbMethod(call method), 3=VbGet(get property)
vb6_VARIANT vb6_CallByName(void* obj, const wchar_t* procName, int32_t callType,
                           void* args, int32_t argc) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    result.vt = (vb6_vartype)VT_EMPTY;

    if (!obj || !procName) return result;

    /* Fix 112: 宿主对象 (窗体/控件 HWND, Font 代理) 用 Win32 语义应答 CallByName */
    if (vb6_Host_IsHostObject(obj)) {
        if (callType == 3) {           /* VbGet */
            vb6_Host_GetProp(obj, procName, &result);
        } else if (callType == 1) {    /* VbLet: 首个实参即要写入的值 */
            void** a = (void**)args;
            if (argc >= 1 && a && a[0]) {
                char hv[64];
                vb6_Host_FromWinVariant(a[0], hv);
                vb6_Host_SetProp(obj, procName, hv);
                vb6_Host_ClearVariant(hv);
            }
        }
        return result;
    }

    IDispatch* disp = (IDispatch*)obj;
    DISPID dispid = 0;
    HRESULT hr;

    // Get DISPID
    hr = disp->lpVtbl->GetIDsOfNames(disp, &IID_NULL, (LPOLESTR*)&procName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return result;

    // Determine INVOKE_KIND from calltype
    INVOKEKIND invKind;
    switch (callType) {
        case 1: invKind = DISPATCH_PROPERTYPUT; break;  // VbLet
        case 2: invKind = DISPATCH_METHOD; break;        // VbMethod
        case 3: invKind = DISPATCH_PROPERTYGET; break;   // VbGet
        default: invKind = DISPATCH_METHOD; break;
    }

    // Build DISPPARAMS from args array
    // Fix 040d: args is now VARIANT** (array of pointers to VARIANTs),
    // same convention as vb6_ComCall. Each element is a packed VARIANT*
    // allocated by vb6_ComPackBSTR/Int/Double/Object/Value.
    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    VARIANT* pArgs = NULL;
    VARIANT** argArr = (VARIANT**)args;

    if (argc > 0 && args) {
        pArgs = (VARIANT*)calloc(argc, sizeof(VARIANT));
        if (pArgs) {
            for (int32_t i = 0; i < argc; i++) {
                VariantInit(&pArgs[argc - 1 - i]);  /* reverse for COM */
                if (argArr[i]) {
                    pArgs[argc - 1 - i] = *argArr[i];
                }
            }
            dp.cArgs = (UINT)argc;
            dp.rgvarg = pArgs;
            // For PropertyPut, also set named arg
            if (invKind == DISPATCH_PROPERTYPUT) {
                DISPID putId = DISPID_PROPERTYPUT;
                dp.cNamedArgs = 1;
                dp.rgdispidNamedArgs = &putId;
            }
        }
    }

    // Invoke
    VARIANT retVal;
    VariantInit(&retVal);
    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    hr = disp->lpVtbl->Invoke(disp, dispid, &IID_NULL, LOCALE_USER_DEFAULT,
                              invKind, &dp, &retVal, &excep, &argErr);

    if (SUCCEEDED(hr)) {
        result.vt = (vb6_vartype)retVal.vt;
        // Copy value from COM VARIANT to vb6_VARIANT
        switch (retVal.vt) {
            case VT_I2: result.iVal = retVal.iVal; break;
            case VT_I4: result.lVal = retVal.lVal; break;
            case VT_R4: result.fltVal = retVal.fltVal; break;
            case VT_R8: result.dblVal = retVal.dblVal; break;
            case VT_BSTR: result.bstrVal = retVal.bstrVal; VariantInit(&retVal); break;
            case VT_DISPATCH: result.pdispVal = retVal.pdispVal; VariantInit(&retVal); break;
            case VT_BOOL: result.boolVal = retVal.boolVal; break;
            case VT_UI1: result.bVal = retVal.bVal; break;
            default: result.lVal = retVal.lVal; break;
        }
    }

    // Cleanup
    // Fix 040d: cleanup packed VARIANT args (same pattern as vb6_ComCall)
    if (pArgs) {
        for (UINT i = 0; i < dp.cArgs; i++) {
            VariantClear(&pArgs[i]);
        }
        free(pArgs);
    }
    if (argArr) {
        for (int32_t i = 0; i < argc; i++) {
            if (argArr[i]) {
                free(argArr[i]);  /* free packed VARIANT struct */
            }
        }
    }
    VariantClear(&retVal);
    if (excep.bstrSource) SysFreeString(excep.bstrSource);
    if (excep.bstrDescription) SysFreeString(excep.bstrDescription);
    if (excep.bstrHelpFile) SysFreeString(excep.bstrHelpFile);

    return result;
}


