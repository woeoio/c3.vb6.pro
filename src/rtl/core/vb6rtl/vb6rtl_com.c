// vb6rtl_com.c - VB6 运行时库: COM 家族：IEnumVARIANT 枚举 + COM 结果转换 + Variant 数组 + LoadPicture/SavePicture/LoadForm
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 4699~4793 行
//   原第 4955~5327 行

#include "vb6rtl.h"
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
// P20-02: IEnumVARIANT support for COM For Each
// ============================================================

void* vb6_ComNewEnum(void* disp) {
    if (!disp) return NULL;
    IDispatch* pDisp = (IDispatch*)disp;

    // _NewEnum has DISPID = -4 (DISPID_NEWENUM)
    DISPID dispidNewEnum = -4;
    OLECHAR* wszNewEnum = L"_NewEnum";
    HRESULT hr = pDisp->lpVtbl->GetIDsOfNames(pDisp, &IID_NULL, &wszNewEnum, 1,
                                                LOCALE_USER_DEFAULT, &dispidNewEnum);
    if (FAILED(hr)) {
        // Try known DISPID directly
        dispidNewEnum = -4;
    }

    DISPPARAMS dp = {0};
    VARIANT result;
    VariantInit(&result);
    EXCEPINFO exInfo = {0};

    hr = pDisp->lpVtbl->Invoke(pDisp, dispidNewEnum, &IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET | DISPATCH_METHOD,
                                &dp, &result, &exInfo, NULL);
    if (FAILED(hr)) return NULL;

    // QI for IEnumVARIANT
    IEnumVARIANT* pEnum = NULL;
    if (result.vt == VT_UNKNOWN) {
        hr = result.punkVal->lpVtbl->QueryInterface(result.punkVal, &IID_IEnumVARIANT, (void**)&pEnum);
        result.punkVal->lpVtbl->Release(result.punkVal);
    } else if (result.vt == VT_DISPATCH) {
        hr = result.pdispVal->lpVtbl->QueryInterface(result.pdispVal, &IID_IEnumVARIANT, (void**)&pEnum);
        result.pdispVal->lpVtbl->Release(result.pdispVal);
    } else {
        VariantClear(&result);
        return NULL;
    }
    return (void*)pEnum;
}

int vb6_EnumNext(void* penum, vb6_VARIANT* outElem) {
    if (!penum || !outElem) return 0;
    IEnumVARIANT* pEnum = (IEnumVARIANT*)penum;

    VARIANT varElem;
    VariantInit(&varElem);
    ULONG fetched = 0;
    HRESULT hr = pEnum->lpVtbl->Next(pEnum, 1, &varElem, &fetched);

    if (hr != S_OK || fetched == 0) {
        VariantClear(&varElem);
        return 0;
    }

    // Convert COM VARIANT to vb6_VARIANT
    memset(outElem, 0, sizeof(vb6_VARIANT));
    switch (varElem.vt) {
        case VT_EMPTY:  outElem->vt = vb6_vtEmpty; break;
        case VT_NULL:   outElem->vt = vb6_vtNull; break;
        case VT_I2:     outElem->vt = vb6_vtInteger; outElem->iVal = varElem.iVal; break;
        case VT_I4:     outElem->vt = vb6_vtLong; outElem->lVal = varElem.lVal; break;
        case VT_R4:     outElem->vt = vb6_vtSingle; outElem->fltVal = varElem.fltVal; break;
        case VT_R8:     outElem->vt = vb6_vtDouble; outElem->dblVal = varElem.dblVal; break;
        case VT_BSTR:   outElem->vt = vb6_vtBSTR; outElem->bstrVal = varElem.bstrVal; break;
        case VT_BOOL:   outElem->vt = vb6_vtBoolean; outElem->boolVal = varElem.boolVal ? -1 : 0; VariantClear(&varElem); break;
        case VT_DISPATCH: outElem->vt = vb6_vtDispatch; outElem->pdispVal = varElem.pdispVal; break;
        case VT_DATE:   outElem->vt = vb6_vtDate; outElem->dblVal = varElem.date; VariantClear(&varElem); break;
        default: {
            // Convert to BSTR for unknown types
            VariantChangeType(&varElem, &varElem, 0, VT_BSTR);
            if (varElem.vt == VT_BSTR) {
                outElem->vt = vb6_vtBSTR;
                outElem->bstrVal = varElem.bstrVal;
            } else {
                VariantClear(&varElem);
            }
            return 1;
        }
    }
    // Cleanup for types we didn't transfer ownership of
    if (varElem.vt != VT_BSTR && varElem.vt != VT_DISPATCH && varElem.vt != VT_UNKNOWN) {
        VariantClear(&varElem);
    }
    return 1;
}

void vb6_EnumRelease(void* penum) {
    if (!penum) return;
    IEnumVARIANT* pEnum = (IEnumVARIANT*)penum;
    pEnum->lpVtbl->Release(pEnum);
}

// ============================================================
// LoadPicture 相关
// 纪律: LoadPicture 返回的是 **VB6 的 StdPicture 对象** (实现 IPicture + IPictureDisp,
// 即 IDispatch*), 是一个**活着的 COM 对象** —— 不是"Release 掉只剩句柄"的残骸。
// 调用方 (ImageList.ListImages.Add 等) 要拿它 AddRef 后长期持有, 镜像 VB6 里
// ListImage.Picture 会 keep 一份引用的行为。谁 AddRef 谁在不用时 vb6_ReleasePicture。
// ============================================================

// OleInitialize 必须至少成功调过一次, 否则 OleLoadPicturePath 静默失败。
// 0=还没试过 / 1=OleInitialize 成 / 2=退化为 CoInitialize
static int vb6_OleEnsureInit(void) {
    static int oleInited = 0;
    if (oleInited) return oleInited;
    if (SUCCEEDED(OleInitialize(NULL))) {
        oleInited = 1;
    } else {
        CoInitialize(NULL);
        oleInited = 2;
    }
    return oleInited;
}

// P21-18: LoadPictureEx — OleLoadPicturePath for all image types (BMP/ICO/EMF/WMF/JPG/GIF/PNG)
// 返回活着的 IPicture* (带一次引用); 失败返回 NULL。用完请 vb6_ReleasePicture。
void* vb6_LoadPictureEx(BSTR pathname) {
    if (!pathname) return NULL;
    vb6_OleEnsureInit();
    IPicture* pPicture = NULL;
    HRESULT hr = OleLoadPicturePath(pathname, NULL, 0, 0, &IID_IPicture, (void**)&pPicture);
    if (FAILED(hr) || !pPicture) return NULL;
    return (void*)pPicture;
}

// 从活着的 IPicture 里取图形句柄 (不销毁 picture 本身)。
// *kind 出 1=BITMAP 2=METAFILE 3=ICON (IPicture::Type 的 PICTTYPE 取值), 可不传 NULL。
// 注意 IPicture::Handle 在对象存活期才有效, 别把句柄单独存起来脱离 picture。
void* vb6_PictureHandleOf(void* picture, int32_t* kind) {
    if (!picture) return NULL;
    IPicture* p = (IPicture*)picture;
    OLE_HANDLE h = 0;
    if (p->lpVtbl->get_Handle(p, &h) != S_OK || !h) return NULL;
    if (kind) {
        short t = 0;
        if (p->lpVtbl->get_Type(p, &t) == S_OK) *kind = (int32_t)t;
        else *kind = 1;
    }
    return (void*)(intptr_t)h;
}

// 释放 LoadPicture 返回的 picture 对象 (Release)。
void vb6_ReleasePicture(void* picture) {
    if (!picture) return;
    ((IPicture*)picture)->lpVtbl->Release((IPicture*)picture);
}

// ============================================================
// COM调用结果→vb6_VARIANT转换
// 将后期绑定COM调用的VARIANT*结果转为vb6_VARIANT并释放原指针
// 用于: v = dic.Item("hello") 等场景
// 此函数消耗VARIANT*所有权, 调用后原指针不可再使用
// ============================================================
vb6_VARIANT vb6_VariantFromComResult(void* variant_ptr) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    if (!variant_ptr) { result.vt = vb6_vtEmpty; return result; }
    VARIANT* pv = (VARIANT*)variant_ptr;
    // Windows VARENUM 与 vb6_vartype 值一致, 可直接赋值
    result.vt = (vb6_vartype)pv->vt;
    switch (pv->vt) {
        case VT_EMPTY: break;
        case VT_NULL:  break;
        case VT_I2:    result.iVal = pv->iVal; break;
        case VT_I4:    result.lVal = pv->lVal; break;
        case VT_R4:    result.fltVal = pv->fltVal; break;
        case VT_R8:    result.dblVal = pv->dblVal; break;
        case VT_CY:    result.cyVal = *(int64_t*)&pv->cyVal; break;
        case VT_DATE:  result.dblVal = pv->date; break;
        case VT_BSTR:
            result.bstrVal = pv->bstrVal;
            pv->bstrVal = NULL;  // 转移所有权, 防止VariantClear释放
            break;
        case VT_DISPATCH:
            result.pdispVal = pv->pdispVal;
            pv->pdispVal = NULL;  // 转移所有权
            break;
        case VT_BOOL:   result.boolVal = pv->boolVal; break;
        case VT_UI1:    result.bVal = pv->bVal; break;
        case VT_ERROR:  result.lVal = pv->scode; break;
        case VT_DECIMAL:
#ifdef _WIN64
            memcpy(&result.decVal, &pv->decVal, sizeof(result.decVal));
#else
            /* Fix <vbeclipse> rev11: x86 的 OLE VARIANT union 只 8 字节, DECIMAL 在
             * 外部由 pdecVal 指向 (vb6_VARIANT 现同布局); 深拷贝到自持缓冲,
             * 由 vb6_VariantClear 释放。 */
            if (pv->pdecVal) {
                result.pdecVal = (DECIMAL*)CoTaskMemAlloc(sizeof(DECIMAL));
                if (result.pdecVal) memcpy(result.pdecVal, pv->pdecVal, sizeof(DECIMAL));
            }
#endif
            break;
        default:
            // P24-02: VT_ARRAY - 将Windows SAFEARRAY转换为vb6_SafeArray1D
        case VT_ARRAY|VT_BSTR:
        case VT_ARRAY|VT_VARIANT:
        case VT_ARRAY|VT_I4:
        case VT_ARRAY|VT_I2:
        case VT_ARRAY|VT_R8:
        case VT_ARRAY|VT_R4: {
            SAFEARRAY* psa = pv->parray;
            if (psa && SafeArrayGetDim(psa) == 1) {
                LONG lBound = 0, uBound = 0;
                SafeArrayGetLBound(psa, 1, &lBound);
                SafeArrayGetUBound(psa, 1, &uBound);
                int32_t count = uBound - lBound + 1;
                VARTYPE baseVt = pv->vt & VT_TYPEMASK;
                vb6_safearray_elemtype et;
                int32_t esz;
                switch (baseVt) {
                    case VT_BSTR:    et = vb6_sa_bstr;     esz = sizeof(BSTR); break;
                    case VT_VARIANT: et = vb6_sa_variant;  esz = sizeof(VARIANT); break;
                    case VT_I4:      et = vb6_sa_long;     esz = sizeof(int32_t); break;
                    case VT_I2:      et = vb6_sa_int;      esz = sizeof(int16_t); break;
                    case VT_R8:      et = vb6_sa_double;   esz = sizeof(double); break;
                    case VT_R4:      et = vb6_sa_single;   esz = sizeof(float); break;
                    default:         et = vb6_sa_variant;  esz = sizeof(VARIANT); break;
                }
                vb6_SafeArray1D* arr = (vb6_SafeArray1D*)calloc(1, sizeof(vb6_SafeArray1D));
                arr->signature = 0x5A1D;  /* Fix 082g */
                arr->elemType = et;
                arr->elemSize = esz;
                arr->lBound = lBound;
                arr->uBound = uBound;
                arr->count = count;
                arr->data = calloc(count, esz);
                arr->isDynamic = 0;
                // Copy elements from Windows SAFEARRAY
                char* srcData = (char*)psa->pvData;
                if (baseVt == VT_VARIANT) {
                    // VARIANT elements: convert each to vb6_VARIANT
                    for (int32_t i = 0; i < count; i++) {
                        VARIANT* srcElem = (VARIANT*)(srcData + i * sizeof(VARIANT));
                        vb6_VARIANT* dstElem = (vb6_VARIANT*)((char*)arr->data + i * sizeof(vb6_VARIANT));
                        // Use recursive call for each element
                        VARIANT* copy = (VARIANT*)malloc(sizeof(VARIANT));
                        VariantInit(copy);
                        VariantCopy(copy, srcElem);
                        *dstElem = vb6_VariantFromComResult(copy);
                    }
                } else if (baseVt == VT_BSTR) {
                    // BSTR elements: copy pointer (transfer ownership)
                    for (int32_t i = 0; i < count; i++) {
                        BSTR srcBstr = ((BSTR*)srcData)[i];
                        BSTR* dstBstr = &((BSTR*)arr->data)[i];
                        *dstBstr = srcBstr;
                        // Clear source so VariantClear doesn't free it
                        ((BSTR*)srcData)[i] = NULL;
                    }
                } else {
                    // Numeric types: direct memcpy
                    memcpy(arr->data, srcData, (size_t)count * esz);
                }
                result.vt = (vb6_vartype)(vb6_vtArray | (baseVt == VT_BSTR ? vb6_vtBSTR : baseVt == VT_VARIANT ? vb6_vtVariant : vb6_vtLong));
                result.parray = arr;
            }
            // Null out the SAFEARRAY pointer before VariantClear
            pv->parray = NULL;
            break;
        }
// 未知类型: 尝试转换为BSTR
            {
                VARIANT vBSTR;
                VariantInit(&vBSTR);
                if (SUCCEEDED(VariantChangeType(&vBSTR, pv, 0, VT_BSTR))) {
                    result.vt = vb6_vtBSTR;
                    result.bstrVal = vBSTR.bstrVal;
                    vBSTR.bstrVal = NULL;
                } else {
                    result.vt = vb6_vtEmpty;
                }
                VariantClear(&vBSTR);
            }
            break;
    }
    VariantClear(pv);  // 清理原始VARIANT (BSTR/pdispVal已转移)
    free(pv);
    return result;
}
// P24-03: 栈上VARIANT→vb6_VARIANT转换 (不释放源VARIANT)
// 用于For Each等场景, 源VARIANT在栈上而非堆上
vb6_VARIANT vb6_VariantFromStackVARIANT(VARIANT* pv) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    if (!pv) { result.vt = vb6_vtEmpty; return result; }
    result.vt = (vb6_vartype)pv->vt;
    switch (pv->vt) {
        case VT_EMPTY: break;
        case VT_NULL:  break;
        case VT_I2:    result.iVal = pv->iVal; break;
        case VT_I4:    result.lVal = pv->lVal; break;
        case VT_R4:    result.fltVal = pv->fltVal; break;
        case VT_R8:    result.dblVal = pv->dblVal; break;
        case VT_CY:    result.cyVal = *(int64_t*)&pv->cyVal; break;
        case VT_DATE:  result.dblVal = pv->date; break;
        case VT_BSTR:
            result.bstrVal = SysAllocString(pv->bstrVal);  /* copy (不转移所有权) */
            break;
        case VT_DISPATCH:
            result.pdispVal = pv->pdispVal;
            if (pv->pdispVal) pv->pdispVal->lpVtbl->AddRef(pv->pdispVal);
            break;
        case VT_BOOL:   result.boolVal = pv->boolVal; break;
        case VT_UI1:    result.bVal = pv->bVal; break;
        case VT_ERROR:  result.lVal = pv->scode; break;
        case VT_DECIMAL:
#ifdef _WIN64
            memcpy(&result.decVal, &pv->decVal, sizeof(result.decVal));
#else
            /* Fix <vbeclipse> rev11: x86 的 OLE VARIANT union 只 8 字节, DECIMAL 在
             * 外部由 pdecVal 指向 (vb6_VARIANT 现同布局); 深拷贝到自持缓冲,
             * 由 vb6_VariantClear 释放。 */
            if (pv->pdecVal) {
                result.pdecVal = (DECIMAL*)CoTaskMemAlloc(sizeof(DECIMAL));
                if (result.pdecVal) memcpy(result.pdecVal, pv->pdecVal, sizeof(DECIMAL));
            }
#endif
            break;
        case VT_ARRAY|VT_BSTR:
        case VT_ARRAY|VT_VARIANT:
        case VT_ARRAY|VT_I4:
        case VT_ARRAY|VT_I2:
        case VT_ARRAY|VT_R8:
        case VT_ARRAY|VT_R4: {
            SAFEARRAY* psa = pv->parray;
            if (psa && SafeArrayGetDim(psa) == 1) {
                LONG lBound = 0, uBound = 0;
                SafeArrayGetLBound(psa, 1, &lBound);
                SafeArrayGetUBound(psa, 1, &uBound);
                int32_t count = uBound - lBound + 1;
                VARTYPE baseVt = pv->vt & VT_TYPEMASK;
                vb6_safearray_elemtype et;
                int32_t esz;
                switch (baseVt) {
                    case VT_BSTR:    et = vb6_sa_bstr;     esz = sizeof(BSTR); break;
                    case VT_VARIANT: et = vb6_sa_variant;  esz = sizeof(VARIANT); break;
                    case VT_I4:      et = vb6_sa_long;     esz = sizeof(int32_t); break;
                    case VT_I2:      et = vb6_sa_int;      esz = sizeof(int16_t); break;
                    case VT_R8:      et = vb6_sa_double;   esz = sizeof(double); break;
                    case VT_R4:      et = vb6_sa_single;   esz = sizeof(float); break;
                    default:         et = vb6_sa_variant;  esz = sizeof(VARIANT); break;
                }
                vb6_SafeArray1D* arr = (vb6_SafeArray1D*)calloc(1, sizeof(vb6_SafeArray1D));
                arr->signature = 0x5A1D;  /* Fix 082g */
                arr->elemType = et;
                arr->elemSize = esz;
                arr->lBound = lBound;
                arr->uBound = uBound;
                arr->count = count;
                arr->data = calloc(count, esz);
                arr->isDynamic = 0;
                char* srcData = (char*)psa->pvData;
                if (baseVt == VT_BSTR) {
                    for (int32_t i = 0; i < count; i++) {
                        ((BSTR*)arr->data)[i] = SysAllocString(((BSTR*)srcData)[i]);
                    }
                } else if (baseVt == VT_VARIANT) {
                    for (int32_t i = 0; i < count; i++) {
                        VARIANT* srcElem = (VARIANT*)(srcData + i * sizeof(VARIANT));
                        vb6_VARIANT* dstElem = (vb6_VARIANT*)((char*)arr->data + i * sizeof(vb6_VARIANT));
                        *dstElem = vb6_VariantFromStackVARIANT(srcElem);  /* recursive */
                    }
                } else {
                    memcpy(arr->data, srcData, (size_t)count * esz);
                }
                result.vt = (vb6_vartype)(vb6_vtArray | (baseVt == VT_BSTR ? vb6_vtBSTR : baseVt == VT_VARIANT ? vb6_vtVariant : vb6_vtLong));
                result.parray = arr;
            }
            break;
        }
        default:
            {
                VARIANT vBSTR;
                VariantInit(&vBSTR);
                if (SUCCEEDED(VariantChangeType(&vBSTR, pv, 0, VT_BSTR))) {
                    result.vt = vb6_vtBSTR;
                    result.bstrVal = vBSTR.bstrVal;
                    vBSTR.bstrVal = NULL;
                } else {
                    result.vt = vb6_vtEmpty;
                }
                VariantClear(&vBSTR);
            }
            break;
    }
    /* 注意: 不调用 free(pv), 因为源VARIANT在栈上 */
    return result;
}

/* P24-04: Extract IDispatch pointer from Variant for COM late-binding */
/* When vb6_VARIANT holds an object (vt==vb6_vtDispatch), return its pdispVal; */
/* otherwise return NULL. Used when Variant-typed vars access COM members. */
void* vb6_VariantToObject(vb6_VARIANT* v) {
    if (!v) return NULL;
    if (v->vt == vb6_vtDispatch) return v->pdispVal;
    return NULL;
}

/* P24-04: Pack vb6_VARIANT (by value) into Windows VARIANT for COM call args */
/* Used when Variant-typed variable member access result is passed as COM arg */
void* vb6_ComPackVariant(vb6_VARIANT v) {
    /* Fix 150: 必须用 malloc —— 所有消费方 (vb6_ComSetProp/ComSetPropArg/ComCall)
       都用 free() 释放打包 VARIANT。此前用 CoTaskMemAlloc 而 free() 释放,
       跨分配器释放损坏堆, 表现为 OLEAUT32.dll 内 0xC0000005 (NewTab Theme 赋值实测) */
    VARIANT* pv = (VARIANT*)malloc(sizeof(VARIANT));
    if (!pv) return NULL;
    VariantInit(pv);
    switch (v.vt) {
        case vb6_vtEmpty: pv->vt = VT_EMPTY; break;
        case vb6_vtNull:  pv->vt = VT_NULL; break;
        case vb6_vtInteger: pv->vt = VT_I2; pv->iVal = v.iVal; break;
        case vb6_vtLong:    pv->vt = VT_I4; pv->lVal = v.lVal; break;
        case vb6_vtSingle:  pv->vt = VT_R4; pv->fltVal = v.fltVal; break;
        case vb6_vtDouble:  pv->vt = VT_R8; pv->dblVal = v.dblVal; break;
        case vb6_vtBSTR:    pv->vt = VT_BSTR; pv->bstrVal = SysAllocString(v.bstrVal); break;
        case vb6_vtDispatch: pv->vt = VT_DISPATCH; pv->pdispVal = (IDispatch*)v.pdispVal; break;
        case vb6_vtBoolean: pv->vt = VT_BOOL; pv->boolVal = v.boolVal ? VARIANT_TRUE : VARIANT_FALSE; break;
        case vb6_vtByte:    pv->vt = VT_UI1; pv->bVal = v.bVal; break;
        default: pv->vt = VT_EMPTY; break;
    }
    return pv;
}


/* Fix 084f: Variant数组嵌套索引按值版本 (见 vb6rtl.h 说明) */
vb6_VARIANT vb6_VariantArrayGetVal(vb6_VARIANT v, int32_t index) {
    return vb6_VariantArrayGet(&v, index);
}

/* Fix 158f: Variant数组元素的取址 (见 vb6rtl_variant.h 说明) */
vb6_VARIANT* vb6_VariantArrayElemPtr(vb6_VARIANT* v, int32_t index) {
    if (!v || !(v->vt & vb6_vtArray) || !v->parray) return NULL;
    vb6_SafeArray1D* arr = v->parray;
    if (index < arr->lBound || index > arr->uBound) return NULL;
    if (arr->elemType != vb6_sa_variant) return NULL;
    return &((vb6_VARIANT*)arr->data)[index - arr->lBound];
}

/* Variant数组索引: 从持有SafeArray的Variant中取/设元素 */
vb6_VARIANT vb6_VariantArrayGet(vb6_VARIANT* v, int32_t index) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    result.vt = vb6_vtEmpty;
    if (!v || !(v->vt & vb6_vtArray) || !v->parray) return result;
    vb6_SafeArray1D* arr = v->parray;
    if (index < arr->lBound || index > arr->uBound) return result;
    int32_t off = index - arr->lBound;
    switch (arr->elemType) {
        case vb6_sa_long:
            result.vt = vb6_vtLong;
            result.lVal = ((int32_t*)arr->data)[off];
            break;
        case vb6_sa_int:
            result.vt = vb6_vtInteger;
            result.iVal = ((int16_t*)arr->data)[off];
            break;
        case vb6_sa_double:
            result.vt = vb6_vtDouble;
            result.dblVal = ((double*)arr->data)[off];
            break;
        case vb6_sa_single:
            result.vt = vb6_vtDouble; /* VB6 Single -> Double promotion */
            result.dblVal = (double)((float*)arr->data)[off];
            break;
        case vb6_sa_bstr:
            result.vt = vb6_vtBSTR;
            result.bstrVal = ((BSTR*)arr->data)[off];
            break;
        case vb6_sa_bool:
            result.vt = vb6_vtBoolean;
            result.boolVal = ((int16_t*)arr->data)[off];
            break;
        case vb6_sa_byte:
            result.vt = vb6_vtInteger; /* VB6 Byte -> Integer promotion */
            result.iVal = (int16_t)((uint8_t*)arr->data)[off];
            break;
        case vb6_sa_variant:
            result = ((vb6_VARIANT*)arr->data)[off];
            break;
        case vb6_sa_ptr:
            result.vt = vb6_vtLong; /* object pointer as Long */
            result.lVal = (int32_t)(intptr_t)((void**)arr->data)[off];
            break;
        case vb6_sa_currency:
            result.vt = vb6_vtLong;
            result.lVal = (int32_t)((int64_t*)arr->data)[off]; /* truncate to Long */
            break;
        default:
            result.vt = vb6_vtEmpty;
            break;
    }
    return result;
}

void vb6_VariantArraySet(vb6_VARIANT* v, int32_t index, vb6_VARIANT val) {
    if (!v || !(v->vt & vb6_vtArray) || !v->parray) return;
    vb6_SafeArray1D* arr = v->parray;
    if (index < arr->lBound || index > arr->uBound) return;
    int32_t off = index - arr->lBound;
    if (arr->elemType == vb6_sa_variant) {
        vb6_VARIANT* slot = &((vb6_VARIANT*)arr->data)[off];
        /* Fix <vbeclipse> rev16: 原先 `vb6_VariantClear(slot); *slot = val;` 是
           Clear + **浅拷贝** —— val 若是从别处借来的 Variant (ByRef 实参/数组元素),
           槽就与它共用同一只 BSTR, 宿主一 Clear 槽即悬垂。改走"接管"语义。 */
        vb6_VariantAssign(slot, val);
    }
    /* 对于非Variant数组, 赋值时需要按目标类型转换(简化: 仅Variant数组支持赋值) */
}

// Fix 048: LoadResData — stub (resource loading not supported in C3 runtime)
vb6_VARIANT vb6_LoadResData(int32_t resourceId, int32_t resourceType) {
    (void)resourceId; (void)resourceType;
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); return v;  /* empty Variant */
}

// Fix <vbeclipse>: LoadResPicture / LoadResString — stub (同 LoadResData 口径:
// 资源段加载暂不支持, 返回 empty Variant; 调用侧拿到 Nothing/空串不崩)
vb6_VARIANT vb6_LoadResPicture(int32_t resourceId, int32_t resourceType) {
    (void)resourceId; (void)resourceType;
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); return v;
}

vb6_VARIANT vb6_LoadResString(int32_t resourceId) {
    (void)resourceId;
    vb6_VARIANT v; memset(&v, 0, sizeof(v)); return v;
}

// ============================================================
// Fix 093a: P21-14 / P21-15 — SavePicture / Load 语句
// (vb6rtl.h 早已声明, 但 RTL 里一直没有实现 → LNK2019)
// ============================================================

// SavePicture Picture, "file" — 把 StdPicture (运行时为 IPicture*/IDispatch) 按
// 扩展名编码保存 (jpg/bmp/gif). 用 OLE 的 OleSavePictureFile: 它内部只通过 vtable
// 槽 0 (QueryInterface) 取 IPersistStream, IPicture/IDispatch 的该槽布局相同,
// 故直接强转 LPDISPATCH 安全 (该 API 不做接口专属的虚调用).
void vb6_SavePicture(void* hBitmap, BSTR filename) {
    if (!hBitmap || !filename) return;
#ifdef _WIN32
    OleSavePictureFile((LPDISPATCH)hBitmap, filename);
#else
    (void)hBitmap; (void)filename;
#endif
}

// Load Form — 预加载窗体 (不显示). 本运行时的窗体默认实例由
// vb6_form_show_<Form>() 首次调用时创建, 对象恒可用, 无独立预加载阶段.
void vb6_LoadForm(void* hwnd) {
    (void)hwnd;
}

// ============================================================
// Fix 105: UserControl/PropertyPage host built-in objects
// Generated C emits UserControl.* / Ambient.* / Extender.* / PropertyPage.* as
// bare global identifiers (e.g. vb6_UserControl_ScaleWidth); previously
// undefined -> C2065. Provide per-process single-instance host state here;
// multi-instance hosting is a later feature block.
// ============================================================
#include <windows.h>
#include "vb6rtl_userctl.h"
#include "vb6forms.h"

// --- Font object instance (ambient font; source for control Font when unset) ---
// czUI fix: 环境字体必须有有效 Name — Bag 重放路径 ReadProperty("Font",
// Ambient.Font) 会把此对象设为控件字体; Name=NULL 时 GdipCreateFontFamily
// 失败 → 所有 GDI+ 文字静默消失 (Charts2020 图表标题/百分比实测)
vb6_ComIface_Font g_vb6_UserControl_FontObj = { NULL, 8.25f, 0, 0, 0, 0, 400, 0 };

// --- UserControl host state ---
int32_t vb6_UserControl_ScaleWidth  = 0;
int32_t vb6_UserControl_ScaleHeight = 0;
int32_t vb6_UserControl_ScaleMode   = 1;     // Twip (VB6 default)
void*   vb6_UserControl_hDC          = NULL;
int32_t vb6_UserControl_ContainerHwnd = 0;
int16_t vb6_UserControl_Enabled     = -1;
int32_t vb6_UserControl_MousePointer = 0;
void*   vb6_UserControl_MouseIcon   = NULL;
int32_t vb6_UserControl_OLEDropMode = 0;
// Fix 154-B: UserControl 其余宿主属性定义 (设计期/宿主状态, 初值随意).
void*   vb6_UserControl_Picture = NULL;
int32_t vb6_UserControl_BackColor = 0;
int32_t vb6_UserControl_ForeColor = 0;
int16_t vb6_UserControl_RightToLeft = 0;
void*   vb6_UserControl_ParentControls = NULL;
void*   vb6_UserControl_Controls = NULL;
void*   vb6_Screen_MouseIcon = NULL;       // Fix <vbeclipse>: Screen.MouseIcon 槽   // Fix <vbeclipse>: UserControl.Controls (集合未建模, 恒 NULL)
vb6_ComIface_Font* vb6_UserControl_Font = &g_vb6_UserControl_FontObj;
struct vb6_UserControl_Ambient_Type vb6_UserControl_Ambient = { &g_vb6_UserControl_FontObj };

// --- Ambient host environment ---
vb6_ComIface_Font* vb6_Ambient_Font = &g_vb6_UserControl_FontObj;
int16_t vb6_Ambient_UserMode   = -1;         // compiled output is runtime
BSTR    vb6_Ambient_DisplayName = NULL;     // set to control instance name at runtime
int32_t vb6_Ambient_ForeColor  = 0;          // black
int32_t vb6_Ambient_BackColor  = 0x8000000F; // BTNFACE (VB6 default)
int16_t vb6_Ambient_RightToLeft = 0;         // Fix 154-B: TriState 从环境属性

// --- Extender ---
int32_t vb6_Extender_Left = 0;
int32_t vb6_Extender_Top  = 0;
// Fix 153-B: Extender 属性转发槽位.
// Fix 162 更正: 下面这句"设计期对象, 运行期恒不活动 → 初值随意"是**错的** ——
// Width/Height 在运行期被真实读取 (VBFlexGrid.ctl 28728 用它做 MouseUp 命中测试:
// `If (X >= 0 And X <= UserControl.Width) And ... Then RaiseEvent Click`)。
// 当时两者只有定义、无任何赋值 → 恒为 0 → 命中测试恒假 → Click 事件永不触发,
// 且 UserControl.Width/Height 连声明都没有 → C2065。
// 现由 vb6_uc_push/vb6_uc_pop 按当前实例同步 (源: vb6_UCRec.extWidth/extHeight,
// 创建入口收到的缇值, 精确非近似)。
int32_t vb6_Extender_Width  = 0;
int32_t vb6_Extender_Height = 0;
// Fix 162: `UserControl.Width/Height` 与 Extender 同值 (cgen 对宿主伪对象
// UserControl 发 vb6_UserControl_<Member>, 对 Extender 发 vb6_Extender_<Member>)。
int32_t vb6_UserControl_Width  = 0;
int32_t vb6_UserControl_Height = 0;
int16_t vb6_Extender_Visible = -1;
BSTR    vb6_Extender_Tag = NULL;
BSTR    vb6_Extender_ToolTipText = NULL;
int32_t vb6_Extender_DragMode = 0;
int32_t vb6_Extender_Align = 0;
int32_t vb6_Extender_HelpContextID = 0;
int32_t vb6_Extender_WhatsThisHelpID = 0;
void*   vb6_Extender_Container = NULL;
void*   vb6_Extender_DragIcon = NULL;

// --- Fix 133u: Extender.Visible/Height (czUI.ctl) ---
struct vb6_UserControl_Extender_Type vb6_UserControl_Extender = { -1, 0 };

// --- Fix 133u: UserControl.hWnd / AutoRedraw (czUI.ctl) ---
void*   vb6_UserControl_hWnd = NULL;
int16_t vb6_UserControl_AutoRedraw = 1;

// --- PropertyPage ---
void*   vb6_PropertyPage_hwnd        = NULL;
void*   vb6_PropertyPage_hWnd        = NULL;  // both spellings denote the same concept
int32_t vb6_PropertyPage_ScaleMode   = 1;
int32_t vb6_PropertyPage_ScaleHeight = 0;
int16_t vb6_PropertyPage_Changed     = 0;

// --- PropertyPage built-in Changed property (Fix 108d) ---
int16_t Changed = 0;

// --- Picture.Line mode constants ---
const int32_t B  = 1;
const int32_t BF = 2;

// VB6 UserControl.TextWidth/TextHeight: measure with GDI using current Font
// (unit = ScaleMode; simplified to pixels here; Twip handled once hosting lands)
//
// Fix 113k 已回退: 曾尝试按 vb6_UserControl_Font 建 HFONT 选入 DC 再测, 更"正确",
// 但实测使 Charts 2020 整体布局**更差** —— 本实现里 ScaleWidth/ScaleHeight 是
// 像素, 而 .ctl 的布局常数(PT16=(SW+SH)*2.5/100 等)是按 VB6 的 TWIP 语义推的,
// 两者本就不同源; 换成更大的字体度量后饼图被图例挤没、柱图 X 轴标签裁切更重
// (见离屏 dump ucPieChart/ucChartBar). 故保留原"默认 DC 字体测量"行为.
static int32_t vb6_uc_measureText(BSTR text, int wantWidth) {
    if (!text) return 0;
    HDC hdc = GetDC(NULL);
    if (!hdc) return 0;
    SIZE sz = { 0, 0 };
    int len = SysStringLen(text);

    // Fix 129: VB6 的 UserControl.TextWidth/TextHeight 用**控件自身的 Font** 度量。
    // 此前直接用 GetDC(NULL) 的默认字体 → 控件内部的文字/标题/图例排版全部算错:
    // 标题预留高度(TopHeader)不足 → 图例压在标题上; 轴标签宽度不对 → 文字裁切。
    HFONT hOld = NULL, hNew = NULL;
    if (len) {
        vb6_ComIface_Font* f = (vb6_ComIface_Font*)vb6_UserControl_Font;
        LOGFONTW lf;
        memset(&lf, 0, sizeof(lf));
        double pt = (f && f->Size > 0) ? (double)f->Size : 8.25;
        lf.lfHeight = -(int)(pt * 96.0 / 72.0 + 0.5);          // 磅 → 像素 @96dpi
        lf.lfWeight = (f && f->Weight) ? f->Weight : ((f && f->Bold) ? 700 : 400);
        lf.lfItalic = (BYTE)((f && f->Italic) ? 1 : 0);
        lf.lfUnderline = (BYTE)((f && f->Underline) ? 1 : 0);
        lf.lfStrikeOut = (BYTE)((f && f->Strikethrough) ? 1 : 0);
        lf.lfCharSet = (BYTE)((f && f->Charset) ? f->Charset : DEFAULT_CHARSET);
        if (f && f->Name) {
            size_t n = wcslen(f->Name);
            if (n > 31) n = 31;
            memcpy(lf.lfFaceName, f->Name, n * sizeof(wchar_t));
        } else {
            wcscpy(lf.lfFaceName, L"MS Sans Serif");
        }
        hNew = CreateFontIndirectW(&lf);
        if (hNew) hOld = (HFONT)SelectObject(hdc, hNew);
    }
    if (len) GetTextExtentPoint32W(hdc, text, len, &sz);
    if (hNew) { SelectObject(hdc, hOld); DeleteObject(hNew); }
    ReleaseDC(NULL, hdc);
    return wantWidth ? sz.cx : sz.cy;
}

int32_t vb6_UserControl_TextWidth(BSTR text) {
    return vb6_uc_measureText(text, 1);
}

int32_t vb6_UserControl_TextHeight(BSTR text) {
    return vb6_uc_measureText(text, 0);
}

// UserControl.Size: VB6 `UserControl.Size width, height` (unit = ScaleMode).
// Windowless controls take their size from the container; update host state and
// refresh the container so ScaleWidth/ScaleHeight stay self-consistent.
void vb6_UserControl_Size(double width, double height) {
    if (!(width != width))  vb6_UserControl_ScaleWidth  = (int32_t)width;
    if (!(height != height)) vb6_UserControl_ScaleHeight = (int32_t)height;
    vb6_UserControl_Refresh();
}

// UserControl.Refresh: windowless controls have no own window; refresh the
// most recently entered UserControl host window (see Fix 112).
void vb6_UserControl_Refresh(void) {
    vb6_UC_RefreshCurrent();
}

// UserControl.CancelAsyncRead: async read unimplemented, no-op.
void vb6_UserControl_CancelAsyncRead(BSTR propName) {
    (void)propName;
}

// Fix <vbeclipse>: `UserControl.Line (x1,y1)-(x2,y2), [color], [mode]`
// (ucTab.ctl:231-241 画标签边框/渐变分隔线; Shape 控件的 Line 语法).
// 生成端把 `-` 连写的坐标对拍平成 5 个固定实参 (x1,y1,x2,y2,color), `, B` /
// `, BF` 模式常量按 Fix 102 的口径原样作**可变参**追加 —— 故此处必须变参, 否则
// 5 参形态 (源码 `UserControl.Line (1, h-10)-(w, h-10), m_Scheme.FrameColor`)
// 无原型匹配 → LNK2001.
// 可变参里第 6 个才是模式, 但模式类型(int32_t)与 color 相同, 逐个 va_arg 会串味;
// 这里改为**不读可变参**: 图形最终态由下一轮 WM_PAINT 重绘路径决定, 保持 no-op
// 语义安全 (不依赖 mode). 坐标/颜色实参签名化, 便于将来接 GDI 画线时不必改生成端.
void vb6_UserControl_Line(double x1, double y1, double x2, double y2, int32_t color, ...) {
    (void)x1; (void)y1; (void)x2; (void)y2; (void)color;
}

// Fix 111: UserControl built-in methods (declared in vb6rtl_userctl.h).
//
// ScaleX/ScaleY: convert x from fromScale to toScale (VB6 ScaleMode constants).
// 96dpi baseline: Twip = 1/15 px, Point = 96/72 px, Inch = 96 px ...
// User(0)/ContainerPosition(8)/ContainerSize(9,10)/unknown are treated as
// pixels -- matching Charts 2020 usage (Extender.Left is already container
// pixels; target UserControl.ScaleMode = 3 = Pixel -> identity).
static double vb6_ucScaleToPixels(int32_t mode) {
    // Fix 184: 单位表必须用**真实 DPI**。此前整张表按 96 写死，而容器侧
    // (vb6_TwipToX / vb6_XToTwipX) 已按 DPI，于是 UserControl 内部每做一次
    // 缇<->像素往返就缩 20% (VBFlexGrid 内层窗口 914px -> 731px)。
    static double s_dpi = 0.0;
    double dpi;
    if (s_dpi <= 0.0) {
        HDC dc = GetDC(NULL);
        int d = dc ? GetDeviceCaps(dc, LOGPIXELSX) : 96;
        if (dc) ReleaseDC(NULL, dc);
        s_dpi = (d > 0) ? (double)d : 96.0;
    }
    dpi = s_dpi;
    switch (mode) {
        case 1: return dpi / 1440.0;    /* Twips */
        case 2: return dpi / 72.0;      /* Points */
        case 3: return 1.0;             /* Pixels */
        case 4: return 1.0;             /* Characters (approx) */
        case 5: return dpi;             /* Inches */
        case 6: return dpi / 25.4;      /* Millimeters */
        case 7: return dpi / 2.54;      /* Centimeters */
        default: return 1.0;            /* User / Container* / unknown */
    }
}

double vb6_UserControl_ScaleX(double x, int32_t fromScale, int32_t toScale) {
    double px = x * vb6_ucScaleToPixels(fromScale);
    double f = vb6_ucScaleToPixels(toScale);
    return (f == 0.0) ? x : (px / f);
}

double vb6_UserControl_ScaleY(double y, int32_t fromScale, int32_t toScale) {
    double px = y * vb6_ucScaleToPixels(fromScale);
    double f = vb6_ucScaleToPixels(toScale);
    return (f == 0.0) ? y : (px / f);
}

// UserControl.AsyncRead: no container/async message pump in compiled form, so
// real async reads are unavailable; record as no-op (symmetric with
// CancelAsyncRead). Callers depending on UserControl_AsyncReadComplete simply
// skip the async image load; the rest of the control is unaffected.
void vb6_UserControl_AsyncRead(BSTR url, int32_t asyncType, BSTR propertyName,
                               int32_t flags) {
    (void)url; (void)asyncType; (void)propertyName; (void)flags;
}

// UserControl.PropertyChanged: notify container "property changed". No container
// callback registry in compiled form; no-op, consistent with the built-in
// Changed (vb6_PropertyPage_Changed / bare Changed).
void vb6_UserControl_PropertyChanged(BSTR propName) {
    (void)propName;
}

// Fix 174: Extender / UserControl host methods referenced by VBFlexGrid.ctl
// (Standard EXE 编译形态下容器不活动, 方法做最小可行实现即可满足链接/运行):
//   UserControl.OLEDrag → vb6_UserControl_OLEDrag      (VBFlexGrid.ctl 2832)
//   Extender.Drag       → vb6_Extender_Drag            (VBFlexGrid.ctl 3047)
//   Extender.SetFocus   → vb6_Extender_SetFocus        (VBFlexGrid.ctl 3052)
//   Extender.ZOrder     → vb6_Extender_ZOrder          (VBFlexGrid.ctl 3057)
// OLE drag/drop 在无容器运行时无意义; Drag 与 ZOrder 的可选参数保留签名
// (cgen 以 `vb6_Extender_Drag(&(vb6_VARIANT){0}, 0)` 形态调用), 忽略即可.
void vb6_UserControl_OLEDrag(void) {
    (void)0;
}

void vb6_Extender_Drag(vb6_VARIANT* Action, int _has_Action) {
    (void)Action; (void)_has_Action;
    (void)0;
}

void vb6_Extender_SetFocus(void) {
    if (vb6_UserControl_hWnd) SetFocus((HWND)vb6_UserControl_hWnd);
}

void vb6_Extender_ZOrder(vb6_VARIANT* Position, int _has_Position) {
    (void)Position; (void)_has_Position;
    (void)0;
}

// ============================================================
// tB Interface (ai/022 B06a): 薄指针 IUnknown 前缀槽的运行时支撑
// ============================================================

int32_t vb6_IidEqual(const void* a, const void* b) {
    const unsigned char* x = (const unsigned char*)a;
    const unsigned char* y = (const unsigned char*)b;
    int i;
    if (!x || !y) return 0;
    for (i = 0; i < 16; i++) { if (x[i] != y[i]) return 0; }
    return 1;
}

int32_t vb6_IfaceSupports(void* ifacePtr, const void* iid) {
    vb6_ivtbl_prefix* vt;
    void* got = NULL;
    if (!ifacePtr || !iid) return 0;
    vt = *(vb6_ivtbl_prefix**)ifacePtr;      /* 薄指针首字 = 该接口的 vtable */
    if (!vt || !vt->QueryInterface) return 0;
    if (vt->QueryInterface(ifacePtr, iid, &got) != 0 || !got) return 0;
    if (*(vb6_ivtbl_prefix**)got) (*(vb6_ivtbl_prefix**)got)->Release(got);
    return 1;
}
