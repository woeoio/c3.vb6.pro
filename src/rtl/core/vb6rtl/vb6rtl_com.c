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
#include <wincodec.h>   /* Fix <vbeclipse> 2026-10-06: WIC — WebP/SVG 等扩展解码 */
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
            // Fix <vbeclipse>: 用受守卫的 vb6_ComAddRefDispatch (内含 vb6_ComIsDispatchable),
            // 与 vb6_VariantObject 同口径。`For Each ctrl In Controls` 交回的是**裸子控件
            // HWND** (宿主分派层的对象表示), 不是真 IDispatch —— 裸 AddRef 会把 HWND 首字段
            // 当 vtable 解引用 → av read (ucTabStrip.UserControl_Resize 实测)。真 COM 接收者
            // 仍照常 AddRef; 析构侧 vb6_ReleaseObject 同判据跳过, 收支平衡。
            if (pv->pdispVal) vb6_ComAddRefDispatch(pv->pdispVal);
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

// ============================================================
// Fix <vbeclipse> (2026-10-06): LoadRes* 实装 — .res 资源段加载。
// 此前三个都是空桩返 empty Variant; 实际上用户 .res (VBP 的 ResFile32) 早在
// P23-03 就随链接进了 exe (driver_link.cpp "Pass user .res file to linker"),
// 运行期 FindResource/LoadResource 直读即可, 一直缺的只是这一层。
// 形参口径: VB6 里三个函数的实参本就是 Variant —— LoadResString(101) 数字 id、
// LoadResData("BIN1","CUSTOM") 字符串名都合法, 故形参统一 vb6_VARIANT, 调用侧
// 由 cgen 的 vb6_VariantFromValue 包装。资源找不到按 VB6 抛错误 326。
// ============================================================

// 实参解包: VT_BSTR → 资源名 (按名查找); 其余数值 → MAKEINTRESOURCEW。
static const wchar_t* vb6_ResNameOf(const vb6_VARIANT* v) {
    if (v && (vb6_vartype)v->vt == VT_BSTR && v->bstrVal) return v->bstrVal;
    return NULL;
}

// 实参解包: 数值档 (I2/I4/UI1/BYTE); BSTR 走 _wtoi 兜底 ("101" 形的字符串 id)。
static int32_t vb6_ResIdOf(const vb6_VARIANT* v) {
    if (!v) return 0;
    switch ((vb6_vartype)v->vt) {
        case (vb6_vartype)VT_I2:  return v->iVal;
        case (vb6_vartype)VT_I4:  return v->lVal;
        case (vb6_vartype)VT_INT: return v->lVal;
        case (vb6_vartype)VT_UI1: return v->bVal;
        case (vb6_vartype)VT_BSTR: return v->bstrVal ? _wtoi(v->bstrVal) : 0;
        default: return (int32_t)v->lVal;
    }
}

// 找到并锁定资源; 返回数据指针, *outSize 收字节数。找不到返回 NULL (调用侧抛 326)。
static const void* vb6_ResLoad(const vb6_VARIANT* id, const wchar_t* typeName,
                               int32_t typeId, uint32_t* outSize) {
    *outSize = 0;
    const wchar_t* name = vb6_ResNameOf(id);
    LPCWSTR rName = name ? name : MAKEINTRESOURCEW(vb6_ResIdOf(id));
    LPCWSTR rType = typeName ? typeName : MAKEINTRESOURCEW(typeId);
    HRSRC hr = FindResourceW(NULL, rName, rType);
    if (!hr) return NULL;
    HGLOBAL h = LoadResource(NULL, hr);
    if (!h) return NULL;
    const void* p = LockResource(h);
    if (!p) return NULL;
    *outSize = SizeofResource(NULL, hr);
    return p;
}

// rc.exe 会把 .ico 展开成 RT_GROUP_ICON(用户写的那个 id) + 若干重编号的 RT_ICON;
// LoadResPicture 的 id 指的是**组**。读组里首枚 entry 的 nID, 再取对应裸图。
// (RT_GROUP_CURSOR/RT_CURSOR 同构。)
static const void* vb6_ResResolveGroup(const vb6_VARIANT* id, int isCursor, uint32_t* outSize) {
    *outSize = 0;
    const wchar_t* name = vb6_ResNameOf(id);
    LPCWSTR rName = name ? name : MAKEINTRESOURCEW(vb6_ResIdOf(id));
    LPCWSTR grpType = isCursor ? RT_GROUP_CURSOR : RT_GROUP_ICON;
    LPCWSTR imgType = isCursor ? RT_CURSOR : RT_ICON;
    HRSRC hr = FindResourceW(NULL, rName, grpType);
    if (!hr) return NULL;
    HGLOBAL h = LoadResource(NULL, hr);
    const uint8_t* g = h ? (const uint8_t*)LockResource(h) : NULL;
    if (!g) return NULL;
    int32_t count = g[4] | (g[5] << 8);          /* idCount */
    if (count < 1) return NULL;
    const uint8_t* e = g + 6;                     /* 首枚 GRPICONDIRENTRY, 末 2 字节 nID */
    int32_t nID = e[12] | (e[13] << 8);
    HRSRC hr2 = FindResourceW(NULL, MAKEINTRESOURCEW(nID), imgType);
    if (!hr2) return NULL;
    HGLOBAL h2 = LoadResource(NULL, hr2);
    const void* p = h2 ? LockResource(h2) : NULL;
    if (!p) return NULL;
    *outSize = SizeofResource(NULL, hr2);
    return p;
}

vb6_VARIANT vb6_LoadResString(vb6_VARIANT resourceId) {
    vb6_VARIANT ret; memset(&ret, 0, sizeof(ret));
    // Win32 字符串表: 16 条一档 — 资源块 id = id/16 + 1, 档内下标 = id%16
    // (101 → block 7 slot 5, 2026-10-06 probe 实测)。
    // 条目格式 = **WORD 长度前缀 + 该长度的字符** (不是零结尾!) — 本机 rc.exe 产物
    // 逐字节实测: block 7 = 5×[0000](空条目) + [0C 00]"ResString-OK" + 10×[0000],
    // 长度恰好 56 字节。旧的 wcslen 游走把长度词当首字符, 读出 CHR$(12)&s。
    int32_t id = vb6_ResIdOf(&resourceId);
    if (id >= 1) {
        int32_t block = id / 16 + 1;
        int32_t slot = id % 16;
        HRSRC hr = FindResourceW(NULL, MAKEINTRESOURCEW(block), RT_STRING);
        if (hr) {
            HGLOBAL h = LoadResource(NULL, hr);
            const wchar_t* tab = h ? (const wchar_t*)LockResource(h) : NULL;
            if (tab) {
                for (int32_t i = 0; i < slot; i++) tab += 1 + (int32_t)*tab;
                int32_t len = (int32_t)*tab;
                if (len > 0) {
                    wchar_t* buf = (wchar_t*)malloc(((size_t)len + 1) * sizeof(wchar_t));
                    if (buf) {
                        memcpy(buf, tab + 1, (size_t)len * sizeof(wchar_t));
                        buf[len] = L'\0';
                        BSTR b = vb6_BSTR_FromStr(buf);
                        free(buf);
                        return vb6_VariantString(b);
                    }
                }
                return vb6_VariantString(vb6_BSTR_Empty());
            }
        }
    }
    vb6_ErrRaiseNumber(326);   /* Resource with identifier not found */
    return ret;
}

// 常见图片文件的字节签名 (vb6_ResSigKind 用)
static int vb6_ResSigIsOleStreamable(const void* data, uint32_t size) {
    const uint8_t* d = (const uint8_t*)data;
    if (size < 8 || !d) return 0;
    if (d[0] == 0x42 && d[1] == 0x4D) return 1;                          /* 'BM' */
    if (d[0] == 0xFF && d[1] == 0xD8 && d[2] == 0xFF) return 1;          /* JPEG */
    if (d[0] == 0x89 && d[1] == 0x50 && d[2] == 0x4E && d[3] == 0x47) return 1;  /* PNG */
    if (d[0] == 0x47 && d[1] == 0x49 && d[2] == 0x46 && d[3] == 0x38) return 1;  /* GIF8 */
    if (d[0] == 0x00 && d[1] == 0x00 && d[2] == 0x01 && d[3] == 0x00) return 1;  /* ICO */
    return 0;
}

// OleLoadPicture 流路径 — PNG/JPEG/GIF/BMP/ICO 全套 (OLE 内建 GDI+ 解码)。
static IPicture* vb6_ResPicViaOle(const void* data, uint32_t size) {
    IPicture* pic = NULL;
    HGLOBAL hg = GlobalAlloc(GMEM_MOVEABLE, size);
    if (!hg) return NULL;
    void* pv = GlobalLock(hg);
    if (!pv) { GlobalFree(hg); return NULL; }
    memcpy(pv, data, size);
    GlobalUnlock(hg);
    IStream* st = NULL;
    if (SUCCEEDED(CreateStreamOnHGlobal(hg, TRUE, &st)) && st) {
        OleLoadPicture(st, 0, FALSE, &IID_IPicture, (void**)&pic);
        st->lpVtbl->Release(st);
    } else {
        GlobalFree(hg);
    }
    return pic;
}

// WIC 路径 — OleLoadPicture 不认的格式 (WebP; SVG 在装了 SVG 解码扩展的机器上)。
// CreateDecoderFromStream 是厂商无关入口, 系统装了什么 WIC 解码器就能吃什么。
static IPicture* vb6_ResPicViaWic(const void* data, uint32_t size) {
    static const GUID kWICFactory   = {0xcacaf262,0x9370,0x4615,{0xa1,0x3b,0x9f,0x55,0x39,0xda,0x4c,0x0a}};
    static const GUID kWICFactoryI  = {0xec5ec8a9,0xc395,0x4314,{0x9c,0x77,0x54,0xd7,0xa9,0x35,0xff,0x70}};
    static const GUID kPixFmtBGRA   = {0x6fddc324,0x4e03,0x4bfe,{0xb1,0x85,0x3d,0x77,0x76,0x8d,0xc9,0x10}};
    IPicture* pic = NULL;
    IWICImagingFactory* fac = NULL;
    IStream* st = NULL;
    HGLOBAL hg = NULL;
    IWICBitmapDecoder* dec = NULL;
    IWICBitmapFrameDecode* frame = NULL;
    IWICFormatConverter* conv = NULL;
    HRESULT wicHr = CoCreateInstance(&kWICFactory, NULL, CLSCTX_INPROC_SERVER,
                                &kWICFactoryI, (void**)&fac);
    hg = GlobalAlloc(GMEM_MOVEABLE, size);
    if (hg) {
        void* pv = GlobalLock(hg);
        if (pv) { memcpy(pv, data, size); GlobalUnlock(hg); }
        if (SUCCEEDED(CreateStreamOnHGlobal(hg, TRUE, &st)) && st) {
            HRESULT hrDec = fac->lpVtbl->CreateDecoderFromStream(fac, st, NULL,
                              WICDecodeMetadataCacheOnDemand, &dec);
            HRESULT hrFrame = dec ? dec->lpVtbl->GetFrame(dec, 0, &frame) : -1;
            HRESULT hrConv = (!dec || SUCCEEDED(hrFrame)) ? fac->lpVtbl->CreateFormatConverter(fac, &conv) : -1;
            HRESULT hrInit = conv ? conv->lpVtbl->Initialize(conv, (IWICBitmapSource*)frame,
                              &kPixFmtBGRA, WICBitmapDitherTypeNone, NULL, 0.0,
                              WICBitmapPaletteTypeCustom) : -1;
            if (SUCCEEDED(hrDec) && dec && SUCCEEDED(hrFrame) && frame
                && SUCCEEDED(hrConv) && conv && SUCCEEDED(hrInit)) {
                UINT w = 0, hgt = 0;
                conv->lpVtbl->GetSize(conv, &w, &hgt);
                if (w && hgt) {
                    BITMAPINFO bmi; memset(&bmi, 0, sizeof(bmi));
                    bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
                    bmi.bmiHeader.biWidth = (LONG)w;
                    bmi.bmiHeader.biHeight = -(LONG)hgt;   /* top-down */
                    bmi.bmiHeader.biPlanes = 1;
                    bmi.bmiHeader.biBitCount = 32;
                    bmi.bmiHeader.biCompression = BI_RGB;
                    void* bits = NULL;
                    HBITMAP hb = CreateDIBSection(NULL, &bmi, DIB_RGB_COLORS, &bits, NULL, 0);
                    /* CopyPixels(prc, cbStride, cbBufferSize, buf) — 第三参是**总缓冲**
                       (stride*height), 此前误传单行步长 → INSUFFICIENTBUFFER。 */
                    HRESULT hrCopy = hb ? conv->lpVtbl->CopyPixels(conv, NULL, w * 4, w * 4 * hgt, (BYTE*)bits) : -1;
                    if (hb && bits && SUCCEEDED(hrCopy)) {
                        PICTDESC pd; memset(&pd, 0, sizeof(pd));
                        pd.cbSizeofstruct = sizeof(pd);
                        pd.picType = PICTYPE_BITMAP;
                        pd.bmp.hbitmap = hb;
                        if (FAILED(OleCreatePictureIndirect(&pd, &IID_IPicture, TRUE, (void**)&pic)))
                            pic = NULL;
                        if (!pic) DeleteObject(hb);   /* fOwn=TRUE 失败时句柄归我们收尾 */
                    } else if (hb) {
                        DeleteObject(hb);
                    }
                }
            }
        } else {
            GlobalFree(hg);
        }
    }
    if (conv) conv->lpVtbl->Release(conv);
    if (frame) frame->lpVtbl->Release(frame);
    if (dec) dec->lpVtbl->Release(dec);
    if (st) st->lpVtbl->Release(st);
    if (fac) fac->lpVtbl->Release(fac);
    return pic;
}

vb6_VARIANT vb6_LoadResPicture(vb6_VARIANT resourceId, vb6_VARIANT resourceType) {
    vb6_VARIANT ret; memset(&ret, 0, sizeof(ret));
    // VB6 的 format: 0=vbResBitmap, 1=vbResIcon, 2=vbResCursor; 字符串格式名也收
    // ("BITMAP"/"ICON"/"CURSOR" 及自定义类型名 —— 语料常见把 PNG/JPEG/WebP 整个
    // 文件挂成 "PNG" 之类的自定义资源类型, 这里按**字节签名**分流解码, 不认类型名)。
    const wchar_t* typeName = vb6_ResNameOf(&resourceType);
    int32_t fmt = typeName ? -1 : vb6_ResIdOf(&resourceType);
    LPCWSTR rType;
    if      (typeName)                      rType = typeName;
    else if (fmt == 0 /*vbResBitmap*/)      rType = RT_BITMAP;
    else if (fmt == 1 /*vbResIcon*/)        rType = RT_ICON;
    else if (fmt == 2 /*vbResCursor*/)      rType = RT_CURSOR;
    else                                    rType = RT_BITMAP;  /* 越界值落 bitmap, 比静默空值可排查 */
    int32_t fmtIsIcon = (fmt == 1) || (typeName && wcsicmp(typeName, L"ICON") == 0);
    int32_t fmtIsCur  = (fmt == 2) || (typeName && wcsicmp(typeName, L"CURSOR") == 0);
    int32_t fmtIsBmp  = (fmt == 0) || (typeName && wcsicmp(typeName, L"BITMAP") == 0)
                        || (!typeName && fmt != 0 && fmt != 1 && fmt != 2);

    uint32_t size = 0;
    const void* data = vb6_ResLoad(&resourceId, rType, fmt, &size);
    if ((!data || !size) && (fmtIsIcon || fmtIsCur)) {
        /* rc.exe 展开的图标: id 落在组 (RT_GROUP_ICON/CURSOR), 按组解析出裸图 */
        data = vb6_ResResolveGroup(&resourceId, fmtIsCur, &size);
    }
    if (!data || !size) { vb6_ErrRaiseNumber(326); return ret; }

    vb6_OleEnsureInit();
    IPicture* pic = NULL;
    const uint8_t* d8 = (const uint8_t*)data;
    if (fmtIsCur) {
        // RT_CURSOR 裸图标图 (BITMAPINFOHEADER+XOR/AND), CreateIconFromResourceEx 直接吃。
        HICON hc = CreateIconFromResourceEx((PBYTE)data, size, FALSE, 0x00030000,
                                            0, 0, LR_DEFAULTCOLOR);
        if (hc) {
            PICTDESC pd; memset(&pd, 0, sizeof(pd));
            pd.cbSizeofstruct = sizeof(pd);
            pd.picType = PICTYPE_ICON;   /* OLE PICTDESC 没有 cursor 档、本 SDK 也没有 PICTYPE_CURSOR —— HCURSOR 经 icon 槽传 HANDLE, 取句柄侧不分这两类 */
            pd.icon.hicon = (HICON)hc;
            if (FAILED(OleCreatePictureIndirect(&pd, &IID_IPicture, TRUE, (void**)&pic))) pic = NULL;
        }
    } else if (fmtIsIcon && !(size >= 4 && d8[0] == 0 && d8[1] == 0 && d8[2] == 1)) {
        // RT_ICON 裸图标图 (无 .ico 文件头); 若实际存的是整枚 .ico 文件则落通用流路径。
        HICON hi = CreateIconFromResourceEx((PBYTE)data, size, TRUE, 0x00030000,
                                            0, 0, LR_DEFAULTCOLOR);
        if (hi) {
            PICTDESC pd; memset(&pd, 0, sizeof(pd));
            pd.cbSizeofstruct = sizeof(pd);
            pd.picType = PICTYPE_ICON;
            pd.icon.hicon = hi;
            if (FAILED(OleCreatePictureIndirect(&pd, &IID_IPicture, TRUE, (void**)&pic))) pic = NULL;
        }
    } else if (fmtIsBmp && !(size >= 2 && d8[0] == 0x42 && d8[1] == 0x4D)) {
        // RT_BITMAP 数据 = 打包 DIB (BITMAPINFOHEADER + 调色板 + 位数据, 无文件头)。
        // 补一个 BITMAPFILEHEADER 走 OleLoadPicture 流, 与 vb6_LoadPictureEx 同出口。
        const BITMAPINFOHEADER* bi = (const BITMAPINFOHEADER*)data;
        DWORD colors = bi->biClrUsed ? bi->biClrUsed
                     : (bi->biBitCount <= 8 ? (1u << bi->biBitCount) : 0u);
        uint32_t total = 14 + size;
        BYTE* bmp = (BYTE*)malloc(total);
        if (bmp) {
            BITMAPFILEHEADER* fh = (BITMAPFILEHEADER*)bmp;
            fh->bfType = 0x4D42;                       /* 'BM' */
            fh->bfSize = total;
            fh->bfReserved1 = fh->bfReserved2 = 0;
            fh->bfOffBits = 14 + bi->biSize + colors * 4;
            memcpy(bmp + 14, data, size);
            pic = vb6_ResPicViaWic(bmp, total);
            if (!pic) pic = vb6_ResPicViaOle(bmp, total);
            free(bmp);
        }
    } else {
        // 整张图片文件 (BMP/PNG/JPEG/GIF/ICO/WebP/… 签名): WIC 优先 (新系统上
        // OleLoadPicture 的流路已返 E_FAIL, 且 WIC 覆盖 WebP/已装扩展的 SVG), OLE 兜底。
        pic = vb6_ResPicViaWic(data, size);
        if (!pic) pic = vb6_ResPicViaOle(data, size);
    }
    if (!pic) pic = vb6_ResPicViaWic(data, size);   /* WebP / 装了解码扩展的 SVG 等 */
    if (!pic) { vb6_ErrRaiseNumber(326); return ret; }
    return vb6_VariantObject((void*)pic);
}

vb6_VARIANT vb6_LoadResData(vb6_VARIANT resourceId, vb6_VARIANT resourceType) {
    vb6_VARIANT ret; memset(&ret, 0, sizeof(ret));
    // format: 字符串 (rc 侧的资源类型名, 如 "CUSTOM"/"JSCRIPT"/"BITMAP") 或
    // 数字 (直接当 Win32 资源类型码: 1=CURSOR 2=BITMAP 3=ICON 10=RCDATA…)。
    const wchar_t* typeName = vb6_ResNameOf(&resourceType);
    int32_t typeId = typeName ? 10 /*RT_RCDATA*/ : vb6_ResIdOf(&resourceType);
    uint32_t size = 0;
    const void* data = vb6_ResLoad(&resourceId, typeName, typeId, &size);
    if (!data || !size) { vb6_ErrRaiseNumber(326); return ret; }
    struct vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_byte, 0, (int32_t)size - 1);
    if (!arr) return ret;
    memcpy(arr->data, data, size);
    return vb6_VariantArray(arr);
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

// Load Form — 预加载窗体 (建窗+触发 Form_Load, 不显示)。实现见 vb6forms.c 的
// vb6_LoadForm (抽干延迟 Form_Load 消息); 本 TU 旧版是空桩, 与 vb6forms.c 真实版
// 重定义 (LNK2005), 故删除桩, 单一权威定义落在 vb6forms.c。

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
void*   vb6_UserControl_ContainerHwnd = NULL;  // 账 #180: 句柄成员按指针宽度存
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

// --- PropertyPage 的 Changed 不再有 C 侧裸名字 (账 #219) ---
// 上面那枚 vb6_PropertyPage_Changed 就是它唯一的存储; 以前这里还有一份 `int16_t Changed`,
// 专为"源码里裸写 Changed = True"留的落脚处。发码侧实测从来交的是带前缀那一个
// (语料 vb6_PropertyPage_Changed 186 处、裸名 0 处)，而裸名全局与用户模块级变量共享 C 名字空间
// ⇒ `Public Changed As Long` 直接 C2371 编不过 (探针 .build/b229out/pjChanged.bas 实测 no exe)。

// --- Picture.Line 的模式旗标不再有 C 侧名字 (账 #220) ---
// 以前这里写着 `const int32_t B = 1; const int32_t BF = 2;`，让 parser 原样发出去的裸名
// 有个落脚处。VB6 允许工程里有个叫 B 的模块级变量（`For B = 1 To 3` 这种写法到处都是），
// 而那两枚是**外部链接的 C 全局** ⇒ 撞名直接 C2373 重定义 + C2166，连 exe 都出不来。
// 旗标现在由 parser 在 Line 的 style 位置折成字面量 1/2，RTL 不再需要名字。

// VB6 UserControl.TextWidth/TextHeight: measure with GDI using current Font
// (unit = ScaleMode; simplified to pixels here; Twip handled once hosting lands)
//
// Fix 113k 已回退: 曾尝试按 vb6_UserControl_Font 建 HFONT 选入 DC 再测, 更"正确",
// 但实测使 Charts 2020 整体布局**更差** —— 本实现里 ScaleWidth/ScaleHeight 是
// 像素, 而 .ctl 的布局常数(PT16=(SW+SH)*2.5/100 等)是按 VB6 的 TWIP 语义推的,
// 两者本就不同源; 换成更大的字体度量后饼图被图例挤没、柱图 X 轴标签裁切更重
// (见离屏 dump ucPieChart/ucChartBar). 故保留原"默认 DC 字体测量"行为.
/* 账 #177: 控件坐标/文字量纲的单位表只剩一份权威, 在 vb6forms.c。vb6rtl 与 vb6forms
   是两个模块、不互相 include 头, 故就地 extern (同 vb6com_internal.h 里
   extern vb6_RaiseError 的做法)。 */
extern double vb6_ScalePxToUser(double px, int32_t mode, int vert);
extern double vb6_ScaleUnitsPerPx(int32_t mode, int vert);

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
    /* 账 #177: VB6 的 TextWidth/TextHeight 交的是**控件 ScaleMode 单位**, 不是设备像素。
       vb6_uc_measureText 量的是 GetDC(NULL) 上的像素, 所以这里必须折算一次:
       缇型控件 (语料里 FontMemRes / VbEclipse 的 ucFolder·ucPerspective 那一族) 此前
       拿到的是像素 ⇒ 比同一枚控件的 ScaleWidth 小 15 倍, 图例/标题全挤在一起。
       像素型控件 (Charts 2020 / czUI / VBFlexGrid) 折算系数为 1 ⇒ 读数逐字节不变。
       ⚠ 这一对读的是进程级 vb6_UserControl_ScaleMode, 只在宿主上下文已换入时正确;
       控件代码里请走按实例的那一对 (vb6_UC_TextWidthOf, 账 #178, 由 cgen 用 #define
       把 UserControl.TextWidth 重定向过去)。 */
    return (int32_t)vb6_ScalePxToUser((double)vb6_uc_measureText(text, 1),
                                      vb6_UserControl_ScaleMode, 0);
}

int32_t vb6_UserControl_TextHeight(BSTR text) {
    return (int32_t)vb6_ScalePxToUser((double)vb6_uc_measureText(text, 0),
                                      vb6_UserControl_ScaleMode, 1);
}

/* 账 #177/#178: 原始像素量 (不做单位折算), 给"按实例取 ScaleMode"的那一对宿主出口用。
   vb6_uc_measureText 是本文件 static, 外面只能从这里拿。 */
int32_t vb6_UC_MeasureTextPx(BSTR text, int wantWidth) {
    return vb6_uc_measureText(text, wantWidth);
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
// 账 #177: 单位表**只剩一份** —— vb6forms.c 的 vb6_ScaleUnitsPerPx (Fix 184 把它
// 接上真实 DPI 时, 这里另抄了一张, 于是同一件事两处口径: 那张表纵向也用 LOGPIXELSX,
// 而权威按 vert 分 X/Y)。vb6rtl 与 vb6forms 是两个模块、不互相 include 头,
// 故就地 extern (同 vb6com_internal.h 里 extern vb6_RaiseError 的做法)。
// User(0)/ContainerPosition(8)/ContainerSize(9,10)/unknown 仍按像素 —— 与 Charts 2020
// 的用法一致 (Extender.Left 已是容器像素; 目标 ScaleMode=3=Pixel ⇒ 恒等)。
extern double vb6_ScaleUnitsPerPx(int32_t mode, int vert);

static double vb6_ucScaleToPixels(int32_t mode, int vert) {
    double u = vb6_ScaleUnitsPerPx(mode, vert);
    return (u == 0.0) ? 1.0 : (1.0 / u);
}

// 账 #196 第三条: 这一对是**单位换算的唯一一份实现**，名字不带宿主前缀 —— 因为要接的接收者不止
// UserControl：`picA.ScaleX(...)` / `Me.ScaleX(...)` / `With Picture1 : .ScaleX(...)` / 窗体模块里
// 裸写 `ScaleX(...)` 都是 VB6 的同一件事 (`Object.ScaleX(x, fromScale, toScale)`)，而换算本身只吃
// 那两个显式的 from/to 参数(接收者自己的 ScaleMode 是由发码侧算好后当参数交进来的，见
// `vb6_WindowScaleModeSelf`)。
// 接上之前这四形的形状：显式接收者与 `Me.` 那一形发成 `vb6_ComCallDouble(hwnd, L"ScaleX", …)` ——
// 对一枚假 IDispatch 发 Invoke ⇒ **编得过、链接过、跑起来回个 0**(#143 那一族)；With 那一形发成
// `hwnd.ScaleX(…)` ⇒ **编译不过**(#150 那一族)；窗体模块里裸写那一形发成裸 `ScaleX(…)` ⇒
// 隐式声明，今天只在真工程里被 /OPT:REF 把整个调用者删掉才没响。
// UserControl 那一档保留 `vb6_UserControl_ScaleX/Y` 这两个**名字**是宿主伪成员表的命名契约
// (`vb6_<Host>_<Member>`，见 cgen_util_com.cpp 的 kHostPseudoRows)，它们只是转手到这里。
double vb6_ScaleUnitX(double x, int32_t fromScale, int32_t toScale) {
    double px = x * vb6_ucScaleToPixels(fromScale, 0);
    double f = vb6_ucScaleToPixels(toScale, 0);
    return (f == 0.0) ? x : (px / f);
}

double vb6_ScaleUnitY(double y, int32_t fromScale, int32_t toScale) {
    double px = y * vb6_ucScaleToPixels(fromScale, 1);
    double f = vb6_ucScaleToPixels(toScale, 1);
    return (f == 0.0) ? y : (px / f);
}

double vb6_UserControl_ScaleX(double x, int32_t fromScale, int32_t toScale) {
    return vb6_ScaleUnitX(x, fromScale, toScale);
}

double vb6_UserControl_ScaleY(double y, int32_t fromScale, int32_t toScale) {
    return vb6_ScaleUnitY(y, fromScale, toScale);
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
