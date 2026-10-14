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
        /* ai/032 §6.6: 无符号两档 (缺档时 default 是把数值转 BSTR —— 不是答错而是换型)。 */
        case VT_UI4:    outElem->vt = (vb6_vartype)VT_UI4; outElem->ulVal = varElem.ulVal; break;
        case VT_UI8:    outElem->vt = (vb6_vartype)VT_UI8; outElem->ullVal = varElem.ullVal; break;
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
    HRESULT hrO = OleInitialize(NULL);
    if (SUCCEEDED(hrO)) {
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
    int oinit = vb6_OleEnsureInit();
    IPicture* pPicture = NULL;
    HRESULT hr = OleLoadPicturePath(pathname, NULL, 0, 0, &IID_IPicture, (void**)&pPicture);
    if (FAILED(hr) || !pPicture) {
        /* DETERMINISTIC GDI FALLBACK — oleaut32's own picture parser is
         * process-fragile under the C3 runtime (observed: OleLoadPicturePath /
         * OleLoadPicture / StdPicture::IPersistStream all E_FAIL in-app on valid
         * files, then S_OK in a plain exe; even the tiny in-memory BMP case flips
         * between runs). GDI file decode (LoadImageW) has been stable every run.
         * Missing files must still yield NULL (the app's CaricaIcone loop relies on
         * LoadPicture failing to stop at the first absent file), so only attempt
         * the fallback when the file actually exists. */
        DWORD fa = GetFileAttributesW(pathname);
        if (fa != INVALID_FILE_ATTRIBUTES) {
            HBITMAP hBmp = (HBITMAP)LoadImageW(GetModuleHandleW(NULL), pathname,
                                               IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE);
            if (hBmp) {
                PICTDESC pd;
                memset(&pd, 0, sizeof(pd));
                pd.cbSizeofstruct = sizeof(pd);
                pd.picType = PICTYPE_BITMAP;
                pd.bmp.hbitmap = hBmp;
                pd.bmp.hpal = NULL;
                IPicture* pWrap = NULL;
                if (SUCCEEDED(OleCreatePictureIndirect(&pd, &IID_IPicture, TRUE, (void**)&pWrap))) {
                    
                    vb6_PictureRegister(pWrap);
                    return (void*)pWrap;
                }
                /* OleCreatePictureIndirect is crippled too: hand the raw HBITMAP
                 * over instead. vb6_SetControlPicture's registered==0 path stores
                 * it as a plain GDI handle and paints it (STATIC + STM_SETIMAGE).
                 * Only reachable when oleaut32's COM picture layer is dead. */
                
                return (void*)hBmp;
            }
            /* Last resort: same file via an IStream + OleLoadPicture. GDI's
             * LoadImageW only decodes BMP/ICO/CUR, so for an existing file that
             * reached here (e.g. a JPEG/PNG that LoadImageW cannot read) this is
             * the only remaining decode path. */
            IPicture* pAlt = NULL;
            DWORD rd = 0, sizeF = 0;
            HANDLE hf = CreateFileW(pathname, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, 0, NULL);
            if (hf != INVALID_HANDLE_VALUE) {
                sizeF = GetFileSize(hf, NULL);
                HANDLE hg = GlobalAlloc(GMEM_MOVEABLE, sizeF);
                void* pv = GlobalLock(hg);
                if (pv) { ReadFile(hf, pv, sizeF, &rd, NULL); GlobalUnlock(hg); }
                CloseHandle(hf);
                IStream* st = NULL;
                if (pv && SUCCEEDED(CreateStreamOnHGlobal(hg, TRUE, &st))) {
                    OleLoadPicture(st, (LONG)rd, FALSE, &IID_IPicture, (void**)&pAlt);
                    st->lpVtbl->Release(st);
                } else if (pv) {
                    GlobalFree(hg);
                }
            }
            if (pAlt) pAlt->lpVtbl->Release(pAlt);
        }
        /* VB6 语义: LoadPicture 指到**不存在**的文件要抛运行时错误 53 (File not found),
         * 而不是静默返回 Nothing。3DMenu 的 CaricaIcone 循环正是靠
         *   For i = 0 To 40 : On Error GoTo Err: Set Img(Num,i)=LoadPicture(Nfile) : ... : Next
         * 里"第一个缺帧 ⇒ 抛错 ⇒ 跳 Err ⇒ 退出循环"来定 Totale(Num) 的。
         * 早先这里只 return NULL: 循环跑满 40 次, Totale(Num) 被写成 40 (真实只有 18 帧),
         * Timer_Shift 于是把 22 个空帧喂给 ImgMenu(sel) → Picture 被清空 → AutoRedraw 位图
         * 回到 BackColor → `TrspCol = GetPixel(hdc,0,0)` 读到 0xE0E0E0 而不是图标角上的
         * 透明色 0xFAE4E4 → TransBltNow 的 XOR/AND/XOR 掩码合成只能整片糊掉 ⇒
         * 环顶那枚"中间的图标"消失/被背景色方块盖住。
         * 空 pathname 不抛 (VB6 用 LoadPicture("") 清空 Picture, 合法且无错)。
         * On Error Resume Next 下 vb6_RaiseError 会提前返回 —— CaricaIconeSub/CaricaTitoli/
         * CaricaFondi 那些 Resume Next 站点行为不变。 */
        if (pathname && pathname[0] && GetFileAttributesW(pathname) == INVALID_FILE_ATTRIBUTES) {
            extern void vb6_RaiseError(int32_t errNum, BSTR description);
            vb6_RaiseError(53, vb6_BSTR_FromStr(L"File not found"));
        }
        return NULL;
    }
    
    vb6_PictureRegister(pPicture);
    return (void*)pPicture;
}

// ============================================================
// Picture 双支持注册表 (Fix P-BMP-3D):
//   vb6_LoadPictureEx 返回的是 COM IPicture* (VB6 的 StdPicture), 而设计期/资源
//   图片走的是原始 GDI 句柄 (HBITMAP/HICON)。cgen 对 `.Picture =` 一律发
//   vb6_SetControlPicture(句柄语义), 导致 IPicture* 被当 HBITMAP 用 → 图标画不出。
//   这里维护一张"本管线活 IPicture 对象"的开放寻址哈希表, vb6_SetControlPicture
//   据此安全区分 COM 对象与 GDI 句柄。IPicture* 是用户态指针, 绝不与 GDI
//   句柄 (内核句柄表索引, 值很小) 重叠, 因此无歧义无探测。不主动注销:
//   残留条目顶多把某个解放后的地址误判为 COM(概率可忽略, 且不影响正确性)。
// ============================================================
#define VB6_PICREG_SIZE 4096
static uintptr_t g_vb6_picReg[VB6_PICREG_SIZE];

static uint32_t vb6_picRegHash(uintptr_t p) {
    uintptr_t h = (uintptr_t)(p >> 4);
    h ^= h >> 16;
    h *= 0x9E3779B97F4A7C15ull;
    return (uint32_t)(h & (VB6_PICREG_SIZE - 1));
}

void vb6_PictureRegister(void* picture) {
    if (!picture) return;
    uintptr_t p = (uintptr_t)picture;
    uint32_t i = vb6_picRegHash(p);
    for (uint32_t n = 0; n < VB6_PICREG_SIZE; n++, i = (i + 1) & (VB6_PICREG_SIZE - 1)) {
        if (g_vb6_picReg[i] == p) return;
        if (g_vb6_picReg[i] == 0) { g_vb6_picReg[i] = p; return; }
    }
}

int vb6_PictureIsRegistered(const void* picture) {
    if (!picture) return 0;
    uintptr_t p = (uintptr_t)picture;
    uint32_t i = vb6_picRegHash(p);
    for (uint32_t n = 0; n < VB6_PICREG_SIZE; n++, i = (i + 1) & (VB6_PICREG_SIZE - 1)) {
        if (g_vb6_picReg[i] == p) return 1;
        if (g_vb6_picReg[i] == 0) return 0;
    }
    return 0;
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
// COM 调用结果→vb6_VARIANT 转换
// 账 #305 §B142 (第五十三刀): 「这一枚 Windows VARIANT 的数值住在槽的哪一格」以前答了
// **两遍** —— vb6_VariantFromComResult(134 行) 与 vb6_VariantFromStackVARIANT(121 行) 的
// 23 档逐条相同(剥掉注释 diff 只余四处不同)，而那四处全是**所有权**：BSTR(偷 / 复制)、
// DISPATCH(置 NULL 转移 / 受守卫 AddRef)、SAFEARRAY 的 BSTR 元素(偷指针 / 逐个复制)、
// 以及「没列到的那一档交什么」。两枚公开名照旧留给发码(它们是全仓按名字发射的)，
// 差的那一件事改成一枚参数(steal = 消耗所有权，与原 FromComResult 同口径)。
// 一起拍掉的第二件事：数组性以前是一张**八行的名单**，名单外的 VT_ARRAY|VT_DATE 那一类
// 落进 default 的**数组块**(把标量字段当 SAFEARRAY 解 = SafeArrayGetDim(垃圾)，块尾还
// pv->parray = NULL 抹掉原值)，名单外的标量也走同一块；那块底下那格换 BSTR 的兜底
// 因此从来没被任何一枚走到过。现在数组只问 `vt & VT_ARRAY` 这一面旗标，认不出的标量
// 统一走换 BSTR 那一格。
// ============================================================
static vb6_VARIANT vb6_ComVariantIntoSlot(VARIANT* pv, int steal) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    if (!pv) { result.vt = vb6_vtEmpty; return result; }
    // Windows VARENUM 与 vb6_vartype 值一致, 可直接赋值
    result.vt = (vb6_vartype)pv->vt;

    if (pv->vt & VT_ARRAY) {
        /* P24-02: VT_ARRAY - 将Windows SAFEARRAY转换为vb6_SafeArray1D。
           元素档问 vb6_SaElemTypeFromVt(唯一那张反向表)、宽度交回 vb6_SaElemSize、标签交回
           vb6_SaVtFromElemType —— 这三张表吃同一个枚举，不在这里第二次答(§B159 第五十刀)。 */
        SAFEARRAY* psa = pv->parray;
        if (psa && SafeArrayGetDim(psa) == 1) {
            LONG lBound = 0, uBound = 0;
            SafeArrayGetLBound(psa, 1, &lBound);
            SafeArrayGetUBound(psa, 1, &uBound);
            int32_t count = uBound - lBound + 1;
            VARTYPE baseVt = pv->vt & VT_TYPEMASK;
            vb6_safearray_elemtype et = vb6_SaElemTypeFromVt(baseVt);
            int32_t esz = vb6_SaElemSize(et);

            vb6_SafeArray1D* arr = (vb6_SafeArray1D*)calloc(1, sizeof(vb6_SafeArray1D));
            vb6_Sa1DInit(arr);
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
                    VARIANT* copy = (VARIANT*)malloc(sizeof(VARIANT));
                    VariantInit(copy);
                    VariantCopy(copy, srcElem);
                    *dstElem = vb6_ComVariantIntoSlot(copy, 1);
                }
            } else if (baseVt == VT_BSTR) {
                BSTR* dstBstr = (BSTR*)arr->data;
                for (int32_t i = 0; i < count; i++) {
                    if (steal) {
                        /* 消耗式: 把指针搬过来并清掉源里那一格，免得外层的 VariantClear 放它 */
                        dstBstr[i] = ((BSTR*)srcData)[i];
                        ((BSTR*)srcData)[i] = NULL;
                    } else {
                        dstBstr[i] = SysAllocString(((BSTR*)srcData)[i]);
                    }
                }
            } else {
                // Numeric types: direct memcpy
                memcpy(arr->data, srcData, (size_t)count * esz);
                /* 账 #305 §B142 (第五十三刀)刻意留的一格，别当成已经修好：认不出档的元素
                   (VT_DATE / VT_CY / VT_I8 / VT_DECIMAL / VT_DISPATCH / VT_UI4 ...)在这里回落
                   variant 档 ⇒ 对端一格 8/4/16 字节、我们一格 sizeof(vb6_VARIANT)，这一条
                   memcpy 就是越界读。第五十三刀先试过「问 SafeArrayGetElemsize 拿对端宽度、
                   不符就逐元素装箱」，实测把语义改坏了(见 emit 后的元素读回 3 + 堆写坏 0xC0000374
                   —— 那张宽度答案与尺寸表打架时，逐元素那一支写的是**槽**而描述符按 4 字节步长
                   分配)。宽度这一问要在动手前先量准(§B165 旁边另记)，不在本刀半修。 */
            }
            result.vt = (vb6_vartype)(vb6_vtArray | vb6_SaVtFromElemType(et));
            result.parray = arr;
        }
        if (steal) pv->parray = NULL;  /* 只有消耗式才动源里那一格 */
        return result;
    }

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
            if (steal) {
                result.bstrVal = pv->bstrVal;
                pv->bstrVal = NULL;  // 转移所有权, 防止VariantClear释放
            } else {
                result.bstrVal = SysAllocString(pv->bstrVal);  /* copy (不转移所有权) */
            }
            break;
        case VT_DISPATCH:
            result.pdispVal = pv->pdispVal;
            if (steal) {
                pv->pdispVal = NULL;  // 转移所有权
            } else if (pv->pdispVal) {
                // Fix <vbeclipse>: 用受守卫的 vb6_ComAddRefDispatch (内含 vb6_ComIsDispatchable),
                // 与 vb6_VariantObject 同口径。`For Each ctrl In Controls` 交回的是**裸子控件
                // HWND** (宿主分派层的对象表示), 不是真 IDispatch —— 裸 AddRef 会把 HWND 首字段
                // 当 vtable 解引用 → av read (ucTabStrip.UserControl_Resize 实测)。真 COM 接收者
                // 仍照常 AddRef; 析构侧 vb6_ReleaseObject 同判据跳过, 收支平衡。
                vb6_ComAddRefDispatch(pv->pdispVal);
            }
            break;
        case VT_BOOL:   result.boolVal = pv->boolVal; break;
        case VT_UI1:    result.bVal = pv->bVal; break;
        case VT_ERROR:  result.lVal = pv->scode; break;
        /* 账 #305 §B145: 缺这一档时 FromComResult 的 default 走的是**数组**那段(pv->parray
         * 把 llVal 当指针解) —— 不是「答错数」而是直接垃圾；而 VarAdd(VT_I8, VT_I4) 本来
         * 就交回 vt=20 的结果(实测)，所以这一格是 Variant 算术的必经之路。 */
        case VT_I8:    result.llVal = pv->llVal; break;
        /* ai/032 §6.6: 无符号两档与 VT_I8 同条 —— 缺档时 default 那段会把 ullVal
         * 当数组指针解 (与 §B145 那次同形的垃圾读)。原样带值过界。 */
        case VT_UI4:   result.ulVal = pv->ulVal; break;
        case VT_UI8:   result.ullVal = pv->ullVal; break;
        case VT_DECIMAL:
            /* 账 #305 §B144: 这一档只问唯一出口「DECIMAL 住在哪一字节」—— 两架构的布局
             * 不在这里抄第二遍(这里曾经抄过两次，都是抄错的)。 */
            {
                const DECIMAL *d144 = vb6_DecimalPtrFromVariant(pv);
#ifdef _WIN64
                if (d144) memcpy(&result.decVal, d144, sizeof(result.decVal));
#else
                /* 两种所有权都要一份私有副本: x86 那台的 DECIMAL 是对端 VARIANT **内联**那
                   16 字节(§B144)，借用出去的存活期只到调用方清完它那枚 VARIANT 为止。 */
                result.pdecVal = (DECIMAL *)CoTaskMemAlloc(sizeof(DECIMAL));
                if (result.pdecVal && d144)
                    memcpy(result.pdecVal, d144, sizeof(DECIMAL));
#endif
            }
            break;
        default:
            /* 认不出的标量档(原两枚双胞胎都没走到过这里): 交 oleaut32 换算成串，
               换不动就交 Empty —— 与「答错数」相比，这一格交的是一个读得出的答案。 */
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
    return result;
}

// 此函数**消耗** VARIANT* 的所有权, 调用后原指针不可再使用(交给 VariantClear + free)。
// 用法: v = dic.Item("hello") 这类「COM 交回一枚堆上的 VARIANT」的场景。
vb6_VARIANT vb6_VariantFromComResult(void* variant_ptr) {
    VARIANT* pv = (VARIANT*)variant_ptr;
    if (!pv) {
        vb6_VARIANT empty;
        memset(&empty, 0, sizeof(empty));
        empty.vt = vb6_vtEmpty;
        return empty;
    }
    vb6_VARIANT result = vb6_ComVariantIntoSlot(pv, 1);
    VariantClear(pv);  // 清理原始VARIANT (BSTR/pdispVal 已转移的那几格是 NULL)
    free(pv);
    return result;
}

// P24-03: 栈上 VARIANT → vb6_VARIANT(带权副本, 不动源) —— For Each 那一族用这一枚。
vb6_VARIANT vb6_VariantFromStackVARIANT(VARIANT* pv) {
    return vb6_ComVariantIntoSlot(pv, 0);
}

/* P24-04: Extract IDispatch pointer from Variant for COM late-binding */
/* When vb6_VARIANT holds an object (vt==vb6_vtDispatch), return its pdispVal; */
/* otherwise return NULL. Used when Variant-typed vars access COM members. */
void* vb6_VariantToObject(vb6_VARIANT* v) {
    if (!v) return NULL;
    if (v->vt == vb6_vtDispatch) return v->pdispVal;
    return NULL;
}

/* 账 #305 §B159 (第五十刀): 「这一档元素许不许原样交给对端」。可搬的判据只有一条: 我们那一格的
 * 宽与 SAFEARRAY 对同一个 vt 算出来的宽**逐档相等** —— 于是整块 memcpy 就是正确的编组。
 * bstr / variant 要逐元素深拷贝、udt 的元素尺寸由调用方给、ptr 那一档涉及引用计数:
 * 口径都还没拍 (§B161 记着), 交出去一律答 VT_EMPTY —— 宁可少给, 不猜 (§B144 那条口径)。
 * ⚠ 「这枚数组是哪一档」仍只问 vb6_SaVtFromElemType(vb6rtl_array.c), 这里只补「哪几档许出站」;
 *   VB6 的 VarType 元素码与 Windows 的 VARENUM 同值, 所以那张表的答案可以直接当 vt 用 (第四十六刀注明)。
 */
static VARTYPE vb6_ComSaVtBlittable(vb6_safearray_elemtype et) {
    switch (et) {
        case vb6_sa_bool:
        case vb6_sa_int:
        case vb6_sa_byte:
        case vb6_sa_long:
        case vb6_sa_single:
        case vb6_sa_double:
            return (VARTYPE)vb6_SaVtFromElemType(et);
        default:
            return 0;   /* 0 = 这一档不出站 */
    }
}

/* 账 #305 §B162 (第五十刀): 反向 —— 对端递进来的真 SAFEARRAY → 我们那一格的描述符。
 * 口径 = **拷**（自己建一枚描述符 + 原样搬元素），不是借：服务端这一趟结束就 free 掉槽阵列
 *   (vb6comserver_obj.c 那句 CoTaskMemFree)，而 VB 代码会把这枚 Variant 存进变量、带出这次调用,
 *   借来的 psa->pvData 在那之后就是悬垂的。回写（VB6 的数组按引用传）刻意不做 —— 与 §B160/
 *   §B161 同一口径，那一格补的时候两头一起补。
 * 白名单与出站**同一张表**：元素档认不出、或认出来的档与真数组的 vt 不是同一档，一律交回
 *   vb6_vtEmpty。改前这里是「对端 VARIANT 头 16 字节的位拷贝」—— 于是一枚看着像数组、
 *   parray 却指着一整条 SAFEARRAY 头部的槽被递进被调方 (实测语料那 7 行 oPC.AddPaletteColors
 *   m_PF_Colors 走的正是这一格)。
 */
static void vb6_VariantArrayIntoSlot(const VARIANT *src, vb6_VARIANT *s) {
    SAFEARRAY* psa = src->parray;
    VARTYPE base = (VARTYPE)(src->vt & VT_TYPEMASK);
    vb6_safearray_elemtype et;
    VARTYPE allow;
    LONG lo = 0, hi = 0;
    int32_t n, esz;
    void* srcData = NULL;
    vb6_SafeArray1D* arr;
    void* data;
    s->vt = vb6_vtEmpty;
    if (!psa || SafeArrayGetDim(psa) != 1) return;          /* 多维: 描述符是另一套 (§B151) */
    et = vb6_SaElemTypeFromVt(base);
    allow = vb6_ComSaVtBlittable(et);
    if (!allow || allow != base) return;                   /* 白名单外 / 认不出 = 不猜 */
    SafeArrayGetLBound(psa, 1, &lo);
    SafeArrayGetUBound(psa, 1, &hi);
    n = hi - lo + 1;
    esz = vb6_SaElemSize(et);
    if (n <= 0 || esz <= 0) return;
    if (FAILED(SafeArrayAccessData(psa, &srcData)) || !srcData) {
        if (srcData) SafeArrayUnaccessData(psa);
        return;
    }
    arr = (vb6_SafeArray1D*)calloc(1, sizeof(*arr));
    data = calloc((size_t)n, (size_t)esz);
    if (!arr || !data) {
        free(arr); free(data);
        SafeArrayUnaccessData(psa);
        return;
    }
    memcpy(data, srcData, (size_t)n * (size_t)esz);
    SafeArrayUnaccessData(psa);
    vb6_Sa1DInit(arr);
    arr->elemType = et;
    arr->elemSize = esz;
    arr->lBound = lo;
    arr->uBound = hi;
    arr->count = n;
    arr->data = data;
    arr->isDynamic = 0;
    s->vt = (vb6_vartype)(vb6_vtArray | allow);
    s->parray = arr;
}

/* 账 #305 §B162: 上面那一臂造出来的描述符归**这一趟调用**所有 —— 服务端 free 槽阵列之前
 * 必须经这里，否则它就漏了。别的档一律不动（借用契约：串与对象仍归调用方）。 */
void vb6_VariantSlotRelease(void *slot) {
    vb6_VARIANT* s = (vb6_VARIANT*)slot;
    if (!s || !(s->vt & vb6_vtArray)) return;
    vb6_SafeArrayDestroy1D(s->parray);
    s->parray = NULL;
    s->vt = vb6_vtEmpty;
}

/* 账 #305 §B159 (第五十刀): 我们的数组载体 → 真 WINDOWS SAFEARRAY (出站那一臂唯一的造处)。
 * 认载体只问首字段: 一维是魔数 0x5A1D、多维是秩数 1..16, 魔数落不进 1..16 —— 与
 * vb6_SaVtOfArrayPtr / vb6_SaNdElemPtr 同一形, 不新增第三种认法。
 * 界照抄 (lBound/cElements), 多维那一支不重排: 我们的 ND 是 d=0 变得最快, 与 SAFEARRAY 的列主序同序。
 * 所有权: 新建的 psa 挂在 out->parray 上, 归**收到这一枚 VARIANT 的那一头** VariantClear ——
 *   客户端实参: vb6_ComCall 的发码 epilogue 逐枚 VariantClear (vb6com_invoke.c:277-286);
 *   服务端返回值: 交对端, 由对端 VariantClear (oleaut32 的标准 ABI);
 *   Variant 算术 (vb6_VarArith): 借用契约不许 VariantClear, 那里就地 SafeArrayDestroy。
 * 谁要把这一份留住, 自己 VariantCopy (加引用) —— 载体本体一律不动。
 */
static void vb6_ComPackArrayOut(const void* arrV, VARIANT* out) {
    SAFEARRAYBOUND rg[16];
    const vb6_SafeArray1D* a1 = (const vb6_SafeArray1D*)arrV;
    VARTYPE vt = 0;
    int32_t esz = 0, n = 0, rank = 0;
    const void* data = NULL;
    SAFEARRAY* psa;
    void* pvData = NULL;
    out->vt = VT_EMPTY;
    if (!arrV) return;
    if (vb6_SaIs1D(a1)) {
        vt = vb6_ComSaVtBlittable(a1->elemType);
        if (!vt) return;
        esz = vb6_SaElemSize(a1->elemType);
        /* 步长与那张尺寸表不符 = 这枚载体当年按错尺寸建过 (§B150 记的坑), 数据不可信, 不搬。 */
        if (a1->elemSize != esz || a1->count <= 0 || !a1->data) return;
        rg[0].lLbound = (ULONG)a1->lBound;
        rg[0].cElements = (ULONG)a1->count;
        rank = 1; n = a1->count; data = a1->data;
    } else {
        const vb6_SafeArrayND* nd = (const vb6_SafeArrayND*)arrV;
        int64_t total = 1;
        int32_t d;
        if (nd->dimCount < 1 || nd->dimCount > 16) return;
        vt = vb6_ComSaVtBlittable(nd->elemType);
        if (!vt) return;
        esz = vb6_SaElemSize(nd->elemType);
        if (nd->elemSize != esz || !nd->data) return;
        for (d = 0; d < nd->dimCount; d++) {
            if (nd->bounds[d].cElements <= 0) return;
            rg[d].lLbound = (ULONG)nd->bounds[d].lBound;
            rg[d].cElements = (ULONG)nd->bounds[d].cElements;
            total *= nd->bounds[d].cElements;
            if (total > 0x7FFFFFFFll) return;   /* 溢出/零元素: 不搬 */
        }
        rank = nd->dimCount; n = (int32_t)total; data = nd->data;
    }
    psa = SafeArrayCreate(vt, (ULONG)rank, rg);
    if (!psa) return;
    if (FAILED(SafeArrayAccessData(psa, &pvData)) || !pvData) {
        SafeArrayDestroy(psa);
        return;
    }
    memcpy(pvData, data, (size_t)n * (size_t)esz);
    SafeArrayUnaccessData(psa);
    out->vt = (VARTYPE)(VT_ARRAY | vt);
    out->parray = psa;
}

/* 账 #305 §B145: 槽 → 真 VARIANT 那张 vt 表的**唯一一份**(借用语义: 不复制 BSTR、不 AddRef)。
 * 「住在哪一格」与「谁负责释放」是两个问题，这里只答前一个 —— 出站要带权副本的那一族
 * 在下面 vb6_ComPackVariant 里自己补那一步，于是这张表只有一份，两处不会再各自漏档
 * (§B142 的 census 量到的正是「出站的表」与「Variant 算术」都要它)。
 * ⚠ 接收者叫 `v` 的 `switch (v.vt)` 是 scripts/check_variant_i8_coverage.ps1 认的形状，
 *   改名或换接收者 = 这张表从 census 里隐身(§B143 那条「变目只认一种写法」的老坑)。 */
/* 账 #305 §B159 (第五十刀): 下面那条 switch 吃的是**整枚** vt, 而数组那一档的标签是
 * vb6_vtArray|元素码, 对不上任何一条 case ⇒ 改前一路落 default, 交给 COM 的是 VT_EMPTY
 * (数组**静默变成空值**)。所以数组那一问排在表之前、只问一次, 真数组由上面的
 * vb6_ComPackArrayOut 现造。 */
void vb6_VariantPackBorrow(vb6_VARIANT v, VARIANT *out) {
    VariantInit(out);
    if (v.vt & vb6_vtArray) { vb6_ComPackArrayOut(v.parray, out); return; }
    switch (v.vt) {
        case vb6_vtEmpty: out->vt = VT_EMPTY; break;
        case vb6_vtNull:  out->vt = VT_NULL; break;
        case vb6_vtInteger: out->vt = VT_I2; out->iVal = v.iVal; break;
        case vb6_vtLong:    out->vt = VT_I4; out->lVal = v.lVal; break;
        case vb6_vtSingle:  out->vt = VT_R4; out->fltVal = v.fltVal; break;
        case vb6_vtDouble:  out->vt = VT_R8; out->dblVal = v.dblVal; break;
        case vb6_vtBSTR:    out->vt = VT_BSTR; out->bstrVal = v.bstrVal; break;   /* 借用: 不复制 */
        case vb6_vtDispatch: out->vt = VT_DISPATCH; out->pdispVal = (IDispatch*)v.pdispVal; break;
        case vb6_vtBoolean: out->vt = VT_BOOL; out->boolVal = v.boolVal ? VARIANT_TRUE : VARIANT_FALSE; break;
        case vb6_vtByte:    out->vt = VT_UI1; out->bVal = v.bVal; break;
        /* 账 #305: 以前这一档落 default ⇒ 交出去是 VT_EMPTY —— 地址**静默变成空值**，不是报错。*/
        case VT_I8:       out->vt = VT_I8; out->llVal = v.llVal; break;
        /* ai/032 §6.6: 无符号两档按原档交给 COM (缺档时 default 交 VT_EMPTY —— 
         * 与 VT_I8 那条同一条严重度)。 */
        case (vb6_vartype)VT_UI4: out->vt = VT_UI4; out->ulVal = v.ulVal; break;
        case (vb6_vartype)VT_UI8: out->vt = VT_UI8; out->ullVal = v.ullVal; break;
        /* 账 #305 §B145: §B142 census 点名的那三格 —— Decimal / Currency / Date 以前都落 default
         * ⇒ 交给 COM 或交给 VarAdd 的是 VT_EMPTY，与上面 VT_I8 同一条严重度分级。
         * Decimal 走 §B142 那枚唯一出口(两份布局不在这里抄第二遍)；Currency 那一格是
         * CURRENCY(=LONGLONG) 而不是标量，与上面「往里读」的 *(int64_t*)&pv->cyVal 同形。 */
        case vb6_vtDecimal:  if (!vb6_DecimalAsVariant(v, out)) out->vt = VT_EMPTY; break;
        case vb6_vtCurrency: out->vt = VT_CY;   *(int64_t*)&out->cyVal = v.cyVal; break;
        case vb6_vtDate:     out->vt = VT_DATE; out->date = v.dblVal; break;
        default: out->vt = VT_EMPTY; break;
    }
}

/* P24-04: Pack vb6_VARIANT (by value) into Windows VARIANT for COM call args */
/* Used when Variant-typed variable member access result is passed as COM arg */
void* vb6_ComPackVariant(vb6_VARIANT v) {
    /* Fix 150: 必须用 malloc —— 所有消费方 (vb6_ComSetProp/ComSetPropArg/ComCall)
       都用 free() 释放打包 VARIANT。此前用 CoTaskMemAlloc 而 free() 释放,
       跨分配器释放损坏堆, 表现为 OLEAUT32.dll 内 0xC0000005 (NewTab Theme 赋值实测) */
    VARIANT* pv = (VARIANT*)malloc(sizeof(VARIANT));
    if (!pv) return NULL;
    vb6_VariantPackBorrow(v, pv);
    /* 出站实参要的是**带权**的串: 接收方(IDispatch::Invoke)按自己的副本处理它，而槽里
       那一份还要继续用。对象那一档维持原状(借用，不 AddRef) —— 那是另一格的口径。 */
    if (pv->vt == VT_BSTR && pv->bstrVal) pv->bstrVal = SysAllocString(pv->bstrVal);
    return pv;
}
/* 账 #305 §B144: 服务端**出站**那一格（COM 客户端拿到的返回值）唯一出口。
 * 改前那里是 `memcpy(pVarResult, &tmpRet, sizeof(VARIANT))`，注释写着「槽的前 16 字节
 * 与 OLE VARIANT 内容一致」—— 标量那几档确实一致，VT_DECIMAL 不一致：对端(和 oleaut32)
 * 从**结构体头**读那 16 字节，而我们 x64 把它放在 offset 8、x86 放的是指针。实测两台都
 * 静默答错数(x64 三问全答 0；x86 把堆地址当数值答出 18416048)，一个错误都不报。
 * 现在：表仍问 §B145 那份唯一的 `vb6_VariantPackBorrow`(它自己再问 §B142 的布局出口)，
 * 这里只补「带权」那一步 —— 交出去的那份归对端 VariantClear，所以串要新分配、对象要 AddRef
 * (改前那两档是把槽里的指针直接递出去 = 对端释放我们的引用)。 */
void vb6_VariantPackCopyInto(const void *slot, VARIANT *out) {
    if (!out) return;
    VariantInit(out);
    if (!slot) return;
    vb6_VariantPackBorrow(*(const vb6_VARIANT *)slot, out);
    if (out->vt == VT_BSTR && out->bstrVal) out->bstrVal = SysAllocString(out->bstrVal);
    /* 用受守卫的那枚出口(内含 vb6_ComIsDispatchable)：槽里挂着 VT_DISPATCH 标签的可能是
     * 裸子控件 HWND，直接 ->AddRef() 会把 HWND 首字段当 vtable 解引用 ⇒ av read。
     * ⚠ 本文件按 **C** 编(/std:c11)，`p->AddRef()` 那种 C++ 形态的写法根本进不了编译器。 */
    if (out->vt == VT_DISPATCH && out->pdispVal) vb6_ComAddRefDispatch(out->pdispVal);
}

/* 账 #305 §B144: 服务端**入站**那一格（调用方递进来的实参）唯一出口，语义是**借用**。
 * 不能改问现成的 `vb6_VariantFromComResult` —— 那一枚的契约是**消耗所有权**(把 VT_BSTR /
 * VT_DISPATCH 的指针字段置 NULL 拿走)，而 `IDispatch::Invoke` 的实参数组是对端的，
 * 被调方不许改动它。先 memcpy 一份私有副本再交给它也不行：转移走的是副本里的裸指针，
 * 等于最后 clear 掉的是调用方拥有的那只串。所以这里只搬字节、只问布局出口，谁也不接管。 */
void vb6_VariantBorrowIntoSlot(const VARIANT *src, void *slot) {
    vb6_VARIANT *s = (vb6_VARIANT *)slot;
    memset(s, 0, sizeof(*s));
    if (!src) { s->vt = vb6_vtEmpty; return; }
    /* 账 #305 §B162 (第五十刀): 数组那一档进不了下面那条 memcpy —— 位拷贝会把对端的
     * SAFEARRAY 头部当成我们的描述符递给被调方。这一问排在最前（标签是 VT_ARRAY|元素码）。 */
    if (src->vt & VT_ARRAY) { vb6_VariantArrayIntoSlot(src, s); return; }
    if (src->vt == VT_DECIMAL) {
        const DECIMAL *d = vb6_DecimalPtrFromVariant(src);
        s->vt = (vb6_vartype)VT_DECIMAL;
#ifdef _WIN64
        if (d) memcpy(&s->decVal, d, sizeof(*d));
#else
        s->pdecVal = (DECIMAL *)d;   /* 借用：指向对端内联的那 16 字节，存活期 = 这一趟 Invoke */
#endif
        return;
    }
    memcpy(s, src, sizeof(VARIANT));  /* 其余各档：对端的头 16 字节与我们那一格本来就对齐 */
}

/* 账 #305 §B145: Variant 参与 + - * / 时的**唯一运行期出口**。
 * 改前发码层替运行期猜了一档整数(Variant 操作数一律发成 vb6_VariantToLong(...))，实测两架构
 * 一致: 盒里是 Double 时 1.5+1 答 **3**、1.5*2 答 **4**；是 Decimal 时 1.25+1 答 **2**、
 * 0.1+0.2 答 **0** —— 小数与 Decimal 的精确值整个丢掉(探针 .build/b1095_arith，读数在台账)。
 * 这里不猜档: 两侧装箱成真 VARIANT，交 oleaut32 的 VarAdd/VarSub/VarMul/VarDiv —— VB6 自己的
 * Variant 算术就是这台引擎，于是 0.1+0.2 交 Decimal 0.3(96 位精确)、1.5+1 交 R8 2.5、
 * Null 传播 Null、Date+天数交 Date、1/3 交 scale=28 的 Decimal，舍入与越界政策也一起继承。
 * 只有错误号映射是我们写的，而且**按运算符分家**(两台实测): DISP_E_OVERFLOW 在加减乘是 VB 6
 * (Overflow)，在除法是 VB 11 —— 因为 x86 那台的 1.0/0 交的是 DISP_E_OVERFLOW，而 x64 交
 * DISP_E_DIVBYZERO；DISP_E_TYPEMISMATCH(如 "abc"+1)两架构都交回 VB 13。 */
static vb6_VARIANT vb6_VarArith(int op, vb6_VARIANT a, vb6_VARIANT b) {
    VARIANT va, vb, res;
    vb6_VARIANT out;
    HRESULT hr;
    vb6_VariantPackBorrow(a, &va);
    vb6_VariantPackBorrow(b, &vb);
    VariantInit(&res);
    switch (op) {
    case 0:  hr = VarAdd(&va, &vb, &res); break;
    case 1:  hr = VarSub(&va, &vb, &res); break;
    case 2:  hr = VarMul(&va, &vb, &res); break;
    default: hr = VarDiv(&va, &vb, &res); break;
    }
    /* 两份实参是借用的: 这里绝不许 VariantClear(那会把槽自己的 BSTR / 对象释放掉)。 */
    /* 账 #305 §B159 (第五十刀): 唯一的例外是数组那一臂 —— PackBorrow 对数组交的是**新建**的真
     * SAFEARRAY (借用契约里只有它是新分配的), 它归这一趟所有。不在这两条早退路上都收掉,
     * 每算一次 `arr + 1` 就漏一份。 */
    if (va.vt & VT_ARRAY) { SafeArrayDestroy(va.parray); va.parray = NULL; }
    if (vb.vt & VT_ARRAY) { SafeArrayDestroy(vb.parray); vb.parray = NULL; }
    if (FAILED(hr)) {
        int32_t err;
        if (hr == DISP_E_DIVBYZERO) err = 11;
        else if (hr == DISP_E_OVERFLOW) err = (op == 3) ? 11 : 6;
        else err = 13;
        vb6_ErrRaiseNumber(err);
        return vb6_VariantEmpty();
    }
    out = vb6_VariantFromStackVARIANT(&res);   /* 复制 BSTR / AddRef 对象，所有权留在 res 上 */
    VariantClear(&res);
    /* oleaut32 造的 Decimal 这一族自带标签(b1093 实测: VarDec* 那一族**不写**，而
     * VarAdd 这一族**写**)，但标签只许有一个写者 —— 再补一次是幂等的。 */
    vb6_DecSetTag(out);
    return out;
}

vb6_VARIANT vb6_VarAddV(vb6_VARIANT a, vb6_VARIANT b) { return vb6_VarArith(0, a, b); }
vb6_VARIANT vb6_VarSubV(vb6_VARIANT a, vb6_VARIANT b) { return vb6_VarArith(1, a, b); }
vb6_VARIANT vb6_VarMulV(vb6_VARIANT a, vb6_VARIANT b) { return vb6_VarArith(2, a, b); }
vb6_VARIANT vb6_VarDivV(vb6_VARIANT a, vb6_VARIANT b) { return vb6_VarArith(3, a, b); }


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

/* 账 #305 §B150: 上面那一枚的多维版 —— 发码侧「整枚元素引用交址」的两路 (As Any ByRef
   实参 / VarPtr) 以前按老前缀比文本, 本刀新那枚名不认就静默退化成临时盒 (写不回填)
   或 C2440。参数形态与 vb6_VariantArrayGetND 同形, 所以发码侧只换函数名。
   越界交回错误 9 (唯一那枚 vb6_SaNdElemPtr, 账 #209 那一族的口径); 非 Variant 数组槽
   交 NULL, 与上面一维那条同形。 */
vb6_VARIANT* vb6_VariantArrayElemPtrND(vb6_VARIANT* v, int32_t rank, ...) {
    if (!v || !(v->vt & vb6_vtArray) || !v->parray || rank < 1 || rank > 16) return NULL;
    vb6_SafeArrayND* nd = (vb6_SafeArrayND*)v->parray;
    if (nd->elemType != vb6_sa_variant) return NULL;
    int32_t idx[16];
    va_list ap;
    va_start(ap, rank);
    for (int32_t d = 0; d < rank; d++) idx[d] = va_arg(ap, int32_t);
    va_end(ap);
    if (nd->dimCount < 1 || nd->dimCount > 16) return vb6_VariantArrayElemPtr(v, idx[0]);
    int32_t es = vb6_SaElemSize(nd->elemType);
    if (es <= 0 || !nd->data) { vb6_SaNdElemFail(nd, idx, rank); }
    return (vb6_VARIANT*)vb6_SaNdElemPtr(nd, rank, idx, es);
}

// 账 #305 §B150 (第四十刀): 「这枚槽里的数组第 off 格是什么档」以前只写在
// vb6_VariantArrayGet 的函数体里, 而同一枚槽的**多维**读根本没进过这一问 ——
// 发码侧把 v(1,2) 发成 vb6_VariantArrayGet(&v, 1), 第二个下标整块丢掉 (实测读回空),
// 可同一份产物里 UBound(v, 2) 却按 ND 问 ⇒ 同一个"这枚槽是几维"的决定说了两遍、漏了一遍。
// 现在档位表提出来只有一份, 一维与多维都从它出来; 多维的下标检查走那唯一一枚
// vb6_SaNdElemPtr (越界/NULL 一律抛 9, 与账 #209 那一族同口径)。
static vb6_VARIANT vb6_VariantElemFromSlotData(void* data, int32_t off, vb6_safearray_elemtype t) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    result.vt = vb6_vtEmpty;
    switch (t) {
        case vb6_sa_long:
            result.vt = vb6_vtLong;
            result.lVal = ((int32_t*)data)[off];
            break;
        case vb6_sa_int:
            result.vt = vb6_vtInteger;
            result.iVal = ((int16_t*)data)[off];
            break;
        case vb6_sa_double:
            result.vt = vb6_vtDouble;
            result.dblVal = ((double*)data)[off];
            break;
        case vb6_sa_single:
            result.vt = vb6_vtDouble; /* VB6 Single -> Double promotion */
            result.dblVal = (double)((float*)data)[off];
            break;
        case vb6_sa_bstr:
            result.vt = vb6_vtBSTR;
            result.bstrVal = ((BSTR*)data)[off];
            break;
        case vb6_sa_bool:
            result.vt = vb6_vtBoolean;
            result.boolVal = ((int16_t*)data)[off];
            break;
        case vb6_sa_byte:
            result.vt = vb6_vtInteger; /* VB6 Byte -> Integer promotion */
            result.iVal = (int16_t)((uint8_t*)data)[off];
            break;
        case vb6_sa_variant:
            result = ((vb6_VARIANT*)data)[off];
            break;
        case vb6_sa_ptr:
            result.vt = vb6_vtLong; /* object pointer as Long */
            result.lVal = (int32_t)(intptr_t)((void**)data)[off];
            break;
        case vb6_sa_currency:
            result.vt = vb6_vtLong;
            result.lVal = (int32_t)((int64_t*)data)[off]; /* truncate to Long */
            break;
        default:
            result.vt = vb6_vtEmpty;
            break;
    }
    return result;
}

/* Variant数组索引: 从持有SafeArray的Variant中取/设元素 */
vb6_VARIANT vb6_VariantArrayGet(vb6_VARIANT* v, int32_t index) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    result.vt = vb6_vtEmpty;
    if (!v || !(v->vt & vb6_vtArray) || !v->parray) return result;
    vb6_SafeArray1D* arr = v->parray;
    if (index < arr->lBound || index > arr->uBound) return result;
    return vb6_VariantElemFromSlotData(arr->data, index - arr->lBound, arr->elemType);
}

/* 账 #305 §B150: 晚绑定槽的**多维**元素读。形参表 = (槽, 秩数, 秩数个 int32_t 下标)。
   非数组槽 / 秩数非法交 Empty (与上面一维那一条同形, 越界则由 ND 权威抛 9 —— 那一档
   与一维不一致的旧口径另记在账里, 本刀不去改既有的一维语义)。 */
vb6_VARIANT vb6_VariantArrayGetND(vb6_VARIANT* v, int32_t rank, ...) {
    vb6_VARIANT result;
    memset(&result, 0, sizeof(result));
    result.vt = vb6_vtEmpty;
    if (!v || !(v->vt & vb6_vtArray) || !v->parray || rank < 1 || rank > 16) return result;
    vb6_SafeArrayND* nd = (vb6_SafeArrayND*)v->parray;
    int32_t idx[16];
    va_list ap;
    va_start(ap, rank);
    for (int32_t d = 0; d < rank; d++) idx[d] = va_arg(ap, int32_t);
    va_end(ap);
    /* 兜底: 槽里其实挂着一枚 1D 描述符 (自家把多维折成一维那一族, 见 §B90 / §B151) ——
     * 1D 描述符的首字段是魔数 0x5A1D = 23069, 落不进 1..16, 所以这一问就认得出来。
     * 这种槽改前读的是第一个下标 (VBFlexGridDemo 的 ArrRows(iCol, iRow) 实测就是这个形,
     * 见 .build/b1274_emit_new.txt 那一行), 所以这里退回**同一枚一维出口**: 修的是真 2-D,
     * 不把已经在跑的写法换成错误 9。折叠本身是 §B151, 不在本格。 */
    if (nd->dimCount < 1 || nd->dimCount > 16) return vb6_VariantArrayGet(v, idx[0]);
    int32_t es = (nd->elemType == vb6_sa_udt) ? nd->elemSize : vb6_SaElemSize(nd->elemType);
    if (es <= 0 || !nd->data) { vb6_SaNdElemFail(nd, idx, rank); }
    void* p = vb6_SaNdElemPtr(nd, rank, idx, es);   /* 唯一一枚带上下界检查的 ND 落点 */
    int32_t off = (int32_t)(((char*)p - (char*)nd->data) / es);
    return vb6_VariantElemFromSlotData(nd->data, off, nd->elemType);
}

/* 同上, 写侧: (槽, 秩数, 秩数个 int32_t 下标, 末尾一枚 vb6_VARIANT 值) —— 这个顺序是**故意**
   与 vb6_VariantArraySet(v, index, val) 同形, 发码侧那两处「Get 改 Set」的改写只换函数名、
   把装箱后的值追加在末尾, 不必再拼一次参数。 */
void vb6_VariantArraySetND(vb6_VARIANT* v, int32_t rank, ...) {
    if (!v || !(v->vt & vb6_vtArray) || !v->parray || rank < 1 || rank > 16) return;
    vb6_SafeArrayND* nd = (vb6_SafeArrayND*)v->parray;
    if (nd->elemType != vb6_sa_variant) return;   /* 与一维那条同口径: 只有 Variant 数组可写 */
    int32_t idx[16];
    va_list ap;
    va_start(ap, rank);
    for (int32_t d = 0; d < rank; d++) idx[d] = va_arg(ap, int32_t);
    vb6_VARIANT val = va_arg(ap, vb6_VARIANT);
    va_end(ap);
    if (nd->dimCount < 1 || nd->dimCount > 16) {
        /* 同一枚兜底: 一维描述符按第一个下标写回, 与改前一维那条形 */
        vb6_VariantArraySet(v, idx[0], val);
        return;
    }
    int32_t es = vb6_SaElemSize(nd->elemType);
    if (es <= 0 || !nd->data) { vb6_SaNdElemFail(nd, idx, rank); }
    vb6_VARIANT* slot = (vb6_VARIANT*)vb6_SaNdElemPtr(nd, rank, idx, es);
    /* 接管语义 (rev16 那一课): Clear + 浅拷贝会让槽与借来的 Variant 共用同一只 BSTR。 */
    vb6_VariantAssign(slot, val);
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
vb6_ComIface_Font vb6_UserControl_FontObj = { NULL, 8.25f, 0, 0, 0, 0, 400, 0 };

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
int32_t vb6_UserControl_BackStyle = 1;   // Fix 255: VB6 的 BackStyle 缺省 Opaque (1)
int32_t vb6_UserControl_ForeColor = 0;
int16_t vb6_UserControl_RightToLeft = 0;
void*   vb6_UserControl_ParentControls = NULL;
void*   vb6_UserControl_Controls = NULL;
void*   vb6_Screen_MouseIcon = NULL;       // Fix <vbeclipse>: Screen.MouseIcon 槽   // Fix <vbeclipse>: UserControl.Controls (集合未建模, 恒 NULL)
vb6_ComIface_Font* vb6_UserControl_Font = &vb6_UserControl_FontObj;
struct vb6_UserControl_Ambient_Type vb6_UserControl_Ambient = { &vb6_UserControl_FontObj };

// --- Ambient host environment ---
vb6_ComIface_Font* vb6_Ambient_Font = &vb6_UserControl_FontObj;
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

// --- Fix 214: 宿主伪对象其余成员 (ComCtlsDemo / ComCtls 控件组) ---
// 声明在 vb6rtl_userctl.h。设计期/运行期都由控件自己的属性过程收发 (见那头部的注释),
// 初值取 VB6 默认: Appearance=0 (Flat), DrawStyle=0 (vbSolid),
// AccessKeys/DisplayAsDefault/Default/Cancel 空或 0。
int32_t vb6_UserControl_Appearance   = 0;
int32_t vb6_UserControl_DrawStyle    = 0;
BSTR    vb6_UserControl_AccessKeys   = NULL;
int16_t vb6_Ambient_DisplayAsDefault = 0;
int16_t vb6_Extender_vb6_Default = 0;
int16_t vb6_Extender_Cancel  = 0;

// --- Fix 133u: Extender.Visible/Height (czUI.ctl) ---
// Fix 221: 第三格 DataChanged (ComCtlsDemo 的 `UserControl.Extender.DataChanged = True`,
// 32 处). 运行期容器不活动 ⇒ 只写不读, 保留可写字段。
struct vb6_UserControl_Extender_Type vb6_UserControl_Extender = { -1, 0, 0 };

// Fix 221: UserControl.EventsFrozen —— .ctl 里只出现在读位置 (见 vb6rtl_userctl.h
// 的注释), VB6 运行期恒 False ⇒ 恒 0 的槽位就是正确值。
int16_t vb6_UserControl_EventsFrozen = 0;

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

// Fix 224: 这两个名字在 vb6rtl_userctl.h 里已交 `_Generic` 宏按首参类型分派 (VB6 的
// Variant 隐式数值转换), 故**定义必须挂 _Raw** —— 否则宏会把这里的定义整形成乱语。
double vb6_UserControl_ScaleXRaw(double x, int32_t fromScale, int32_t toScale) {
    return vb6_ScaleUnitX(x, fromScale, toScale);
}

double vb6_UserControl_ScaleYRaw(double y, int32_t fromScale, int32_t toScale) {
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

// 族4a: UserControl.SetFocus —— 与 vb6_Extender_SetFocus 同形 (ComCtls ListView.ctl 的
// FColumnHeaderEditFilter 经 `vb6_UserControl_SetFocus()` 把焦点交还宿主 UserControl)。
// 旧日志装饰名 _vb6_UserControl_SetFocus@0 印证 0 实参。
void vb6_UserControl_SetFocus(void) {
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
