// vb6rtl_array.c - VB6 运行时库: 数组家族：SAFEARRAY 一维与多维实现
// 2026-09-17 从 src/rtl/core/vb6rtl/vb6rtl.c 按家族拆出（纯搬移，逐行未改）:
//   原第 3183~3373 行
//   原第 3374~3617 行

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
// SAFEARRAY - VB6 数组实现
// ============================================================

// 安全数组元素大小表
static int32_t vb6_sa_elem_size(vb6_safearray_elemtype t) {
    switch (t) {
        case vb6_sa_bool:    return (int32_t)sizeof(int16_t);
        case vb6_sa_byte:    return (int32_t)sizeof(uint8_t);
        case vb6_sa_int:     return (int32_t)sizeof(int16_t);
        case vb6_sa_long:    return (int32_t)sizeof(int32_t);
        case vb6_sa_single:  return (int32_t)sizeof(float);
        case vb6_sa_double:  return (int32_t)sizeof(double);
        case vb6_sa_bstr:    return (int32_t)sizeof(BSTR);
        case vb6_sa_variant: return (int32_t)sizeof(vb6_VARIANT);
        case vb6_sa_ptr:     return (int32_t)sizeof(void*);
        case vb6_sa_currency: return (int32_t)sizeof(int64_t);  /* P1-9修复: VB6 Currency = 8bytes */
        case vb6_sa_udt:     return 0;  /* UDT: elemSize set externally, cannot determine here */
        default:             return 4;
    }
}

vb6_SafeArray1D* vb6_SafeArrayCreate1D(vb6_safearray_elemtype elemType,
    int32_t lBound, int32_t uBound) {
    vb6_SafeArray1D* arr = (vb6_SafeArray1D*)calloc(1, sizeof(vb6_SafeArray1D));
    if (!arr) return NULL;
    arr->signature = 0x5A1D;  /* Fix 082g: 1D array magic */
    arr->elemType = elemType;
    arr->elemSize = vb6_sa_elem_size(elemType);
    arr->lBound = lBound;
    arr->uBound = uBound;
    arr->count = uBound - lBound + 1;
    arr->isDynamic = 0;
    if (arr->count > 0) {
        arr->data = calloc((size_t)arr->count, (size_t)arr->elemSize);
    }
    return arr;
}

vb6_SafeArray1D* vb6_SafeArrayReDim1D(vb6_safearray_elemtype elemType,
    int32_t lBound, int32_t uBound) {
    vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(elemType, lBound, uBound);
    if (arr) arr->isDynamic = 1;
    return arr;
}

vb6_SafeArray1D* vb6_SafeArrayReDim1D_Udt(int32_t elemSize,
    int32_t lBound, int32_t uBound) {
    vb6_SafeArray1D* arr = (vb6_SafeArray1D*)calloc(1, sizeof(vb6_SafeArray1D));
    if (!arr) return NULL;
    arr->signature = 0x5A1D;  /* Fix 082g: 1D array magic */
    arr->elemType = vb6_sa_udt;
    arr->elemSize = elemSize;
    arr->lBound = lBound;
    arr->uBound = uBound;
    arr->count = uBound - lBound + 1;
    arr->isDynamic = 1;
    if (arr->count > 0) {
        arr->data = calloc((size_t)arr->count, (size_t)arr->elemSize);
    }
    return arr;
}

// Fix <vbeclipse> rev3 (2026-09-30): 本函数**收 elemType**。
// 此前签名没有它, arr==NULL 时只能回落 vb6_sa_variant 做兜底, 于是**任何非 Variant
// 数组的首次 ReDim Preserve 都按 variant 步长分配**：
//   `Dim m_Keys() As String` + `ReDim Preserve m_Keys(Idx)`
//     → 按 16 字节(x86) / 24 字节(x64) 每元素分配, 但 codegen 发的
//       `VB6_SA_AT(BSTR, arr, i) = ...` 按 **4 字节**步进 —— 第 i 个"槽"里其实
//       压着别的元素的 BSTR; 且 elemType 被记成 variant, 于是
//       vb6_SafeArrayDestroy1D 走 Variant 分支, 把那些 BSTR 当 bstrVal **二次 free**。
// 实测 play78.exe: 崩在 `List.Contains` 的 `For i = 0 To UBound(m_Keys)` 里,
// 反汇编 = `mov eax,[edx+ecx*4]` + `call StrComp` (rva 0x18f6), 读出垃圾 BSTR。
// codegen 本来就知道元素类型 (mapSaElemType), 现在把它传下来, 运行时不再猜。
vb6_SafeArray1D* vb6_SafeArrayReDimPreserve1D_T(vb6_safearray_elemtype elemType,
    vb6_SafeArray1D* arr, int32_t newLBound, int32_t newUBound) {
    // UDT 不能走这里 (元素尺寸由 sizeof 给, 见 _Udt 入口)。若真传进来 udt,
    // vb6_sa_elem_size 会返回 0 → calloc(n, 0) 拿到一个**非空但零长**的指针,
    // 后续按 VB6_SA_AT(UDT,…) 写入就是纯堆越界。回落 variant (只多分配不少分配)。
    if (elemType == vb6_sa_udt) elemType = vb6_sa_variant;

    // 首次 ReDim (arr == NULL): 按**调用方给的**元素类型建, 不再猜。
    if (!arr) return vb6_SafeArrayReDim1D(elemType, newLBound, newUBound);

    // 已有数组但元素类型与本次声明不符 (历史按错误尺寸分配过): 数据不可保留,
    // 按正确尺寸重建。这条让"改过的编译器重编旧工程"也能自愈。
    if (elemType != vb6_sa_udt && arr->elemType != vb6_sa_udt
        && arr->elemType != elemType) {
        vb6_SafeArrayDestroy1D(arr);
        return vb6_SafeArrayReDim1D(elemType, newLBound, newUBound);
    }

    int32_t newCount = newUBound - newLBound + 1;
    if (newCount <= 0) {
        vb6_SafeArrayDestroy1D(arr);
        return NULL;
    }

    // 分配新数据区 (零初始化)
    void* newData = calloc((size_t)newCount, (size_t)arr->elemSize);
    if (!newData) return arr;  // 分配失败, 返回原数组

    // 复制旧数据到新区域 (取交集)
    int32_t copyStart = (arr->lBound > newLBound) ? arr->lBound : newLBound;
    int32_t copyEnd   = (arr->uBound < newUBound) ? arr->uBound : newUBound;
    if (copyStart <= copyEnd) {
        int32_t srcOff = copyStart - arr->lBound;
        int32_t dstOff = copyStart - newLBound;
        int32_t copyLen = (copyEnd - copyStart + 1) * arr->elemSize;
        memcpy((char*)newData + dstOff * arr->elemSize,
               (char*)arr->data + srcOff * arr->elemSize,
               (size_t)copyLen);
    }

    // BSTR元素: 旧区域中被丢弃的元素需要释放
    if (arr->elemType == vb6_sa_bstr) {
        for (int32_t i = arr->lBound; i <= arr->uBound; i++) {
            // 在新范围之外的BSTR需要释放
            if (i < newLBound || i > newUBound) {
                BSTR* slot = &VB6_SA_AT(BSTR, arr, i);
                if (*slot) vb6_BSTR_Free(*slot);
            }
        }
    }

    free(arr->data);
    arr->data = newData;
    arr->lBound = newLBound;
    arr->uBound = newUBound;
    arr->count = newCount;
    return arr;
}

// 兼容入口: 元素类型未知 (无 elemType 形参的老调用点)。按 variant 兜底 ——
// 元素尺寸取最大, 只可能多分配不少分配。**新代码别用这个**。
vb6_SafeArray1D* vb6_SafeArrayReDimPreserve1D(vb6_SafeArray1D* arr,
    int32_t newLBound, int32_t newUBound) {
    return vb6_SafeArrayReDimPreserve1D_T(vb6_sa_variant, arr, newLBound, newUBound);
}

// UDT版ReDim Preserve: 元素尺寸由调用方提供 (sizeof(UDT))。
// 关键修复: 当 arr==NULL 时, 不能回落 vb6_sa_empty(4字节), 否则 UDT 元素写入越界。
// 若 arr 已存在但 elemSize 与本次不符 (曾用错误尺寸分配), 则丢弃旧数据按正确尺寸重建。
vb6_SafeArray1D* vb6_SafeArrayReDimPreserve1D_Udt(int32_t elemSize,
    vb6_SafeArray1D* arr, int32_t newLBound, int32_t newUBound) {
    if (!arr) return vb6_SafeArrayReDim1D_Udt(elemSize, newLBound, newUBound);

    if (arr->elemSize != elemSize) {
        // 旧数组用错误元素尺寸分配 (如历史 4 字节), 数据不可保留, 直接重建
        vb6_SafeArrayDestroy1D(arr);
        return vb6_SafeArrayReDim1D_Udt(elemSize, newLBound, newUBound);
    }

    int32_t newCount = newUBound - newLBound + 1;
    if (newCount <= 0) {
        vb6_SafeArrayDestroy1D(arr);
        return NULL;
    }

    void* newData = calloc((size_t)newCount, (size_t)elemSize);
    if (!newData) return arr;  // 分配失败, 返回原数组

    int32_t copyStart = (arr->lBound > newLBound) ? arr->lBound : newLBound;
    int32_t copyEnd   = (arr->uBound < newUBound) ? arr->uBound : newUBound;
    if (copyStart <= copyEnd) {
        int32_t srcOff = copyStart - arr->lBound;
        int32_t dstOff = copyStart - newLBound;
        int32_t copyLen = (copyEnd - copyStart + 1) * elemSize;
        memcpy((char*)newData + dstOff * elemSize,
               (char*)arr->data + srcOff * elemSize,
               (size_t)copyLen);
    }

    free(arr->data);
    arr->data = newData;
    arr->elemSize = elemSize;
    arr->lBound = newLBound;
    arr->uBound = newUBound;
    arr->count = newCount;
    return arr;
}

void vb6_SafeArrayDestroy1D(vb6_SafeArray1D* arr) {
    if (!arr) return;
    // BSTR元素: 逐个释放
    if (arr->elemType == vb6_sa_bstr && arr->data) {
        for (int32_t i = 0; i < arr->count; i++) {
            BSTR* slot = (BSTR*)((char*)arr->data + i * arr->elemSize);
            if (*slot) vb6_BSTR_Free(*slot);
        }
    }
    // vb6_VARIANT元素: 逐个清理BSTR
    if (arr->elemType == vb6_sa_variant && arr->data) {
        for (int32_t i = 0; i < arr->count; i++) {
            vb6_VARIANT* slot = (vb6_VARIANT*)((char*)arr->data + i * arr->elemSize);
            if (slot->vt == vb6_vtBSTR && slot->bstrVal) {
                vb6_BSTR_Free(slot->bstrVal);
            }
        }
    }
    if (arr->data) free(arr->data);
    free(arr);
}

// Fix 170: VB6 整体数组赋值 `A() = B()` —— 返回 src 的**深拷贝**新载体, 并销毁原 dst。
// 不能直接把 src 的指针赋给 dst: 两个名字会指向同一个 vb6_SafeArray1D, 作用域结束时
// 各自 vb6_SafeArrayDestroy1D → double free (且一侧 ReDim 会神秘改变另一侧内容)。
// dst==src 时同样安全 (先建副本再销毁 dst)。
vb6_SafeArray1D* vb6_ArrayAssign1D(vb6_SafeArray1D* dst, vb6_SafeArray1D* src) {
    return vb6_ArrayAssign1D_Cb(dst, src, NULL);
}

// Fix 178: 见 vb6rtl_array.h 的注释 —— cb 非空时逐元素深拷贝 (vb6_sa_udt 载体)。
vb6_SafeArray1D* vb6_ArrayAssign1D_Cb(vb6_SafeArray1D* dst, vb6_SafeArray1D* src,
                                      vb6_udt_elem_copy cb) {
    vb6_SafeArray1D* res = NULL;
    // 两个名字已共用同一载体 (Fix 178 之前的旧别名) 时: 克隆一份且**不销毁** dst,
    // 否则 src 侧会跟着悬垂。
    int32_t sameCarrier = (dst != NULL && dst == src);
    if (src) {
        res = (vb6_SafeArray1D*)calloc(1, sizeof(vb6_SafeArray1D));
        if (!res) { return dst; }
        res->signature = 0x5A1D;  /* 与 Create1D/ReDim1D_Udt 同魔数 (Fix 082g) */
        res->elemType = src->elemType;
        res->elemSize = src->elemSize;
        res->lBound   = src->lBound;
        res->uBound   = src->uBound;
        res->count    = src->count;
        res->isDynamic = src->isDynamic;
        if (src->count > 0 && src->data && src->elemSize > 0) {
            size_t bytes = (size_t)src->count * (size_t)src->elemSize;
            res->data = malloc(bytes);
            if (!res->data) { free(res); return dst; }
            memcpy(res->data, src->data, bytes);
            // 持有所有权的元素类型: memcpy 后每个槽位仍指向 src 的对象, 需逐个复制
            if (res->elemType == vb6_sa_bstr) {
                for (int32_t i = 0; i < res->count; i++) {
                    BSTR* slot = (BSTR*)((char*)res->data + (size_t)i * res->elemSize);
                    if (*slot) *slot = SysAllocString(*slot);
                }
            } else if (res->elemType == vb6_sa_variant) {
                for (int32_t i = 0; i < res->count; i++) {
                    vb6_VARIANT* d = (vb6_VARIANT*)((char*)res->data + (size_t)i * res->elemSize);
                    vb6_VARIANT* s = (vb6_VARIANT*)((char*)src->data + (size_t)i * src->elemSize);
                    // 裸 memcpy 让 d 与 src 共享 BSTR/对象指针: 必须**清空**(而非 Clear)
                    // 后再 VariantCopy, 否则会把 src 仍持有的对象释放掉。
                    memset(d, 0, sizeof(vb6_VARIANT));
                    vb6_VariantCopy(d, s);
                }
            } else if (cb) {
                // Fix 178: UDT 元素 —— 槽位刚被 memcpy 成 src 元素的位拷贝, 交给
                // 生成的 vb6_udtcpy_<T> 逐成员建立新副本 (String/子数组/Variant)。
                for (int32_t i = 0; i < res->count; i++) {
                    void* d = (char*)res->data + (size_t)i * res->elemSize;
                    const void* s = (const char*)src->data + (size_t)i * src->elemSize;
                    cb(d, s);
                }
            }
        }
    }
    if (dst && !sameCarrier) vb6_SafeArrayDestroy1D(dst);
    return res;
}

void* vb6_SafeArrayGetPtr(vb6_SafeArray1D* arr, int32_t index) {
    if (!arr || index < arr->lBound || index > arr->uBound) return NULL;
    return (char*)arr->data + (index - arr->lBound) * arr->elemSize;
}

void vb6_SafeArrayPutElem(vb6_SafeArray1D* arr, int32_t index, void* value) {
    if (!arr || index < arr->lBound || index > arr->uBound) return;
    void* dest = (char*)arr->data + (index - arr->lBound) * arr->elemSize;
    // BSTR: 先释放旧值, 再赋新值
    if (arr->elemType == vb6_sa_bstr) {
        BSTR* slot = (BSTR*)dest;
        if (*slot) vb6_BSTR_Free(*slot);
        BSTR newVal = *(BSTR*)value;
        *slot = newVal;
    } else {
        memcpy(dest, value, (size_t)arr->elemSize);
    }
}

int32_t vb6_UBound(vb6_SafeArray1D* safeArray, int32_t dimension) {
    // Fix <vbeclipse> rev4 (2026-09-30): 数组状态追踪, 由 C3_SA_TRACE=1 门控。
    // 崩溃现场 `[edx+ecx*4]` 里 edx(=arr->data)是 NULL 而 uBound>=0 —— 只有把
    // signature/elemType/elemSize/lBound/uBound/count/data 全打出来才能判是
    // "结构本身是垃圾"还是"结构对但 data 没分配"。纯 ASCII 输出 (窄流+C locale)。
    if (getenv("C3_SA_TRACE")) {
        fprintf(stderr, "[SA] UBound arr=%p sig=0x%X et=%d es=%d lb=%d ub=%d cnt=%d data=%p\n",
                (void*)safeArray,
                safeArray ? (unsigned)safeArray->signature : 0u,
                safeArray ? (int)safeArray->elemType : -1,
                safeArray ? (int)safeArray->elemSize : -1,
                safeArray ? (int)safeArray->lBound : -999,
                safeArray ? (int)safeArray->uBound : -999,
                safeArray ? (int)safeArray->count : -999,
                safeArray ? safeArray->data : NULL);
        fflush(stderr);
    }
    // Fix <vbeclipse> rev2 (2026-09-29, 用户在真 VB6 里实测确认):
    // `UBound(<未分配的动态数组>)` 在 VB6 是**运行时错误 9 (下标越界)**,
    // 不是返回某个数。工程正是靠 `On Error GoTo ErrorHandle` 接住它实现
    // "空表"语义 (List.cls: Contains=False / Count=-1 都写在错误处理器里)。
    // 本函数此前两个版本都是**猜的**:
    //   v1: return 0  → `For i = 0 To UBound(x)` 跑一次, VB6_SA_AT 读 NULL+0xc
    //                    → 0xC0000005 (实测 List.cls Contains);
    //   v2: return -1 → 常见 For 循环零迭代, 看似没事, 但**没有 On Error 的
    //                    调用点被静默放过** —— 真 VB6 会报错, 这里却吞掉。
    // 现按 VB6 语义抛 9: 有 On Error 的过程 longjmp 到它的处理器 (与 VB6 一致),
    // 没有的走 vb6_ErrRaise 的未处理路径 (报错退出, 也与 VB6 编译版一致)。
    if (!safeArray) {
        vb6_ErrRaise(9, vb6_BSTR_FromStr(L"VBA.Information"),
                     vb6_BSTR_FromStr(L"Subscript out of range"));
        return -1;  /* 不可达: vb6_ErrRaise 必 longjmp 或 ExitProcess */
    }
    // Fix 082g: use signature field to distinguish 1D vs ND arrays
    // SafeArray1D has signature=0x5A1D, SafeArrayND has dimCount (2-16)
    if (safeArray->signature == 0x5A1D) {
        // Confirmed 1D array
        return safeArray->uBound;
    }
    // Not a 1D array - try ND path
    if (dimension > 1) {
        return vb6_UBoundND((vb6_SafeArrayND*)safeArray, dimension);
    }
    // Fallback: check if it looks like an ND array
    int32_t possibleDimCount = *(int32_t*)safeArray;
    if (possibleDimCount >= 2 && possibleDimCount <= 16) {
        vb6_SafeArrayND* ndArr = (vb6_SafeArrayND*)safeArray;
        if (ndArr->totalElements > 0 && ndArr->data != NULL) {
            return vb6_UBoundND(ndArr, dimension);
        }
    }
    return safeArray->uBound;
}

int32_t vb6_LBound(vb6_SafeArray1D* safeArray, int32_t dimension) {
    // 与 vb6_UBound 同族: 未分配的动态数组在 VB6 里 LBound 也是错误 9
    // (不是返回 0) —— 见 vb6_UBound rev2 的说明。
    if (!safeArray) {
        vb6_ErrRaise(9, vb6_BSTR_FromStr(L"VBA.Information"),
                     vb6_BSTR_FromStr(L"Subscript out of range"));
        return 0;  /* 不可达 */
    }
    // Fix 082g: use signature field to distinguish 1D vs ND arrays
    if (safeArray->signature == 0x5A1D) {
        return safeArray->lBound;
    }
    if (dimension > 1) {
        return vb6_LBoundND((vb6_SafeArrayND*)safeArray, dimension);
    }
    int32_t possibleDimCount = *(int32_t*)safeArray;
    if (possibleDimCount >= 2 && possibleDimCount <= 16) {
        vb6_SafeArrayND* ndArr = (vb6_SafeArrayND*)safeArray;
        if (ndArr->totalElements > 0 && ndArr->data != NULL) {
            return vb6_LBoundND(ndArr, dimension);
        }
    }
    return safeArray->lBound;
}

// 账 #209: VB6_SA_AT 的冷路径 —— 未分配的动态数组或下标越界。口径与上面
// vb6_UBound/vb6_LBound 的 rev2 完全一致 (用户在真 VB6 实测: 都是错误 9,
// 有 On Error 走处理器, 没有就报错退出), 这里只多带一条越界读数便于定位。
void vb6_SaElemFail(void* arr, int32_t idx) {
    vb6_SafeArray1D* a = (vb6_SafeArray1D*)arr;
    if (getenv("C3_SA_TRACE")) {  // 与 vb6_UBound rev4 同一个开关
        fprintf(stderr, "[SA] elem access out of range: arr=%p idx=%d lb=%d ub=%d\n",
                (void*)a, (int)idx,
                a ? (int)a->lBound : -999, a ? (int)a->uBound : -999);
        fflush(stderr);
    }
    vb6_ErrRaise(9, vb6_BSTR_FromStr(L"VBA.Information"),
                 vb6_BSTR_FromStr(L"Subscript out of range"));
    exit(9);  /* 不可达: vb6_ErrRaise 必 longjmp 或 ExitProcess */
}

// 账 #209 同族: 多维那一支的冷路径。口径与上面那条一模一样 (VB6 的越界就是错误 9),
// 只是把秩数与逐维下标带出来 —— 多维最容易犯的错是"某一维抄错上下界", 光一个 idx 说不清。
void vb6_SaNdElemFail(void* arr, const int32_t* idx, int32_t rank) {
    vb6_SafeArrayND* a = (vb6_SafeArrayND*)arr;
    if (getenv("C3_SA_TRACE")) {  // 与 vb6_UBound rev4 / vb6_SaElemFail 同一个开关
        fprintf(stderr, "[SA] ND elem access out of range: arr=%p rank=%d dimCount=%d",
                (void*)a, (int)rank, a ? (int)a->dimCount : -999);
        for (int32_t d = 0; a && d < rank && d < 16; d++) {
            fprintf(stderr, " d%d=%d[lb=%d n=%d]", (int)d, (int)idx[d],
                    (int)a->bounds[d].lBound, (int)a->bounds[d].cElements);
        }
        fprintf(stderr, "\n");
        fflush(stderr);
    }
    vb6_ErrRaise(9, vb6_BSTR_FromStr(L"VBA.Information"),
                 vb6_BSTR_FromStr(L"Subscript out of range"));
    exit(9);  /* 不可达 */
}

// ============================================================
// SAFEARRAY ND - VB6 多维数组实现
// ============================================================

vb6_SafeArrayND* vb6_SafeArrayCreateND(vb6_safearray_elemtype elemType,
    int32_t dimCount, vb6_SafeArrayBound bounds[]) {
    if (dimCount <= 0 || dimCount > 16) return NULL;

    vb6_SafeArrayND* arr = (vb6_SafeArrayND*)calloc(1, sizeof(vb6_SafeArrayND));
    if (!arr) return NULL;

    arr->dimCount = dimCount;
    arr->elemType = elemType;
    arr->elemSize = vb6_sa_elem_size(elemType);

    int32_t total = 1;
    for (int32_t d = 0; d < dimCount; d++) {
        arr->bounds[d] = bounds[d];
        if (bounds[d].cElements <= 0) {
            arr->totalElements = 0;
            arr->data = NULL;
            return arr;
        }
        total *= bounds[d].cElements;
    }
    arr->totalElements = total;

    if (total > 0) {
        arr->data = calloc((size_t)total, (size_t)arr->elemSize);
        if (!arr->data) {
            free(arr);
            return NULL;
        }
    }

    return arr;
}

void vb6_SafeArrayDestroyND(vb6_SafeArrayND* arr) {
    if (!arr) return;

    if (arr->elemType == vb6_sa_bstr && arr->data) {
        for (int32_t i = 0; i < arr->totalElements; i++) {
            BSTR* slot = (BSTR*)((char*)arr->data + i * arr->elemSize);
            if (*slot) vb6_BSTR_Free(*slot);
        }
    }
    if (arr->elemType == vb6_sa_variant && arr->data) {
        for (int32_t i = 0; i < arr->totalElements; i++) {
            vb6_VARIANT* slot = (vb6_VARIANT*)((char*)arr->data + i * arr->elemSize);
            if (slot->vt == vb6_vtBSTR && slot->bstrVal) {
                vb6_BSTR_Free(slot->bstrVal);
            }
        }
    }

    if (arr->data) free(arr->data);
    free(arr);
}

int32_t vb6_SafeArrayND_Offset(vb6_SafeArrayND* arr, int32_t dimCount, int32_t indices[]) {
    int32_t offset = indices[0] - arr->bounds[0].lBound;
    int32_t stride = arr->bounds[0].cElements;
    for (int32_t d = 1; d < dimCount; d++) {
        offset += (indices[d] - arr->bounds[d].lBound) * stride;
        stride *= arr->bounds[d].cElements;
    }
    return offset;
}

void* vb6_SafeArrayND_GetPtr(vb6_SafeArrayND* arr, ...) {
    if (!arr) return NULL;
    int32_t indices[16];
    va_list ap;
    va_start(ap, arr);
    for (int32_t d = 0; d < arr->dimCount; d++) {
        indices[d] = va_arg(ap, int32_t);
    }
    va_end(ap);

    int32_t offset = vb6_SafeArrayND_Offset(arr, arr->dimCount, indices);
    if (offset < 0 || offset >= arr->totalElements) return NULL;
    return (char*)arr->data + offset * arr->elemSize;
}

vb6_SafeArrayND* vb6_SafeArrayReDimND(vb6_safearray_elemtype elemType,
    int32_t dimCount, vb6_SafeArrayBound bounds[]) {
    return vb6_SafeArrayCreateND(elemType, dimCount, bounds);
}

vb6_SafeArrayND* vb6_SafeArrayReDimND_Udt(int32_t elemSize,
    int32_t dimCount, vb6_SafeArrayBound bounds[]) {
    if (dimCount <= 0 || dimCount > 16) return NULL;

    vb6_SafeArrayND* arr = (vb6_SafeArrayND*)calloc(1, sizeof(vb6_SafeArrayND));
    if (!arr) return NULL;

    arr->dimCount = dimCount;
    arr->elemType = vb6_sa_udt;
    arr->elemSize = elemSize;

    int32_t total = 1;
    for (int32_t d = 0; d < dimCount; d++) {
        arr->bounds[d] = bounds[d];
        if (bounds[d].cElements <= 0) {
            arr->totalElements = 0;
            arr->data = NULL;
            return arr;
        }
        total *= bounds[d].cElements;
    }
    arr->totalElements = total;

    if (total > 0) {
        arr->data = calloc((size_t)total, (size_t)arr->elemSize);
        if (!arr->data) {
            free(arr);
            return NULL;
        }
    }

    return arr;
}

vb6_SafeArrayND* vb6_SafeArrayReDimPreserveND(vb6_SafeArrayND* arr,
    int32_t dimCount, vb6_SafeArrayBound newBounds[]) {
    if (!arr) return vb6_SafeArrayCreateND(vb6_sa_empty, dimCount, newBounds);

    int32_t newTotal = 1;
    for (int32_t d = 0; d < dimCount; d++) {
        if (newBounds[d].cElements <= 0) {
            vb6_SafeArrayDestroyND(arr);
            return NULL;
        }
        newTotal *= newBounds[d].cElements;
    }

    void* newData = calloc((size_t)newTotal, (size_t)arr->elemSize);
    if (!newData) return arr;

    if (arr->data && arr->totalElements > 0) {
        int32_t minDims = (dimCount < arr->dimCount) ? dimCount : arr->dimCount;

        int32_t copyCounts[16];
        int32_t oldCounts[16];
        int32_t newCounts[16];
        int32_t oldStrides[16];
        int32_t newStrides[16];

        for (int32_t d = 0; d < dimCount; d++)
            newCounts[d] = newBounds[d].cElements;
        for (int32_t d = 0; d < minDims; d++)
            oldCounts[d] = arr->bounds[d].cElements;
        for (int32_t d = minDims; d < 16; d++)
            oldCounts[d] = 0;

        for (int32_t d = 0; d < dimCount; d++)
            copyCounts[d] = (oldCounts[d] < newCounts[d]) ? oldCounts[d] : newCounts[d];

        oldStrides[minDims - 1] = 1;
        for (int32_t d = minDims - 2; d >= 0; d--)
            oldStrides[d] = oldStrides[d + 1] * arr->bounds[d + 1].cElements;

        newStrides[dimCount - 1] = 1;
        for (int32_t d = dimCount - 2; d >= 0; d--)
            newStrides[d] = newStrides[d + 1] * newBounds[d + 1].cElements;

        int32_t iterMax = 1;
        for (int32_t d = 0; d < minDims; d++)
            iterMax *= copyCounts[d];

        for (int32_t linear = 0; linear < iterMax; linear++) {
            int32_t tmp = linear;
            int32_t oldOff = 0, newOff = 0;
            int32_t bounds_check = 1;
            for (int32_t d = minDims - 1; d >= 0; d--) {
                int32_t idx = tmp % copyCounts[d];
                tmp /= copyCounts[d];
                if (idx >= oldCounts[d] || idx >= newCounts[d]) {
                    bounds_check = 0;
                    break;
                }
                oldOff += idx * oldStrides[d];
                newOff += idx * newStrides[d];
            }
            if (bounds_check) {
                memcpy((char*)newData + newOff * arr->elemSize,
                       (char*)arr->data + oldOff * arr->elemSize,
                       (size_t)arr->elemSize);
            }
        }
    }

    // P1-12修复: BSTR/VARIANT - 旧区域中被丢弃的元素需要释放
    if (arr->data && (arr->elemType == vb6_sa_bstr || arr->elemType == vb6_sa_variant)) {
        for (int32_t linear = 0; linear < arr->totalElements; linear++) {
            int32_t tmp = linear;
            int32_t indices[16];
            int32_t d;
            for (d = arr->dimCount - 1; d >= 0; d--) {
                indices[d] = tmp % arr->bounds[d].cElements;
                tmp /= arr->bounds[d].cElements;
            }
            int32_t inNewBounds = 1;
            for (d = 0; d < arr->dimCount && d < dimCount; d++) {
                if (indices[d] >= newBounds[d].cElements) { inNewBounds = 0; break; }
            }
            if (arr->dimCount > dimCount) inNewBounds = 0;
            if (!inNewBounds) {
                if (arr->elemType == vb6_sa_bstr) {
                    BSTR* slot = (BSTR*)((char*)arr->data + linear * arr->elemSize);
                    if (*slot) vb6_BSTR_Free(*slot);
                } else if (arr->elemType == vb6_sa_variant) {
                    vb6_VARIANT* slot = (vb6_VARIANT*)((char*)arr->data + linear * arr->elemSize);
                    vb6_VariantClear(slot);
                }
            }
        }
    }

    if (arr->data) free(arr->data);
    arr->data = newData;
    arr->dimCount = dimCount;
    arr->totalElements = newTotal;
    for (int32_t d = 0; d < dimCount; d++)
        arr->bounds[d] = newBounds[d];
    for (int32_t d = dimCount; d < 16; d++) {
        arr->bounds[d].lBound = 0;
        arr->bounds[d].cElements = 0;
    }

    return arr;
}

// UDT版ReDim Preserve(多维): 元素尺寸由调用方提供 (sizeof(UDT))。
// NULL 初值不再回落 vb6_sa_empty(4字节), 避免 UDT 多维数组写入越界。
vb6_SafeArrayND* vb6_SafeArrayReDimPreserveND_Udt(int32_t elemSize,
    vb6_SafeArrayND* arr, int32_t dimCount, vb6_SafeArrayBound newBounds[]) {
    if (!arr) return vb6_SafeArrayReDimND_Udt(elemSize, dimCount, newBounds);

    if (arr->elemSize != elemSize) {
        vb6_SafeArrayDestroyND(arr);
        return vb6_SafeArrayReDimND_Udt(elemSize, dimCount, newBounds);
    }

    int32_t newTotal = 1;
    for (int32_t d = 0; d < dimCount; d++) {
        if (newBounds[d].cElements <= 0) {
            vb6_SafeArrayDestroyND(arr);
            return NULL;
        }
        newTotal *= newBounds[d].cElements;
    }

    void* newData = calloc((size_t)newTotal, (size_t)elemSize);
    if (!newData) return arr;

    if (arr->data && arr->totalElements > 0) {
        int32_t minDims = (dimCount < arr->dimCount) ? dimCount : arr->dimCount;

        int32_t copyCounts[16];
        int32_t oldCounts[16];
        int32_t newCounts[16];
        int32_t oldStrides[16];
        int32_t newStrides[16];

        for (int32_t d = 0; d < dimCount; d++)
            newCounts[d] = newBounds[d].cElements;
        for (int32_t d = 0; d < minDims; d++)
            oldCounts[d] = arr->bounds[d].cElements;
        for (int32_t d = minDims; d < 16; d++)
            oldCounts[d] = 0;

        for (int32_t d = 0; d < dimCount; d++)
            copyCounts[d] = (oldCounts[d] < newCounts[d]) ? oldCounts[d] : newCounts[d];

        oldStrides[minDims - 1] = 1;
        for (int32_t d = minDims - 2; d >= 0; d--)
            oldStrides[d] = oldStrides[d + 1] * arr->bounds[d + 1].cElements;

        newStrides[dimCount - 1] = 1;
        for (int32_t d = dimCount - 2; d >= 0; d--)
            newStrides[d] = newStrides[d + 1] * newBounds[d + 1].cElements;

        int32_t iterMax = 1;
        for (int32_t d = 0; d < minDims; d++)
            iterMax *= copyCounts[d];

        for (int32_t linear = 0; linear < iterMax; linear++) {
            int32_t tmp = linear;
            int32_t oldOff = 0, newOff = 0;
            int32_t bounds_check = 1;
            for (int32_t d = minDims - 1; d >= 0; d--) {
                int32_t idx = tmp % copyCounts[d];
                tmp /= copyCounts[d];
                if (idx >= oldCounts[d] || idx >= newCounts[d]) {
                    bounds_check = 0;
                    break;
                }
                oldOff += idx * oldStrides[d];
                newOff += idx * newStrides[d];
            }
            if (bounds_check) {
                memcpy((char*)newData + newOff * elemSize,
                       (char*)arr->data + oldOff * elemSize,
                       (size_t)elemSize);
            }
        }
    }

    if (arr->data) free(arr->data);
    arr->data = newData;
    arr->dimCount = dimCount;
    arr->elemSize = elemSize;
    arr->totalElements = newTotal;
    for (int32_t d = 0; d < dimCount; d++)
        arr->bounds[d] = newBounds[d];
    for (int32_t d = dimCount; d < 16; d++) {
        arr->bounds[d].lBound = 0;
        arr->bounds[d].cElements = 0;
    }

    return arr;
}

int32_t vb6_UBoundND(vb6_SafeArrayND* arr, int32_t dimension) {
    if (!arr || dimension < 1 || dimension > arr->dimCount) return 0;
    return arr->bounds[dimension - 1].lBound + arr->bounds[dimension - 1].cElements - 1;
}

int32_t vb6_LBoundND(vb6_SafeArrayND* arr, int32_t dimension) {
    if (!arr || dimension < 1 || dimension > arr->dimCount) return 0;
    return arr->bounds[dimension - 1].lBound;
}

