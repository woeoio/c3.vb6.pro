#pragma once
// vb6rtl_array.h - SAFEARRAY（一维 / 多维）与错误状态栈
// 由 vb6rtl.h 伞头 include；生成代码不要直接 include 本文件
#include "vb6rtl_base.h"

#ifdef _WIN32
#include <stdio.h>
#include <stdlib.h>
#include <intrin.h>   /* 越界探针: _ReturnAddress 取发码调用点 */
#endif

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// SAFEARRAY - VB6 动态/静态数组
// ============================================================

// 元素类型标识 (用于ReDim/Erase时正确清理)
typedef enum vb6_safearray_elemtype {
    vb6_sa_empty = 0,
    vb6_sa_bool  = 1,   // int16_t
    vb6_sa_byte  = 2,   // uint8_t
    vb6_sa_int   = 3,   // int16_t
    vb6_sa_long  = 4,   // int32_t
    vb6_sa_single= 5,   // float
    vb6_sa_double= 6,   // double
    vb6_sa_bstr  = 7,   // BSTR (需要逐元素释放)
    vb6_sa_variant=8,   // vb6_VARIANT (需要逐元素清理)
    vb6_sa_ptr   = 9,   // void* (对象引用)
    vb6_sa_currency = 10, // int64_t (VB6 Currency: value*10000)
    vb6_sa_udt    = 11,   // UDT (elemSize set externally)
} vb6_safearray_elemtype;

// 一维数组描述符 (VB6绝大多数用例是一维)
//
// ⚠ **前 24 字节必须与 oleauto.h 的 SAFEARRAY 逐字节同布局** —— 这一条不是洁癖，
// 是 VB6 生态的硬需求。VB6 里从 Variant 掏数组的地址是这么写的：
//
//     CopyMemory Ptr, ByVal UnsignedAdd(VarPtr(SomeVariant), 8), PTR_SIZE  ' union 在偏移 8
//     If Ptr <> NULL_PTR Then
//         CopyMemory DimensionCount, ByVal Ptr, 2                          ' 当 SAFEARRAY* 读首 WORD
//         If DimensionCount = 1 Then ...                                   '  = .cDims
//
// 也就是**把 `Variant.parray` 直接当 `SAFEARRAY*` 用**。实测工程 ComCtlsDemo 里
// **7 处**这么写：CommonDialog.CustomColors / ListView.ColumnOrder /
// ListView.FListItemTileViewIndices / FontCombo.RestoreRecent /
// SpinBox.SetAcceleration(×2) / Common.IsArrayInitialized。
//
// 此前首字段是 `int32_t signature = 0x5A1D`，于是 `.cDims` 读出 23197 ⇒ 恒 ≠ 1 ⇒
// 恒抛**错误 5** ("Array must be single dimensioned")。ComCtlsDemo 的 Form_Load 就死在
// `ListView1.ListItems(1).TileViewIndices = Array(1, 2)` 这一句上（第十六个卡点）。
//
// x86 上 SAFEARRAY 的实际偏移：0 cDims(W) / 2 fFeatures(W) / 4 cbElements(D) /
// 8 cLocks(D) / 12 pvData(P) / 16 rgsabound[0].cElements(D) / 20 rgsabound[0].lLbound(D)。
// 下面逐槽对上 (总尺寸仍是 32 字节，与改动前一致)。
typedef struct vb6_SafeArray1D {
    /* ===== OLE SAFEARRAY 兼容前缀 (偏移与字段语义都对得上) ===== */
    uint16_t cDims;                     /*  0  [.cDims]     1 维数组恒为 1 */
    uint16_t fFeatures;                 /*  2  [.fFeatures] C3 标记 VB6_SA1D_FFEATURES */
    int32_t  elemSize;                  /*  4  [.cbElements] 单个元素字节数 */
    int32_t  elemType;                  /*  8  [.cLocks 槽]  ⚠ C3 扩展: 元素类型枚举 */
    void*    data;                      /* 12  [.pvData]    数据指针 (calloc分配, 零初始化) */
    int32_t  count;                     /* 16  [rgsabound[0].cElements] 元素个数 */
    int32_t  lBound;                    /* 20  [rgsabound[0].lLbound]   下界 */
    /* ===== 以下为 C3 内部扩展 (OLE 侧看不到) ===== */
    int32_t  uBound;                    /* 24  冗余: = lBound + count - 1 */
    int32_t  isDynamic;                 /* 28  是否动态数组 (ReDim创建) */
} vb6_SafeArray1D;

// 1D 描述符的 fFeatures 标记。真实 SAFEARRAY 的 fFeatures 是 FADF_* 位集，这里借它
// 放一枚 C3 标记：既让 `vb6_SaIs1D` 能把 1D 描述符与 `vb6_SafeArrayND`
// (首字段 dimCount = 2..16，offset 2 恒 0) 区分开，又让 `.fFeatures` 保持非零。
#define VB6_SA1D_FFEATURES 0xA1D1u

// ⚠ 注意 offset 8 的 `elemType` 占的是 SAFEARRAY.cLocks 的位置。oleaut32 的
// SafeArrayLock/Unlock 会把 cLocks 当计数器加减 —— 对 C3 描述符调用它就会篡改
// elemType (进而影响销毁时的逐元素释放路径)。C3 自己的数组只经本 RTL 流转，
// 从不交给 oleaut32 的 SafeArray* API，所以这一槽冲突不成立；真要交给它们时，
// 必须先深拷成真 SAFEARRAY (见 vb6_VariantToSafeArray1D 的调用面)。
//
// 这几条断言把"与 SAFEARRAY 同布局"变成编译期事实：谁重排/插入字段都会当场编不过。
// x64 的偏移与 x86 不同 —— `void* pvData` 要 8 字节对齐，于是编译器在 elemType(cLocks 槽)
// 后插 4 字节填充：pvData 落 16、cElements 落 24、lLbound 落 28，struct 长到 40。
// 这**正好**也是 oleaut32 在 x64 上 SAFEARRAY 的布局，所以两档都自然对齐，无需手动填充。
_Static_assert(offsetof(vb6_SafeArray1D, cDims)    == 0,  "SAFEARRAY.cDims 槽");
_Static_assert(offsetof(vb6_SafeArray1D, fFeatures) == 2,  "SAFEARRAY.fFeatures 槽");
_Static_assert(offsetof(vb6_SafeArray1D, elemSize) == 4,  "SAFEARRAY.cbElements 槽");
#ifdef _WIN64
_Static_assert(sizeof(vb6_SafeArray1D) == 40, "vb6_SafeArray1D must stay 40 bytes (x64 SAFEARRAY prefix)");
_Static_assert(offsetof(vb6_SafeArray1D, data)   == 16, "SAFEARRAY.pvData 槽 (x64)");
_Static_assert(offsetof(vb6_SafeArray1D, count)  == 24, "SAFEARRAY.rgsabound[0].cElements 槽 (x64)");
_Static_assert(offsetof(vb6_SafeArray1D, lBound) == 28, "SAFEARRAY.rgsabound[0].lLbound 槽 (x64)");
#else
_Static_assert(sizeof(vb6_SafeArray1D) == 32, "vb6_SafeArray1D must stay 32 bytes (x86 SAFEARRAY prefix)");
_Static_assert(offsetof(vb6_SafeArray1D, data)   == 12, "SAFEARRAY.pvData 槽 (x86)");
_Static_assert(offsetof(vb6_SafeArray1D, count)  == 16, "SAFEARRAY.rgsabound[0].cElements 槽 (x86)");
_Static_assert(offsetof(vb6_SafeArray1D, lBound) == 20, "SAFEARRAY.rgsabound[0].lLbound 槽 (x86)");
#endif

// 1D 描述符初始化 (所有构造点共用：Create1D / ReDim1D_Udt / ArrayAssign1D /
// com 的两处 / codegen 的复合字面量)。
static inline void vb6_Sa1DInit(vb6_SafeArray1D* arr) {
    arr->cDims = 1;
    arr->fFeatures = (uint16_t)VB6_SA1D_FFEATURES;
}

// 这是不是 C3 的 1D 描述符？(ND 描述符首 WORD 是 dimCount = 2..16，offset 2 恒 0)
static inline int vb6_SaIs1D(const void* p) {
    const vb6_SafeArray1D* a = (const vb6_SafeArray1D*)p;
    return a != NULL && a->cDims == 1 && a->fFeatures == (uint16_t)VB6_SA1D_FFEATURES;
}

// P14.2.3: 字符串数组函数
vb6_SafeArray1D* vb6_Split(BSTR expr, BSTR delimiter, int32_t limit, int32_t compare);
BSTR vb6_Join(vb6_SafeArray1D* arr, BSTR delimiter);

// 创建一维静态数组
vb6_SafeArray1D* vb6_SafeArrayCreate1D(vb6_safearray_elemtype elemType,
    int32_t lBound, int32_t uBound);

// 创建一维动态数组 (ReDim)
vb6_SafeArray1D* vb6_SafeArrayReDim1D(vb6_safearray_elemtype elemType,
    int32_t lBound, int32_t uBound);

// UDT版ReDim: elemSize由调用方提供 (用于UDT动态数组)
vb6_SafeArray1D* vb6_SafeArrayReDim1D_Udt(int32_t elemSize,
    int32_t lBound, int32_t uBound);

// ReDim Preserve: 保留原有数据, 调整大小。
// ⚠ elemType 形参不是冗余 —— 它决定**元素步长**, 而步长是运行期唯一的真相来源。
// 历史坑（三处同源，MEMORY §15/§16）: 早期签名没有 elemType, arr==NULL 时只能回落
// vb6_sa_variant(x86 16 字节)。于是 `Dim m_Keys() As String` 首次 ReDim 按 16 字节/元素
// 分配, 而 `VB6_SA_AT(BSTR, arr, i)` 按 4 字节步进读写 → 第 i 个槽里其实躺着别的元素的
// BSTR; 销毁时 vb6_SafeArrayDestroy1D 又按 elemType==variant 走 Variant 分支, 把那些
// BSTR 当 bstrVal **二次 free** → 堆损坏, 且崩点漂移到"后面某次 free" (实测 play78.exe
// 的 List.cls: `For i = 0 To UBound(m_Keys)` 读出垃圾 BSTR 后 StrComp 崩在 mov eax,[edx+ecx*4])。
// 现在 elemType 由 codegen 权威给出 (mapSaElemType), 运行时不再猜。
// 保留三参旧入口: 仍有走不到的调用点, 且 elemType 未知时它按 variant 兜底 (=旧行为)。
vb6_SafeArray1D* vb6_SafeArrayReDimPreserve1D_T(vb6_safearray_elemtype elemType,
    vb6_SafeArray1D* arr, int32_t newLBound, int32_t newUBound);
vb6_SafeArray1D* vb6_SafeArrayReDimPreserve1D(vb6_SafeArray1D* arr,
    int32_t newLBound, int32_t newUBound);
// 账 #305 §B218 (甲) 腿⑤ 第四格 (第一百四十八刀): 上面那位主人的**带回调兄弟** —— 元素是工程类
// 裸实例 (vb6_sa_ptr) 时, 收缩丢尾的那几格要在搬移之前把引用交回。为什么多收一枚函数指针:
// 裸结构体没有 vtable、每类的 Release 是各自一对出口 ⇒ RTL 自己认不出该调谁 (与
// vb6_SA_ReleaseElements 同一口径)。搬移本身**不重复实现**: 这枚只做完出厂就转调 _T,
// 所以两条路的保留/重建规则永远一字相同。非类元素请继续直接用 _T (发码形状逐字节不变)。
vb6_SafeArray1D* vb6_SafeArrayReDimPreserve1D_Tr(vb6_safearray_elemtype elemType,
    vb6_SafeArray1D* arr, int32_t newLBound, int32_t newUBound,
    unsigned long (*release)(void*));

// UDT版ReDim Preserve: elemSize由调用方提供 (用于UDT动态数组)
vb6_SafeArray1D* vb6_SafeArrayReDimPreserve1D_Udt(int32_t elemSize,
    vb6_SafeArray1D* arr, int32_t newLBound, int32_t newUBound);

// 销毁数组 (释放内存)
void vb6_SafeArrayDestroy1D(vb6_SafeArray1D* arr);
// 账 #305 §B218 (甲) 腿⑤ 数组档 (第一百四十刀): 元素是**工程类实例指针** (vb6_sa_ptr) 时,
// 逐元素把引用交回给持有它的那位主人 —— 谁的 Release 只有调用方知道 (每类一对出口),
// 所以这里收函数指针。elemType 不是 vb6_sa_ptr 或 release 为空 ⇒ 一个字都不做 (=旧行为)。
// 必须在销毁载体**之前**调用: 载体一没, 元素地址就没了。
void vb6_SA_ReleaseElements(void* arr, unsigned long (*release)(void*));
// Fix <vbeclipse> 2026-10-06: Erase 语句的数组变量出口 (VB6 文档口径):
// 动态数组 → 释放存储并置 NULL; 固定数组 → 保留存储, 元素重置
// (数值/UDT 置零, BSTR 置零长串, Variant 置 Empty)。
void vb6_EraseArrayVar(vb6_SafeArray1D** parr);

// Fix 170: 整体数组赋值 `A() = B()` —— 深拷贝 src 为新载体返回, 原 dst 销毁
vb6_SafeArray1D* vb6_ArrayAssign1D(vb6_SafeArray1D* dst, vb6_SafeArray1D* src);

// Fix 178: 元素是「含所有权成员的 UDT」(vb6_sa_udt) 时, memcpy 之后每个槽位仍指向
// src 的元素 → 各行别名到同一份存储。cb(dstElem, srcElem) 负责把 src 元素深拷贝进
// dst 元素 (dst 元素此刻是 src 的位拷贝, 其中指针归 src 所有, 不得释放)。
// cb==NULL 等价于 Fix 170 的原行为。
typedef void (*vb6_udt_elem_copy)(void* dst, const void* src);
vb6_SafeArray1D* vb6_ArrayAssign1D_Cb(vb6_SafeArray1D* dst, vb6_SafeArray1D* src,
                                      vb6_udt_elem_copy cb);

// 获取/设置元素 (void*通用版)
void* vb6_SafeArrayGetPtr(vb6_SafeArray1D* arr, int32_t index);
void  vb6_SafeArrayPutElem(vb6_SafeArray1D* arr, int32_t index, void* value);

// 便捷宏: 按类型访问元素
// Fix 106: VB6 允许用非整型表达式做数组下标 (如 Dim i As Single 后 arr(i)),
// 由 VB6 隐式转换为 Long。这里把下标差值显式转 int32_t, 否则 C 端
// float/double 下标直接报 C2108 (Charts 2020 ucChartArea: 202 处),
// 且级联出 C2198 等二次错误。Variant 下标仍由生成端 toLongIfVariant 处理。
//
// 账 #209: 这里是**唯一**对一维数组描述符的裸解引用点, 此前它一个检查都没有, 于是
//   `With m_Serie(Index)`（动态数组从未 ReDim / 已被 Erase）读 (NULL)->lBound
//   ⇒ 0xC0000005 原生崩（实测 ucChartBar demo 点 Random, 三次同偏移 0x1abf2）。
// VB6 在这一条是**运行时错误 9**, 与 vb6_UBound/vb6_LBound 的 rev2 同族 —— 那两处
// 已经抛 9, 本刀补上"元素"这一半。步长仍按调用方写明的 sizeof(type), 不改读描述符的
// elemSize —— 那是 Fix 170/rev3 记下的独立历史坑, 本刀只加检查、不动步长语义。
void vb6_SaElemFail(void* arr, int32_t idx);  // 必抛 9, 不返回
static inline void* vb6_SaElemPtr(void* arrV, int32_t idx, int32_t elemSize) {
    vb6_SafeArray1D* arr = (vb6_SafeArray1D*)arrV;
    if (!arr || idx < arr->lBound || idx > arr->uBound) {
#ifdef _WIN32
        if (getenv("C3_SA_TRACE")) {
            /* 定位用: 这里打的是**发码调用点**的绝对返回地址 (本函数在 Debug 下是独立
               函数, 未内联), 与 vb6_SaElemFail 里打的 base 相减即得可符号化的 RVA。 */
            fprintf(stderr, "[SA]   site ret=%p\n", _ReturnAddress());
            fflush(stderr);
        }
#endif
        vb6_SaElemFail(arrV, idx);
    }
    return (char*)arr->data + (ptrdiff_t)(idx - arr->lBound) * (ptrdiff_t)elemSize;
}
#define VB6_SA_AT(type, arr, idx) \
    (*(type*)vb6_SaElemPtr((arr), (int32_t)(idx), (int32_t)sizeof(type)))

// UBound/LBound (替换旧stub)
int32_t vb6_UBound(vb6_SafeArray1D* safeArray, int32_t dimension);
int32_t vb6_LBound(vb6_SafeArray1D* safeArray, int32_t dimension);

// P21-B: Array() function support
vb6_SafeArray1D* vb6_ArrayCreate(int32_t count);
void vb6_ArraySetLong(vb6_SafeArray1D* arr, int32_t index, int32_t val);
void vb6_ArraySetDouble(vb6_SafeArray1D* arr, int32_t index, double val);
void vb6_ArraySetBSTR(vb6_SafeArray1D* arr, int32_t index, BSTR val);
void vb6_ArraySetVariant(vb6_SafeArray1D* arr, int32_t index, vb6_VARIANT val);
// ============================================================
// SAFEARRAY ND - VB6 多维数组
// ============================================================

typedef struct vb6_SafeArrayBound {
    int32_t lBound;
    int32_t cElements;
} vb6_SafeArrayBound;

typedef struct vb6_SafeArrayND {
    int32_t dimCount;
    vb6_safearray_elemtype elemType;
    int32_t elemSize;
    int32_t totalElements;
    void* data;
    vb6_SafeArrayBound bounds[16];
} vb6_SafeArrayND;

vb6_SafeArrayND* vb6_SafeArrayCreateND(vb6_safearray_elemtype elemType,
    int32_t dimCount, vb6_SafeArrayBound bounds[]);
void vb6_SafeArrayDestroyND(vb6_SafeArrayND* arr);
// 账 #305 §B218 (甲) 腿⑤ 第五格 (第一百四十九刀): 上面 1D 那位 vb6_SA_ReleaseElements 的**多维兄弟** ——
// 元素是本工程类裸实例 (vb6_sa_ptr) 的 ND 载体在销毁前逐枚交回。步进口径与 vb6_SafeArrayDestroyND 里
// bstr / variant 那两档一字相同 (按 i * elemSize 线性走 totalElements 格), 区别只在「这一格是裸实例指针」
// ⇒ 该调谁的 Release 仍只有调用方知道。release 为空 / elemType 不是 vb6_sa_ptr / 载体还没建 ⇒ 一个字不做。
// 同样必须在销毁载体**之前**调用。
void vb6_SA_ReleaseElementsND(void* arr, unsigned long (*release)(void*));
vb6_SafeArrayND* vb6_SafeArrayReDimND(vb6_safearray_elemtype elemType,
    int32_t dimCount, vb6_SafeArrayBound bounds[]);
vb6_SafeArrayND* vb6_SafeArrayReDimND_Udt(int32_t elemSize,
    int32_t dimCount, vb6_SafeArrayBound bounds[]);
/* 账 #305 §B155 格④ (第四十三刀): 这一枚也**收 elemType**, 与上面一维那条 _T 同一个决定。
   此前形参表里没有它, arr==NULL (首次分配就写 Preserve) 时只能按 vb6_sa_empty 建 ——
   尺寸表对 empty 给 default 档 4 字节, 而访问侧按声明档步进:
     `Dim p() As String` + `ReDim Preserve p(1 To 2, 1 To 3)` 在 x64 上造 4 字节一步、
     写 8 字节一步 ⇒ 堆越界 (实测旧编译器 0xC0000374 堆损坏、新编译器 0xC0000005,
     两条都是响的; Variant 档 24/4 同形)。
   UDT 不在这一枚 (它走下面的 _Udt, 元素尺寸由调用方的 sizeof 给)。 */
vb6_SafeArrayND* vb6_SafeArrayReDimPreserveND(vb6_safearray_elemtype elemType,
    vb6_SafeArrayND* arr, int32_t dimCount, vb6_SafeArrayBound newBounds[]);
/* 账 #305 §B218 (甲) 腿⑤ 第六格 (第一百五十四刀): 本体多收一枚元素交回回调, 上面那枚是它的转调。
   与一维那位兄弟 vb6_SafeArrayReDimPreserve1D_Tr 同一个理由 —— 收缩丢尾那几格躺的是本工程
   哪一类的实例, 只有 cgen 知道该递哪一枚 <Cls>_Release。release 为空 = 今天那一发留着。 */
vb6_SafeArrayND* vb6_SafeArrayReDimPreserveND_Tr(vb6_safearray_elemtype elemType,
    vb6_SafeArrayND* arr, int32_t dimCount, vb6_SafeArrayBound newBounds[],
    unsigned long (*release)(void*));

// UDT版ReDim Preserve(多维): elemSize由调用方提供
vb6_SafeArrayND* vb6_SafeArrayReDimPreserveND_Udt(int32_t elemSize,
    vb6_SafeArrayND* arr, int32_t dimCount, vb6_SafeArrayBound newBounds[]);

int32_t vb6_SafeArrayND_Offset(vb6_SafeArrayND* arr, int32_t dimCount, int32_t indices[]);
void* vb6_SafeArrayND_GetPtr(vb6_SafeArrayND* arr, ...);
// 账 #305 §B150: elemType → 单元素字节数 (UDT 档不在这里答, 由描述符自己的 elemSize 说话)。
int32_t vb6_SaElemSize(vb6_safearray_elemtype t);

// 账 #305 §B155 格①（第四十六刀）: elemType → VB6 的 VarType **元素**码（vbArray 那一位由调用方 OR 上）。
// 与上面那张尺寸表同一族 —— 「这枚数组是哪一档」只许一处答。认不出的档（empty / udt）回落 variant：
// 那就是改前每一档的答案。Date 档在这枚枚举里根本没有（数组里的 Date 存成 double），那一格记在 §B159。
int32_t vb6_SaVtFromElemType(vb6_safearray_elemtype t);

// 账 #305 §B159（第五十刀）：与上面那张表并排住的**反向**表 —— Windows 的 VARENUM → 我们的元素枚举。
// 「哪一档对应哪一档」第二次有了唯一答案：改前 vb6rtl_com.c 里两条进口各抄了一份元素 switch，
// 两份都只认 6 档、默认回落 variant；于是 COM 交回一枚 Byte 数组时按 24 字节步长 memcpy 一枚
// 1 字节/元素的源 = **越界读**，而标签答 8204/"Variant()"。认不出的档照旧回落 variant
//（= 改前那两条的答案，行为不动），元素宽度一律交回上面那张尺寸表答。
vb6_safearray_elemtype vb6_SaElemTypeFromVt(VARTYPE comVt);

int32_t vb6_UBoundND(vb6_SafeArrayND* arr, int32_t dimension);
int32_t vb6_LBoundND(vb6_SafeArrayND* arr, int32_t dimension);

// Fix <vbeclipse>: UDT (Type ... End Type) 作为**值**传给 COM 方法时的编组。
// VB6 把 UDT 打成 SAFEARRAY(VT_UI1) —— 对端收到一串字节, 长度 = UDT 字节数。
// 缺这个入口时 C3 落 vb6_ComPackInt(<struct>) → C2440 "无法从 vb6_type_RECT 转换为
// int32_t" (ucSplitBar.ctl:149 `.SplitterMouseDown UserControl.hWnd, m_Rect, x, y`
// —— m_Rect As RECT, 4×int32_t)。
//
// ⚠ 深拷贝而非存指针: 传址会让调用方后续改 UDT 时对端看到变化, 而 VB6 的
// ByVal UDT 传参是值语义 (对端拿到调用瞬间的快照)。内存由 VARIANT clear 路径
// 按 vb6_sa_byte 逐元素释放, 不需要额外钩子。
//
// ⚠ 这个 inline 放在本文件而不是 vb6rtl_variant.h: 伞头里 array.h 排在
// variant.h **之后**, 只有这里才看得见 vb6_SafeArrayCreate1D / vb6_SafeArray1D。
// 放进 variant.h 会 C2021/C4018 (类型未声明)。
static inline vb6_VARIANT vb6_VariantFromUdtBytes(const void* src, int32_t nBytes) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    if (!src || nBytes <= 0) { v.vt = vb6_vtEmpty; return v; }
    vb6_SafeArray1D* arr = vb6_SafeArrayCreate1D(vb6_sa_byte, 0, nBytes - 1);
    if (!arr) { v.vt = vb6_vtEmpty; return v; }
    memcpy(arr->data, src, (size_t)nBytes);
    v.vt = (vb6_vartype)(vb6_vtArray | vb6_SaVtFromElemType(vb6_sa_byte));   /* VT_ARRAY|VT_UI1，问表不写死 */
    v.parray = arr;
    return v;
}

// Fix 106: 同 VB6_SA_AT, 多维下标也显式转 int32_t (VB6 隐式 CLng)。
// 账 #209 同族的下一半: 这一支以前和当初的 VB6_SA_AT 一模一样 —— 描述符在不在、
// 每一维在不在范围内, 一个都不问。三条实测读数:
//   · `Dim a2(1 To 2, 1 To 3)` 读 a2(3,1) ⇒ 静默拿到隔壁那格 (12)，VB6 是错误 9；
//   · `Dim d() As Long` 从没 ReDim 就取 d(1,1) ⇒ 读 (NULL)->data ⇒ 0xC0000005；
//   · 4 秩数组 `Dim a4(1 To 2,1 To 2,1 To 2,1 To 2)` **在范围内也崩** —— 那是发码侧
//     把实参写成 `(int[]){i0, i1}` 只塞了两个下标却按实际秩数交出去 (cgen_expr_call_prelude.inc)。
// 现在全部秩数(1..16)都走这一个出口: 热路径逐维比上下界, 冷路径抛 9。步长仍按调用方
// 写明的 sizeof(type) —— 与一维那一刀同一口径, 不去读描述符的 elemSize。
void vb6_SaNdElemFail(void* arr, const int32_t* idx, int32_t rank);   // 必抛 9, 不返回
static inline void* vb6_SaNdElemPtr(void* arrV, int32_t rank, const int32_t* idx,
                                    int32_t elemSize) {
    vb6_SafeArrayND* a = (vb6_SafeArrayND*)arrV;
    // 一维描述符被按多维形状用 (Fix 056 那条"声明 1D、ReDim 成 ND"的双面形) ——
    // 1D 描述符的首 WORD 现在是 SAFEARRAY.cDims == 1 (与 OLE 同布局, 见上面 struct 注释),
    // 而 ND 的 dimCount 只可能是 2..16 (发码侧 dimCount==1 必走 1D 路径), 所以下界取 2
    // 就同时认出"这不是 ND"且不必多读字段。
    if (!a || a->dimCount < 2 || a->dimCount > 16 || rank < 1 || rank > a->dimCount) {
        vb6_SaNdElemFail(arrV, idx, rank);
    }
    ptrdiff_t off = 0;
    ptrdiff_t stride = 1;
    for (int32_t d = 0; d < rank; d++) {
        int32_t lb = a->bounds[d].lBound;
        int32_t cnt = a->bounds[d].cElements;
        if (idx[d] < lb || idx[d] >= lb + cnt) vb6_SaNdElemFail(arrV, idx, rank);
        off += (ptrdiff_t)(idx[d] - lb) * stride;
        stride *= cnt;
    }
    return (char*)a->data + off * (ptrdiff_t)elemSize;
}

#define VB6_SA_ND_AT1(elemType, arr, i) \
    (*(elemType*)vb6_SaNdElemPtr((arr), 1, \
        (const int32_t[]){ (int32_t)(i) }, (int32_t)sizeof(elemType)))

#define VB6_SA_ND_AT2(elemType, arr, i, j) \
    (*(elemType*)vb6_SaNdElemPtr((arr), 2, \
        (const int32_t[]){ (int32_t)(i), (int32_t)(j) }, (int32_t)sizeof(elemType)))

#define VB6_SA_ND_AT3(elemType, arr, i, j, k) \
    (*(elemType*)vb6_SaNdElemPtr((arr), 3, \
        (const int32_t[]){ (int32_t)(i), (int32_t)(j), (int32_t)(k) }, (int32_t)sizeof(elemType)))

// 4 秩及以上(VB6 到 60 秩)走这一条: 秩数与下标数组都由发码侧一次交全。
#define VB6_SA_ND_ATN(elemType, arr, rank, idx) \
    (*(elemType*)vb6_SaNdElemPtr((arr), (rank), (idx), (int32_t)sizeof(elemType)))

// 文件 I/O
int32_t vb6_FreeFile(void);
int32_t vb6_Open(BSTR pathname, int32_t mode, int32_t access, int32_t filenumber, int32_t reclength);
int32_t vb6_Close(int32_t filenumber);
int32_t vb6_CloseAll();  // Close all open files
int32_t vb6_EOF(int32_t filenumber);
int32_t vb6_LOF(int32_t filenumber);
int32_t vb6_Loc(int32_t filenumber);// P21-13: Seek function (return current file position)
int32_t vb6_SeekFunc(int32_t filenumber);
// P21-13: Seek statement (set file position)
void vb6_SeekStmt(int32_t filenumber, int32_t position);

// P22-08: Width# — set file output line width
void vb6_Width(int32_t filenumber, int32_t width);

void vb6_Print(int32_t filenumber, BSTR s);
void vb6_Write(int32_t filenumber, BSTR s);
BSTR vb6_LineInput(int32_t filenumber);
int32_t vb6_Input(int32_t filenumber, BSTR* outVar);
BSTR vb6_InputString(int32_t count, int32_t filenumber);  // P15.4: Input function — Fix <vbeclipse>: 形参序 = VB6 源码序 (count, filenumber), 旧序从未读到过东西
int32_t vb6_Kill(BSTR pathname);
int32_t vb6_MkDir(BSTR pathname);
int32_t vb6_RmDir(BSTR pathname);
int32_t vb6_ChDir(BSTR pathname);
int32_t vb6_ChDrive(BSTR drive);
int32_t vb6_Name(BSTR oldPath, BSTR newPath);
int32_t vb6_FileCopy(BSTR source, BSTR destination);

// P8.2: 随机/二进制文件访问 (Get/Put)
int32_t vb6_Get(int32_t filenumber, int32_t recnumber, void* varPtr, int32_t varSize);
int32_t vb6_Put(int32_t filenumber, int32_t recnumber, void* varPtr, int32_t varSize);

// 错误处理
// Variant比较 (VB6语义: 数值vs数值, 字符串vs字符串, 混合转Double)
int32_t vb6_VarCmpEq(vb6_VARIANT* a, vb6_VARIANT* b);
int32_t vb6_VarCmpNe(vb6_VARIANT* a, vb6_VARIANT* b);
int32_t vb6_VarCmpLt(vb6_VARIANT* a, vb6_VARIANT* b);
int32_t vb6_VarCmpGt(vb6_VARIANT* a, vb6_VARIANT* b);
int32_t vb6_VarCmpLe(vb6_VARIANT* a, vb6_VARIANT* b);
int32_t vb6_VarCmpGe(vb6_VARIANT* a, vb6_VARIANT* b);
int32_t vb6_VarCmpLongEq(vb6_VARIANT* a, int32_t b);
int32_t vb6_VarCmpLongLt(vb6_VARIANT* a, int32_t b);
int32_t vb6_VarCmpLongGt(vb6_VARIANT* a, int32_t b);
int32_t vb6_VarCmpLongLe(vb6_VARIANT* a, int32_t b);
int32_t vb6_VarCmpLongGe(vb6_VARIANT* a, int32_t b);
int32_t vb6_VarCmpLongNe(vb6_VARIANT* a, int32_t b);

int32_t vb6_ErrNumber(void);
BSTR vb6_ErrDescription(void);
void vb6_ErrClear(void);
void vb6_RaiseError(int32_t errNum, BSTR description);
BSTR vb6_ErrSource(void);
// Fix <vbeclipse> 2026-10-06: Err.LastDllError 快照 (调用点捕获, 访问时返回存储值)
int32_t vb6_ErrLastDllError(void);
void vb6_ErrSetLastDllError(int32_t code);
void vb6_ErrRaise(int32_t errNum, BSTR source, BSTR description);
void vb6_ErrRaiseNumber(int32_t errNum);

// 全局错误处理状态 (由cgen生成的代码直接使用)
extern int32_t vb6_err_resume_next;  // On Error Resume Next 标志
extern int32_t vb6_err_jmp_active;    // On Error GoTo label 标志
extern void* vb6_err_handler_label;   // 错误跳转标签 (MVP, 暂不用)
extern jmp_buf* vb6_error_jmp_ptr;    // 指向当前函数的局部jmp_buf (P12.3)
extern int32_t vb6_error_jmp_set;     // setjmp 是否已设置

// P14.1.2: Resume恢复点跟踪
extern int32_t vb6_err_resume_point;    // 出错语句resume点索引
extern int32_t vb6_err_resume_next_point;   // 出错下一句resume点索引
extern int32_t vb6_err_dispatch;        // dispatch switch变量
extern int32_t vb6_err_in_handler;      // 当前在On Error GoTo处理器中

// P12.3: On Error嵌套 — 保存/恢复调用者的错误处理状态
#define VB6_ERR_STACK_SIZE 8
void vb6_SaveErrState(void);      // 保存当前错误状态到栈 (函数入口调用)
void vb6_RestoreErrState(void);   // 从栈恢复错误状态 (函数出口调用)
// 账 #305 §B188 (第九十刀): 把一截 VB 体跑在**自己带着的错误帧**里, 返回它抛出来的号
// (0 = 干净跑完)。用途只有一处 = 发码里的 `vb6_cls_<X>_Destroy`: `Class_Terminate` 抛错时,
// 体后面那三截尾巴(字段 BSTR / 数组字段 / `vb6_Free(me)`)照旧要走完 —— 边界装在**知道尾巴在
// 哪儿的那一站**, 不装在调用它的那一站(出厂两腿 `destroyFunc` 与 `instanceClaimRelease` 都汇
// 到 `_Destroy`, 在那里套帧就是一份答案被两个地方各答一遍)。Err 里的号与描述原样留着, 由发
// 码在尾巴走完后 `vb6_RaiseError` 交回外层帧(COM 出厂那层记诊断, VB 调用方在 Err 里看见)。
int32_t vb6_RunVbProtected(void (*body)(void*), void* me);

// 字典访问
BSTR vb6_BSTR_Concat(BSTR a, BSTR b);

#ifdef __cplusplus
}
#endif
