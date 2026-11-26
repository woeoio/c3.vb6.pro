#pragma once
// vb6rtl_variant.h - VARIANT（VB6 Variant 类型）构造 / 转换 / 清理
// 由 vb6rtl.h 伞头 include；生成代码不要直接 include 本文件
#include "vb6rtl_base.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// VARIANT - VB6 Variant类型 (简化实现)
// ============================================================

// Vb6VarType 与 Vb6Type 枚举对应
typedef enum vb6_vartype {
    vb6_vtEmpty = 0,
    vb6_vtNull = 1,
    vb6_vtInteger = 2,
    vb6_vtLong = 3,
    vb6_vtSingle = 4,
    vb6_vtDouble = 5,
    vb6_vtCurrency = 6,
    vb6_vtDate = 7,
    vb6_vtBSTR = 8,
    vb6_vtDispatch = 9,
    vb6_vtError = 10,
    vb6_vtBoolean = 11,
    vb6_vtVariant = 12,
    vb6_vtDecimal = 14,
    vb6_vtByte = 17,
    vb6_vtArray = 0x2000,
} vb6_vartype;


typedef struct vb6_VARIANT {
    uint16_t vt;          /* VARTYPE — 必须与 Windows VARIANT 同宽 (2字节) */
    uint16_t wReserved1;
    uint16_t wReserved2;
    uint16_t wReserved3;
    union {
        int16_t iVal;
        int32_t lVal;
        float fltVal;
        double dblVal;
        BSTR bstrVal;
        void* pdispVal;
        int16_t boolVal;
        uint8_t bVal;
        /* ai/032 §6.9 收口: SByte / UInteger 两档的装箱载体。
         * 与 bVal / iVal **同槽**(union 共享存储) ⇒ sizeof(vb6_VARIANT) 逐平台不变,
         * 与 Windows VARIANT 的 x86=16 / x64=24 布局仍然对齐(OLE 的 VT_I1 就是那一字节,
         * VT_UI2 就是那一短字, 只是有符号/无符号之分)。 */
        int8_t  sbVal;   /* VT_I1  (SByte, 8-bit signed)   */
        uint16_t usVal;  /* VT_UI2 (UInteger, 16-bit unsigned) */
        int64_t cyVal;
        int64_t llVal;  /* Fix 133x: VT_I8 (LongPtr 64位指针/句柄) 存回用 */
        uint32_t ulVal;  /* ai/032 §6.6: VT_UI4 (ULong 装箱载体, 与 lVal 同槽) */
        uint64_t ullVal; /* ai/032 §6.6: VT_UI8 (ULongLong 装箱载体, 与 llVal 同槽) */
        struct vb6_SafeArray1D* parray;  /* P20-37: array pointer for GetAllSettings etc */
        /* Fix <vbeclipse> rev11: **与 Windows VARIANT 完全同布局**。
         * VB6/OLE 的 VARIANT 在 x86 = 16 字节 (2B vt + 6B 保留 + 8B union),
         * x64 = 24 字节 (8B 头 + 16B union, DECIMAL 在 union 内内联)。
         * 原先这里无条件放 16 字节的内联 decVal ⇒ x86 的 vb6_VARIANT 膨胀到
         * 24 字节, 与 SDK VARIANT(x86=16) 差 8 字节。任何把 vb6_VARIANT* 当
         * OLE VARIANT* 与 COM 互传 (IDispatch::Invoke 的 pVarResult /
         * DISPPARAMS.rgvarg) 的路径即**越界写堆** → 堆损坏 → 崩溃点漂移、
         * av target 随机 (实测 19MB/7.6MB 级)。
         * MS 的做法: 仅 x64 在 union 内联 DECIMAL(16B), x86 用 DECIMAL*
         * pdecVal(4B) —— 此处照抄同一条件编译, 令 sizeof 逐平台对齐。 */
        DECIMAL* pdecVal;   /* x86: DECIMAL 以指针形式参与 */
#ifdef _WIN64
        struct { uint16_t wReserved1; uint8_t scale; uint8_t sign; uint32_t Hi32; uint32_t Lo32; uint32_t Mid32; } decVal;
#endif
    };
} vb6_VARIANT;

#undef vb6_VARIANT  /* 取消Windows vb6_VARIANT, 使用VB6简化版 */

/* 账 #305 §B142: **`DECIMAL.wReserved` 不是「保留」—— oleaut32 拿它当 vt 标签用**(它自己写的那枚是 14 = VT_DECIMAL)。
 * 手填字段造 DECIMAL 时漏这一枚, OLE 不报错: 它返回 S_OK 却给 0 / 给空串(实测,
 * 见台账 §B142 第四轮那段的字节对照)。两份布局字段名不同(内联匿名 struct 叫 wReserved1、
 * 真 DECIMAL 叫 wReserved), 所以统一从这个宏进, 不许在调用点再把 #ifdef 抄第二遍。 */
#ifdef _WIN64
  #define vb6_DEC_VT(d)  ((d).wReserved1)
#else
  #define vb6_DEC_VT(d)  ((d).wReserved)
#endif

/* Fix <vbeclipse> rev11 配套: DECIMAL 访问统一宏。
 * x64: decVal 内联在 union 里, 直接成员访问。
 * x86: union 保持 8 字节 (对齐 OLE VARIANT), decVal 只能以 DECIMAL* pdecVal
 *      指针形式存在 —— 使用前必须已分配 (见 vb6_CDec / vb6_VariantFromComResult),
 *      vb6_VariantClear 负责释放。 */
#ifdef _WIN64
  #define vb6_VARIANT_DECVAL(v) ((v).decVal)
#else
  #define vb6_VARIANT_DECVAL(v) (*(v).pdecVal)
#endif
/* 账 #305 §B142(第三十五刀): 给槽里那枚 DECIMAL 补上 oleaut32 的 vt 标签，**只在这里写一遍**。
 * 实测(纯 C 探针)：`VarDecFromR8` 造的那枚这两字节是 **0**，只有 `VarDecFromStr` 会写 14 ⇒ 「OLE
 * 自己造的一定带标签」这句错了，标签必须在**造出来的那一步末尾**统一补，不许散在 case 里。
 * x86 那侧要判空：`CDec` 的 default 臂刚把 pdecVal 释放掉，槽还算 Decimal 但没有悬体。 */
static inline void vb6_DecSetTag(vb6_VARIANT v) {
    if (v.vt != vb6_vtDecimal) return;
#ifdef _WIN64
    vb6_DEC_VT(vb6_VARIANT_DECVAL(v)) = vb6_vtDecimal;
#else
    if (v.pdecVal) vb6_DEC_VT(vb6_VARIANT_DECVAL(v)) = vb6_vtDecimal;
#endif
}


// Variant构造
static inline vb6_VARIANT vb6_VariantEmpty(void) {
    vb6_VARIANT v;
    memset(&v, 0, sizeof(v));
    v.vt = vb6_vtEmpty;
    return v;
}

static inline vb6_VARIANT vb6_VariantNull(void) {
    vb6_VARIANT v;
    memset(&v, 0, sizeof(v));
    v.vt = vb6_vtNull;
    return v;
}

/* ucTreeView (2026-11-06): VB6 的 **Missing** —— `Optional As Variant` 实参省略时由
 * 调用方交进来。VarType(Missing) = vbError(10), 槽里是 DISP_E_PARAMNOTFOUND。
 * 这是 VB6 代码区分"实参没给"与"给了个空值 (Empty/Null)"的**唯一公开手段**:
 *   If Not VarType(N.Image) = vbError Then ...      ' 有图标才画
 * 生成侧以前在省略位填 vb6_VariantEmpty() ⇒ 上面那条判断在 C3 里整族翻面
 * (ucTreeViewPlus 展开时拿 SelectedImage 作画 ⇒ icon 丢失; 文本宽度也少算
 * ImageList 宽 ⇒ 文字被裁)。lVal 与 Windows VARIANT 的 scode 是同一格。 */
static inline vb6_VARIANT vb6_VariantMissing(void) {
    vb6_VARIANT v;
    memset(&v, 0, sizeof(v));
    v.vt = vb6_vtError;
    v.lVal = (int32_t)0x80020004;   /* DISP_E_PARAMNOTFOUND */
    return v;
}

static inline vb6_VARIANT vb6_VariantInt(int16_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtInteger; v.iVal = val; return v;
}

static inline vb6_VARIANT vb6_VariantLong(int32_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtLong; v.lVal = val; return v;
}

// Task #44: VT_I8 (LongLong / x64 的 LongPtr=intptr_t) 整型包装。
// 此前 _Generic 的 long long: 落 vb6_VariantLong(int32_t) —— x64 下 LongPtr 变量
// 装 2^32 被**静默截成 0**。x64 上 intptr_t 就是 long long, _Generic 分不开
// LongLong 与 LongPtr, 两者语义都是 64 位, 统一包成 VT_I8 才不丢位。
static inline vb6_VARIANT vb6_VariantLongLong(int64_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = (vb6_vartype)VT_I8; v.llVal = val; return v;
}

// ai/032 §6.6 收口: 无符号两档的 Variant 装箱 (VT_UI4 / VT_UI8)。
// 以前 _Generic 把 unsigned int/long/long long 全落 vb6_VariantLong(int32) ——
// `v = u` (ULong=4294967295) 读回 -1。档位对齐 C 侧实宽: 32 位无符号 → VT_UI4,
// 64 位无符号 → VT_UI8 (与 VT_I8 同槽位 ullVal/llVal, 位模式不截断)。
// (UInteger=uint16_t 走既有 `wchar_t:` 档 —— MSVC C 里 wchar_t 就是 unsigned short,
//  原表那一档本来就是它在接; 别再加 `unsigned short:` 分支, 是同一类型, C7700。)
// 全部读取侧本刀连同补齐 (ToLong/LongPtr/Bool/Double/比较/Format/TypeName/IsNumeric
// /com 包与拆/host 读表/属性包)。
static inline vb6_VARIANT vb6_VariantULong(uint32_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = (vb6_vartype)VT_UI4; v.ulVal = val; return v;
}

static inline vb6_VARIANT vb6_VariantULongLong(uint64_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = (vb6_vartype)VT_UI8; v.ullVal = val; return v;
}

static inline vb6_VARIANT vb6_VariantDouble(double val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtDouble; v.dblVal = val; return v;
}

// <vbeclipse>: VB6 的 Single 是 VT_R4 (VarType=4)、Date 是 VT_DATE (VarType=7)。
// 以前没有这两档构造器: _Generic 里 `float → vb6_VariantDouble` ⇒ Single 一律装箱成
// VT_R8, 而 Date 在 C 侧就是 double, 只能靠 VB 类型说话 (见 codegen 的 boxToVariant)。
// 注意本仓已有路径**已经**会产 VT_DATE (vb6rtl_com.c 的 VARIANT DATE 转换), 而 TypeName
// /IsDate/Format 也早有 vtDate 档 —— 缺的正是这两枚构造器和数值提取端 (见 vb6rtl.c)。
static inline vb6_VARIANT vb6_VariantSingle(float val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtSingle; v.fltVal = val; return v;
}

static inline vb6_VARIANT vb6_VariantDate(double val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtDate; v.dblVal = val; return v;   // 与 vtDate 的既有约定同: 序列号放 dblVal
}

// C30-10 第三十二刀: Currency 的 Variant 装箱 (VT_CY —— VarType 6 / TypeName "Currency")。
// C3 用 double 承载 Currency 的**值** (vb6_CCur / vb6_CCurBSTR 都如此, 见 vb6rtl_misc.c),
// 于是 `v = CCur(1.5)` 以前落 _Generic 的 `double:` 档 ⇒ VT_R8 (VarType 5 / "Double";
// 夹具 tests/dict_native 的 d121: ccur=5, 真组件 6)。这一档按 VB6 的 10000 倍定点
// 换回 VT_CY —— 口径与 COM 实参那一站 (vb6com_pack.c 的 vb6_ComPackCur) 同一份,
// 读取端 (ToDouble/ToLong/CStr/Format/比较/COM 包与拆) 本来就有 vtCurrency 臂。
// 定标口径只有一位主人: vb6rtl_base.h 的 vb6_CurrencyScale (§11.159)。改前这里自己写了
// `(int64_t)(scaled +/- 0.5)`, 与 vb6_CCur / vb6_ComPackCur 一样舍的是乘积 —— 三个站各写一遍。
static inline vb6_VARIANT vb6_VariantCur(double val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtCurrency;
    v.cyVal = vb6_CurrencyScale(val);
    return v;
}

static inline vb6_VARIANT vb6_VariantString(BSTR val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtBSTR; v.bstrVal = val; return v;
}

static inline vb6_VARIANT vb6_VariantBool(int16_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtBoolean; v.boolVal = val; return v;
}

// Variant containing a SafeArray (VB6: a = Array(1,2,3))
// 账 #305 §B155 格①（第四十六刀）: 标签不再写死成 vtArray|vtVariant —— 它问描述符自己那一档。
// 改前任何 typed 数组装箱后 VarType 一律答 8204、TypeName 一律答 "Variant"（VB6: String 数组是
// 8200 / "String()"）；而且拿着 VT_ARRAY|VT_VARIANT 的标签去指一枚 BSTR 载体，交回 COM 就是错档。
// 为什么声明在本文件、定义在 vb6rtl_array.c: 伞头里 variant.h 排在 array.h **之前**（那边的
// vb6_VariantFromUdtBytes 注释记着这条顺序），这里看不见 elemtype 枚举与两套描述符。
int32_t vb6_SaVtOfArrayPtr(const void* arr);
/* 账 §11.179 (第一百九十六刀): 装箱口。1D → 深拷一份副本 (带 OWNED 标记, 随 Variant 销毁);
   ND / 真 SAFEARRAY → 照旧借用。声明在本文件、定义在 vb6rtl_array.c, 与上面那枚同一个理由。 */
vb6_VARIANT vb6_VariantArrayCopy(void* _arr);
/* 账 §11.180 (第一百九十七刀): 带元素档的装箱口 —— 描述符不在 (ReDim 之前的动态数组 /
   Erase 之后) 时按**声明**那一档写标签。参差与上面那一枚同住 vb6rtl_array.c。 */
vb6_VARIANT vb6_VariantArrayTyped(void* _arr, int32_t elemEnum);
static inline vb6_VARIANT vb6_VariantArray(void* _arr) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = (vb6_vartype)vb6_SaVtOfArrayPtr(_arr); v.parray = (struct vb6_SafeArray1D*)_arr; return v;
}

// Fix 024: Variant containing a Byte (VT_UI1)
static inline vb6_VARIANT vb6_VariantByte(uint8_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtByte; v.bVal = val; return v;
}

// ai/032 §6.9 收口: 小整型两档的 Variant 装箱 (VT_I1 / VT_UI2)。
// 为什么这两档以前答不对: `vb6_VariantFromValue` 的 _Generic 按 **C 类型**分派, 而
//   SByte    = int8_t   = char         -> 命中既有的 `char:` 档 -> vb6_VariantLong = VT_I4
//   UInteger = uint16_t = unsigned short -> 命中既有的 `wchar_t:` 档 -> vb6_VariantLong = VT_I4
// (MSVC C 里 wchar_t 就是 unsigned short, 所以那一档本来在接它; 见上面 VT_UI4 那段
//  注释里「别再加 unsigned short 分支, 会 C7700」—— 那个约束依然成立, 所以这里
//  **不去碰 _Generic**, 而是让唯一权威 boxToVariant 直接点名调这两枚出口。)
// 值没丢(8/16 位塞得进 VT_I4), 丢的是**类型身份**: TypeName 答 "Long"、VarType 答 3。
// 与 ULong/VT_UI4、ULongLong/VT_UI8 同一口径 —— 那两档的注释写着「答 "Long" 才是误导」。
static inline vb6_VARIANT vb6_VariantSByte(int8_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = (vb6_vartype)VT_I1; v.sbVal = val; return v;
}

static inline vb6_VARIANT vb6_VariantUInteger(uint16_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = (vb6_vartype)VT_UI2; v.usVal = val; return v;
}

// Fix <vbeclipse>: vb6_ReleaseObject 的反向操作, 定义在 vb6com/ (与 vb6_ReleaseObject
// 共用 vb6_ComIsDispatchable 判据)。此处前向声明, 避免 vb6rtl_variant.h 反向依赖 vb6com.h。
void vb6_ComAddRefDispatch(void* p);

/* ParamArray (Windows SAFEARRAY of VT_VARIANT) → vb6_VARIANT 的那一趟:
   `varParam = arrParam` 曾落进 _Generic default ⇒ 假 VT_DISPATCH ⇒ UBound 错误 9。
   定义在 vb6rtl_com.c (与 vb6_ComVariantIntoSlot 同一 TU)。源数组只读。 */
vb6_VARIANT vb6_VariantFromParamArray(SAFEARRAY* psa);

// Fix 024: Variant containing a COM object pointer (VT_DISPATCH)
// Fix <vbeclipse>: **必须 AddRef** —— vb6_VariantClear() 对 vb6_vtDispatch 会调
// vb6_ReleaseObject, 构造侧不取引用而析构侧释放 = 过度释放 → 0xC0000374 堆损坏
// (实证 ucPerspective.AddView 的 `m_Views.Add ViewId, l_View` + 随后 `Set l_View = Nothing`:
//  容器里留下指向已释放对象的 Variant)。VB6 里 `coll.Add key, obj` 也是 AddRef 的。
// 防御: 只对真 COM 接收者 AddRef (vb6_ComAddRefDispatch 内含 vb6_ComIsDispatchable),
// 裸结构体/UDT 地址照旧只存指针。
static inline vb6_VARIANT vb6_VariantObject(void* val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtDispatch; v.pdispVal = val;
    vb6_ComAddRefDispatch(val);
    return v;
}

// Fix 024: Variant identity passthrough — used by vb6_VariantFromValue
//           when the expression is ALREADY a vb6_VARIANT (no wrapping).
static inline vb6_VARIANT vb6_VariantIdentity(vb6_VARIANT v) { return v; }

/* 账 #305 §11.106 (第一百二十四刀): **不取引用**的对象档装箱 —— 只给只读谓词的实参用。
 *
 * 病根 (第一百二十三刀的 D2 读数): `TypeName(<对象槽>)` / `IsObject(<对象>)` 这类调用的实参
 * 先在装箱站成 `vb6_VariantFromValue(...)` —— 对 void* / 包装器那是 `vb6_VariantObject`,
 * 它按 Fix 024 必须 AddRef; 而按值形参的这枚临时 VARIANT 谁也不会 `vb6_VariantClear`
 * (它在表达式里, 不是具名变量) ⇒ 每读一次漏一枚引用。对象是包装器时那一枚就是包装器的命:
 * 槽清空后仍是 rc=1, 类引用计数到不了 0, `Class_Terminate` 永不发生。
 *
 * 为什么能借: 这九枚 (`TypeName/VarType/IsObject/IsNull/IsEmpty/IsArray/IsDate/IsError/
 * IsNumeric`) 的 RTL 体只读 `vt` 与 payload, 既不存也不清 (vb6rtl_conv.c 逐行可对); 实参的
 * 生存期覆盖整个调用, 借来的指针在调用期间一直有效。改用这一档后**不分配任何东西**,
 * 也就不存在"谁退"那一问。 */
static inline vb6_VARIANT vb6_VariantObjectBorrowed(void* val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtDispatch; v.pdispVal = val;
    return v;
}

/* 账 #305 §11.126 (第一百四十五刀): **所有权转移**的对象档装箱 —— 指针里那一枚引用归
 * 收到这一枚 VARIANT 的那一头，构造侧不再 AddRef。三兄弟只差内核:
 *   vb6_VariantObject          取一枚 (槽自己持有一份, 清槽时退)
 *   vb6_VariantObjectBorrowed  不取 (只读谓词实参, 谁也不必退)
 *   vb6_VariantObjectTransfer  接管 (调用方已把一枚交出来: FromInstance / 函数返回槽)
 * 为什么单开一枚而不复用 Borrowed: 那两个是**两件不同的事**。借来的一枚随时可能被原主
 * 退掉, 转移来的一枚只有收到它的人退 —— 共用一枚出口, 下一次有人给 Borrowed 加口径
 * (比如改成"仅在…")就必然连带改坏转移的语义。 */
static inline vb6_VARIANT vb6_VariantObjectTransfer(void* val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtDispatch; v.pdispVal = val;
    return v;
}

// Fix 024: Polymorphic Variant constructor via C11 _Generic.
//   用法: vb6_VARIANT v = vb6_VariantFromValue(any_C_expr);
//   按实参表达式的C类型(编译期推断)选择合适的Variant构造函数:
//     _Bool/bool                        -> VariantBool
//     整型(char/short/int/long/long long,
//          signed/unsigned, wchar_t)    -> VariantLong 或 VariantByte(uchar) 或 VariantInt(short)
//     浮点(float/double/long double)    -> VariantDouble
//     wchar_t*  (BSTR别名)              -> VariantString
//     struct vb6_SafeArray1D*           -> VariantArray
//     vb6_VARIANT (已是变体)            -> VariantIdentity (no-op)
//     其他指针 (void*/class*/type*)     -> VariantObject
//   用途: 当 C3 的 inferExprType 把标量/LenB(...)误判为 Variant 时,
//         temp 变量声明 "vb6_VARIANT _vcmp_N = scalar;" 会触发 C2440.
//         改用 vb6_VariantFromValue(scalar) 让编译器按实类型自动包装, 消除 C2440.
// P4 崩溃取证 (2026-11-04): 两个 SafeArray 描述符是**两套结构**, _Generic 只列了 1D 那一档,
// 于是 Nd 指针落 default: vb6_VariantObject ⇒ vt=VT_DISPATCH 而 pdispVal 是描述符地址。
// 后果不是"读错值"而是**假对象**: `v = a`(2D) 的 VarType 答 9, IsArray 答 False;
// 再交给 For Each / 成员访问时运行期把它当 IDispatch 解引用 ⇒ 0xC0000005 (读 +0x16)。
// 修法在**装箱口**补这一档, 与 1D 同一枚 vb6_VariantArray (它自己问描述符要 vt)。
struct vb6_SafeArrayND; // forward declaration (1D 的那一枚在 vb6rtl_builtin.h)
#ifdef _MSC_VER
// Fix 092u: MSVC 中 char 默认等价于 signed char, _Generic 关联列表同时出现
// char:/signed char: 被视为重复关联 (error C7700, "char 与之前的 char 不兼容").
// MSVC 下仅保留 char: (已覆盖 signed char); 其他编译器保留标准的三态区分.
#define vb6_VariantFromValue(x) _Generic((x), \
    _Bool:                vb6_VariantBool, \
    char:                 vb6_VariantLong, \
    unsigned char:        vb6_VariantByte, \
    short:                vb6_VariantInt, \
    int:                  vb6_VariantLong, \
    unsigned int:         vb6_VariantULong, \
    long:                 vb6_VariantLong, \
    unsigned long:        vb6_VariantULong, \
    long long:            vb6_VariantLongLong, \
    unsigned long long:   vb6_VariantULongLong, \
    wchar_t:              vb6_VariantLong, \
    float:                vb6_VariantSingle, \
    double:               vb6_VariantDouble, \
    wchar_t*:             vb6_VariantString, \
    struct vb6_SafeArray1D*: vb6_VariantArrayCopy, \
    struct vb6_SafeArrayND*: vb6_VariantArray, \
    struct tagSAFEARRAY*:  vb6_VariantFromParamArray, \
    vb6_VARIANT:          vb6_VariantIdentity, \
    default:              vb6_VariantObject \
)((x))
#else
#define vb6_VariantFromValue(x) _Generic((x), \
    _Bool:                vb6_VariantBool, \
    char:                 vb6_VariantLong, \
    signed char:          vb6_VariantLong, \
    unsigned char:        vb6_VariantByte, \
    short:                vb6_VariantInt, \
    int:                  vb6_VariantLong, \
    unsigned int:         vb6_VariantULong, \
    long:                 vb6_VariantLong, \
    unsigned long:        vb6_VariantULong, \
    long long:            vb6_VariantLongLong, \
    unsigned long long:   vb6_VariantULongLong, \
    wchar_t:              vb6_VariantLong, \
    float:                vb6_VariantSingle, \
    double:               vb6_VariantDouble, \
    wchar_t*:             vb6_VariantString, \
    struct vb6_SafeArray1D*: vb6_VariantArrayCopy, \
    struct vb6_SafeArrayND*: vb6_VariantArray, \
    struct tagSAFEARRAY*:  vb6_VariantFromParamArray, \
    vb6_VARIANT:          vb6_VariantIdentity, \
    default:              vb6_VariantObject \
)((x))
#endif

/* 账 #305 §11.106 (第一百二十四刀): `vb6_VariantFromValue` 的**借用档** —— 整形与标量
 * 那几档与正式档完全同一枚构造器 (它们只写值, 不拥有任何东西); 差异只在最后一条:
 * 指针类落 `vb6_VariantObjectBorrowed` (不 AddRef)。臂表必须与上面的正式档逐条对齐
 * —— 少一条就会落到 default, 把标量当对象。
 * 只给九枚只读谓词 (TypeName/VarType/IsObject/IsNull/IsEmpty/IsArray/IsDate/IsError/
 * IsNumeric) 的实参站用; 任何会**存**实参的出口都不许用这一档。 */
#ifdef _MSC_VER
#define vb6_VariantBorrowFromValue(x) _Generic((x), \
    _Bool:                vb6_VariantBool, \
    char:                 vb6_VariantLong, \
    unsigned char:        vb6_VariantByte, \
    short:                vb6_VariantInt, \
    int:                  vb6_VariantLong, \
    unsigned int:         vb6_VariantULong, \
    long:                 vb6_VariantLong, \
    unsigned long:        vb6_VariantULong, \
    long long:            vb6_VariantLongLong, \
    unsigned long long:   vb6_VariantULongLong, \
    wchar_t:              vb6_VariantLong, \
    float:                vb6_VariantSingle, \
    double:               vb6_VariantDouble, \
    wchar_t*:             vb6_VariantString, \
    struct vb6_SafeArray1D*: vb6_VariantArray, \
    struct vb6_SafeArrayND*: vb6_VariantArray, \
    struct tagSAFEARRAY*:  vb6_VariantFromParamArray, \
    vb6_VARIANT:          vb6_VariantIdentity, \
    default:              vb6_VariantObjectBorrowed \
)((x))
#else
#define vb6_VariantBorrowFromValue(x) _Generic((x), \
    _Bool:                vb6_VariantBool, \
    char:                 vb6_VariantLong, \
    signed char:          vb6_VariantLong, \
    unsigned char:        vb6_VariantByte, \
    short:                vb6_VariantInt, \
    int:                  vb6_VariantLong, \
    unsigned int:         vb6_VariantULong, \
    long:                 vb6_VariantLong, \
    unsigned long:        vb6_VariantULong, \
    long long:            vb6_VariantLongLong, \
    unsigned long long:   vb6_VariantULongLong, \
    wchar_t:              vb6_VariantLong, \
    float:                vb6_VariantSingle, \
    double:               vb6_VariantDouble, \
    wchar_t*:             vb6_VariantString, \
    struct vb6_SafeArray1D*: vb6_VariantArray, \
    struct vb6_SafeArrayND*: vb6_VariantArray, \
    struct tagSAFEARRAY*:  vb6_VariantFromParamArray, \
    vb6_VARIANT:          vb6_VariantIdentity, \
    default:              vb6_VariantObjectBorrowed \
)((x))
#endif

// 账 #305 新账 #407 第一百七十一刀: 一枚**裸实例指针**装成对象档 VARIANT 的唯一出口。
// 视图那一档 (vb6_iface_<I> 是值结构体) 落 vb6_VariantFromValue 的 _Generic default
// = vb6_VariantObject 是 cl C2172「实参不是指针」的主人，那一档改问这一句。
// 定义不在本层 —— 身份表与 IDispatch 包装器都在 src/rtl/core/vb6comserver/vb6comserver_obj.c。
vb6_VARIANT vb6_VariantFromProjectInstance(void* instance);

// Index into a Variant that holds an array -- returns element as vb6_VARIANT
vb6_VARIANT vb6_VariantArrayGet(vb6_VARIANT* v, int32_t index);
// Set element in a Variant that holds an array
void vb6_VariantArraySet(vb6_VARIANT* v, int32_t index, vb6_VARIANT val);
// 账 #305 §B150 (第四十刀): 晚绑定槽的多维元素读/写 —— 秩数与全部下标一次交全,
// 上下界与错误 9 由 vb6rtl_array.h 那唯一一枚 vb6_SaNdElemPtr 管。
vb6_VARIANT vb6_VariantArrayGetND(vb6_VARIANT* v, int32_t rank, ...);
void vb6_VariantArraySetND(vb6_VARIANT* v, int32_t rank, ...);
// Fix 084f: Variant 数组嵌套索引按值版本 — vGateway(lIdx)(0) 中内层
// vb6_VariantArrayGet(&vGateway, lIdx) 返回 vb6_VARIANT 值(非左值无法取地址),
// 此函数按值接收后再按索引取元素, 替代非法的 (vb6_VARIANT){...} 复合字面量.
vb6_VARIANT vb6_VariantArrayGetVal(vb6_VARIANT v, int32_t index);
// Fix 158f: Variant 数组元素的取址访问 — Declare As Any ByRef 调用把
// Variant 数组元素当指针传 (VTableHandle.bas: VariantCopy ArgListRev(i),
// ArgList(UBound(ArgList) - i)). vb6_VariantArrayGet 返回值(右值)无法取址,
// 此前生成 (void*)(intptr_t)(vb6_VariantArrayGet(...)) → C2440 (vb6_VARIANT→
// intptr_t). 此函数返回底层 SafeArray 元素槽位的地址 (仅 vb6_sa_variant 支持,
// 其余类型返回 NULL — 非 Variant 数组槽位不是 VARIANT 布局).
vb6_VARIANT* vb6_VariantArrayElemPtr(vb6_VARIANT* v, int32_t index);
/* 账 #305 §B150: 多维版取址, 形参表 = (槽, 秩数, 秩数个 int32_t 下标) —— 与
   vb6_VariantArrayGetND 同形, 发码侧那两路交址的改写只换函数名。 */
vb6_VARIANT* vb6_VariantArrayElemPtrND(vb6_VARIANT* v, int32_t rank, ...);

// Variant转基本类型
int32_t vb6_VariantToLong(vb6_VARIANT v);
// Variant → LongPtr (指针/句柄语义): 与 vb6_VariantToLong 数值提取一致,
// 但返回 intptr_t, 避免 x64 下 64 位句柄/指针被截断.
intptr_t vb6_VariantToLongPtr(vb6_VARIANT v);
// Fix 093a: Variant → Boolean (CBool 语义). 调用点 (ToolsJsonVba 等) 依赖符号名
// vb6_VariantToBool; 此前缺失 → LNK2019.
int16_t vb6_VariantToBool(vb6_VARIANT v);
double vb6_VariantToDouble(vb6_VARIANT v);
void vb6_VariantSetI4(vb6_VARIANT* v, int32_t x);   // Fix 133: Declare 标量 ByRef 出参回写
void vb6_VariantSetI8(vb6_VARIANT* v, intptr_t x);  // Fix 133x: LongPtr 出参回写 (VT_I8, x64 指针不截断)
BSTR vb6_VariantToString(vb6_VARIANT v);

/* 账 #305 §B142: Decimal 在「往外读」这一侧的**唯一两个换算出口**。都走 oleaut32
 * 的 `VariantChangeType` —— 不是选的, 是本机只够得着它: `VarDecToR8` / `VarDecToString` /
 * `VarDecToI4` 在 SDK x64 的 oleaut32.lib 里根本没有(台账 §B142 ③)。
 * 成功返回真; 失败给 0 / 空串, 不抛错(与其它读表的 default 同形)。 */
BSTR vb6_DecimalToBSTR(vb6_VARIANT v);
int vb6_DecimalToR8(vb6_VARIANT v, double *out);

/* 账 #305 §B142/§B145: 槽里那枚 Decimal 变成合法 OLE VARIANT 的**唯一出口**(见 vb6rtl.c)。
 * 返回 0 = 槽不是 Decimal; 返回 1 = out 已铺好(带 oleaut32 的 vt 标签、且是副本)。 */
int vb6_DecimalAsVariant(vb6_VARIANT v, VARIANT *out);

/* 账 #305 §B144: 与上面那枚成对的**反向一问** ——「这枚 OLE VARIANT 的 DECIMAL 住在哪」。
 * 两架构的答案都是结构体的头，分歧只在 SDK 的称呼；入站读 Decimal 的地方只问这里，
 * 不再各自数偏移(x64 数错成 offset 8、x86 把数值当指针，两处都量到过)。 */
DECIMAL *vb6_DecimalPtrFromVariant(const VARIANT *v);

/* 账 #305 §B144: COM 服务端适配器那两份「槽 ↔ OLE VARIANT」的 memcpy 收成下面这一对出口。
 * 服务端那份 .c 手上只有匿名槽(3×uint64)，不该知道槽里每一档住在第几个字节 ——
 * 它原来那句「槽布局前 16 字节与 OLE VARIANT 内容一致」正是 §B144 静默错的根因。 */
void vb6_VariantPackCopyInto(const void *slot, VARIANT *out);    /* 槽 → 对端 VARIANT(带权副本) */
void vb6_VariantBorrowIntoSlot(const VARIANT *src, void *slot);  /* 对端 VARIANT → 槽(借用) */
void vb6_VariantSlotRelease(void *slot);  /* 账 #305 §B162: 回收上面那臂造出的描述符 */
/* 账 #305 §B222(第二百三十刀): 数组形参的出站回写(VT_BYREF|VT_ARRAY 那一形)。
 * 必须在 vb6_VariantSlotRelease **之前**调用；成功后槽的 vt 被补成数组档，收尾照常回收。 */
void vb6_ArraySlotWriteBack(VARIANT *client, void *slot);
/* 账 #305 §B223(第二百三十三刀): ByRef 标量形参的出站回写(数值/布尔那几档)。
 * 同样必须在 vb6_VariantSlotRelease **之前**调用；BSTR/对象/Variant 三档不猜。 */
void vb6_ByRefScalarWriteBack(VARIANT *client, void *slot);

/* 富贵竹 FIX-011/012 (账 #610 第341刀): [out] VARIANT 出参槽 ↔ C3 描述符 的两向适配。
 *
 * 读侧 (FIX-011): 服务端 (AutoCAD GetBoundingBox 一族) 往槽里写的是**真 SAFEARRAY**
 *   (oleaut32 自家布局), 而本仓的数组槽是 C3 自有描述符 (fFeatures=0xA1D1) ⇒ 读之前
 *   要归一化。vb6_SlotHasRealSafeArray 回答「槽里挂的是不是真 SAFEARRAY」(判据与
 *   vb6_SaIs1D 同源); vb6_VariantNormalizeComSlot 就地把真 SAFEARRAY 换成 C3 描述符。
 *   vb6_VariantArrayGet / vb6_VariantArrayGetND 入口各有一道**惰性兜底**(万一某条路
 *   没走收尾函数, 读的时候也能自愈)。
 * 写侧 (FIX-012 闭环): 下一次把槽递回服务端前, 要先把 C3 描述符换回真 SAFEARRAY,
 *   否则服务端覆写出参前的 VariantClear 会把描述符当真 SAFEARRAY SafeArrayDestroy,
 *   且描述符 fFeatures 里 FADF_BSTR 位是置起的 ⇒ double 数据区被当 BSTR 指针逐个
 *   SysFreeString ⇒ 堆损坏 / Invoke 静默失败。该方向的换手在 vb6com.h
 *   (vb6_VariantComSlotToRealArray), 因为递出点在 vb6com_pack.c。
 *
 * ⚠ 这两个声明**必须在本头**: vb6rtl_com.c 里的 vb6_VariantArrayGet 会调它们,
 *   缺原型 ⇒ C4013 隐式声明 (返回 int 截断指针, x64 崩雷)。 */
int vb6_SlotHasRealSafeArray(void* pVar);
void vb6_VariantNormalizeComSlot(void* pVar);

/* 账 #305 §B219(第二百二十刀): 入站实参「这一槽声明的是什么档」的那一问，值域就两样：
 * 一个真 VARTYPE(要换算到它)，或 VB6_ARGVT_SKIP(这一格本刀不问，槽里保持纯借用)。
 * 声明档不明 / ByRef / Variant / 数组 那几形交回 SKIP —— 它们的 C 存储是不是联合体那一格
 * 还没量过(见 022 第二百二十刀的在册欠项)，猜一档就是造第二套口径。 */
#define VB6_ARGVT_SKIP ((VARTYPE)0xFFFF)

/* 账 #305 §B219: 「往里读」那一族在此之上多出来的一问 ——「调用方递的 vt 与形参声明档
 * 不合时怎么办」。发码那 16 处读臂各自按声明档取联合体成员、从不看标签(实测两架构：
 * VT_BSTR 进 Long 档答指针低 32 位、VT_I4 进 String 档把 40 当 BSTR 解 ⇒ 原生 AV 打死客户端)。
 * 规则三条：同档不动(槽里就是答案)；异档换算产物**先进自己的临时量**再走上面那枚借用出口
 * —— 不许就地覆盖槽(槽借用了调用方的 BSTR，VariantChangeType 会把它 SysFreeString 掉)；
 * 换算不动交回 DISP_E_TYPEMISMATCH(客户端 vb6_ComCheckError 已把它折成 VB 13)。
 * scratch 由调用方在**被调方返回之后** VariantClear —— 顺序与 §B185 那扇边界同规。 */
HRESULT vb6_ComArgNormalize(const VARIANT *src, VARTYPE want, void *slot, VARIANT *scratch);

/* 账 #305 §11.142 (第一百六十刀): 上面那一枚**失败的原因**。
 * 以前失败一律报 DISP_E_TYPEMISMATCH，于是收窄越界（`o.I = 40000` 进 Integer 形参）
 * 在客户端被折成 13「类型不匹配」，而真 VB6 报 **6 Overflow** 且目标不变。
 * 失败那一瞬读回（同一 STA 线程、同一个循环里），不要当跨调用状态用。 */
void vb6_ComArgNormLastHrSet(HRESULT hr);
HRESULT vb6_ComArgNormLastHr(void);

/* 账 #305 §B224 (第二百三十八刀): 「VARIANT → 目标档」这一问的唯一出口(定义在 vb6rtl_com.c)。
 * 除了一档之外全交 oleaut32 的 VariantChangeType；那一档是**布尔转字符串**：
 * Automation 答 "-1"/"0"，VB 语言答 "True"/"False"。真 VB6 6.0 服务端实测
 * `EchoStr(ByVal s As String)` 收 VT_BOOL 答 "[True]"/"[False]"(ai/probes/bool_string
 * A1/A2) ⇒ 服务端入站那一格与 VB 语言同口径，不是 Automation 的换算口径。
 * 全仓凡以 VT_BSTR 为目的档的换算都问这里(服务端入站 / COM 结果解包 / 十进制转串)。
 * dest 必须是 VariantInit 过的空 VARIANT；与 VariantChangeType 一样 dest 可等于 src。 */
HRESULT vb6_ComChangeType(VARIANT *dest, const VARIANT *src, VARTYPE want);

/* 账 #305 §B145: 槽 → 真 VARIANT 的那张 vt 表，**只答「数值住在哪一格」**，一律借用语义
 * (不复制 BSTR、不 AddRef)。需要带权副本的出站那一族自己补下面那一步。 */
void vb6_VariantPackBorrow(vb6_VARIANT v, VARIANT *out);

/* 账 #305 §B145: Variant 参与 + - * / 的唯一运行期出口(定义在 vb6rtl_com.c)。
 * 两侧都装箱成真 VARIANT 后交 oleaut32 的 VarAdd/VarSub/VarMul/VarDiv —— VB6 本人的
 * Variant 算术引擎，于是 Decimal 的 96 位精确、Double 的小数、Null 传播、Date 加天数
 * 与它自己的舍入/越界政策一起继承，发码层不再替运行期猜档。 */
vb6_VARIANT vb6_VarAddV(vb6_VARIANT a, vb6_VARIANT b);
/* C30-10 第四十五刀: 一元 `-` 打在 Variant 上（子类型保持）。
   原有的四个 vb6_Var*V 只管二元运算，一元取负在 VB6 里是**保子类型**的
   （-Integer→Integer、-Long→Long、-Double→Double），而 -BSTR 是错误 13。
   后端 visit(UnaryExpr) 以前直接发 `(-<struct>)` —— C2088，整段编不过。 */
vb6_VARIANT vb6_VarNegV(vb6_VARIANT a);
// C30-10 第五十三刀: 一元 `Not` 打在 Variant 上（Boolean 逻辑非 / 整型位非并保档 / Null 原样）。
vb6_VARIANT vb6_VarNotV(vb6_VARIANT a);
// C30-11 第六十五刀: 比较符的 Null 传播 —— 任一侧槽是 Null ⇒ 整枚 Null；
// 否则就是已算好的布尔结果（VB6 True = -1）。读数 .temp/null/oracle_ops.vbs O5..O10。
vb6_VARIANT vb6_VCmpNullV(vb6_VARIANT* a, vb6_VARIANT* b, int32_t cond);
// C30-11 第六十六刀: `&` 的双 Null 折叠 —— 两侧槽都是 Null ⇒ Null，否则就是正常拼接。
// 读数 .temp/null/oracle_ops3.vbs R1/R4..R6/R10/R12（只有两侧同时 Null 才折）。
vb6_VARIANT vb6_VarConcatNullV(vb6_VARIANT* a, vb6_VARIANT* b, BSTR sa, BSTR sb);
// C30-11 第六十七刀: Or/And 的 Null 真值表（Or 只有另一侧 True 才确定，And 只有另一侧 False）。
// isAnd=0 走 Or、1 走 And；mask 的 bit0/bit1 表示 a/b 是不是槽（不是槽的那侧用 litBits）。
vb6_VARIANT vb6_VarConcatNullVV(vb6_VARIANT a, vb6_VARIANT b);

vb6_VARIANT vb6_VarLogicNullV(int32_t op, vb6_VARIANT* a, vb6_VARIANT* b,
                              int32_t litBits, int32_t mask, int32_t litType);
vb6_VARIANT vb6_VarSubV(vb6_VARIANT a, vb6_VARIANT b);
vb6_VARIANT vb6_VarMulV(vb6_VARIANT a, vb6_VARIANT b);
vb6_VARIANT vb6_VarDivV(vb6_VARIANT a, vb6_VARIANT b);
// Fix 029: Variant → SafeArray extraction (variant holding array).
// 用于调用点反向强制: callee 期望 vb6_SafeArray1D* 但实参是 vb6_VARIANT.
struct vb6_SafeArray1D* vb6_VariantToSafeArray1D(vb6_VARIANT v);
// 账 #305 §B172 (第七十刀): UBound/LBound 变体实参站的**严格**提取 —— 不是数组
// (含 Null / Empty / 标量) 一律抛 13 Type mismatch (真 VB6 IDE 读数), 与上面那枚
// 宽松版 (返回 NULL) 分家, 因为宽松版还被 Byte()/Filter 的探测路径共用。
struct vb6_SafeArray1D* vb6_VariantToSafeArray1DStrict(vb6_VARIANT v);
// Fix 140: Variant → Byte() 数组. Variant 持数组时返回其 parray; 持 BSTR 时
// 按 VB6 语义把字符串转成字节数组 (vb6_StringToByteArray). 用于 Byte() 字段
// 从字符串 Variant 赋值 (LabelPlus.ctl `m_Caption = .ReadProperty("Caption", ...)`
// — PropertyBag 里 Caption 存的是字符串, 旧路径 vb6_VariantToSafeArray1D 对
// 字符串 Variant 返回 NULL, 把 caption 清空 → 右侧 KPI 卡片空白).
struct vb6_SafeArray1D* vb6_VariantToByteArray(vb6_VARIANT v);
// Fix 029: Variant → void* (按值传入, 避免调用点包装时的左值问题).
// 与 vb6_VariantToObject(vb6_VARIANT*) 互补; 后者要求实参是左值 (取地址),
// 但 IndexOrCallExpr 调用点包装的实参可能是函数返回的右值, 无法取址.
void* vb6_VariantToObjectVal(vb6_VARIANT v);

// P3 缺口① / ③: Set / With / `Is Nothing` 的阵列判据 (判据只看运行期 vt 的阵列位)。
// 非阵列档原样转调上面的 vb6_VariantToObjectVal / vb6_VariantToObject, 既有行为不动;
// 阵列档分别报 VB6 实测错误号 424 (目标 Object) / 13 (目标 Variant) / 424 (Is Nothing)。
// 读数出处: .build/vb6base/READINGS.md (P1 表 T1–T4、P3 表 A/E/I/J/K/L)。
void* vb6_VariantToObjectValChecked(vb6_VARIANT v);
vb6_VARIANT vb6_VariantSetCopyChecked(vb6_VARIANT v);
int32_t vb6_VariantIsNothingChecked(vb6_VARIANT* v);

// P8.4: Variant清理(释放内含BSTR等资源)
void vb6_VariantClear(vb6_VARIANT* v);
// P8.4: Variant深拷贝(复制BSTR)
void vb6_VariantCopy(vb6_VARIANT* dst, const vb6_VARIANT* src);
// Fix <vbeclipse> rev16: Variant **槽位赋值**(接管语义) —— 先释放 dst 旧内容, 再把
// src 的值**深拷贝**进去 (BSTR 另分配一只 / Dispatch AddRef / x86 DECIMAL 另分配)。
// 动机: Variant 数组元素赋值原先发的是 `slot = vb6_VariantFromValue((*Item))`, 而
// vb6_VariantFromValue 对 vb6_VARIANT 是 identity **浅拷贝** ⇒ 槽与调用方实参共用
// 同一只 BSTR。调用方 (ByRef Variant 形参的宿主) 下一次 vb6_VariantClear(&v) 就把它
// free 掉, 后续同尺寸分配复用该地址 ⇒ 每个槽都读出"最后一次写入的值"。
// 实测 tests/ve_list (工程内 List.cls 形态): idx=0/1/2 全答 "Gamma";
// 真工程 play78: Folder.Views.Item(0) 取回 1 字符垃圾 → m_Views.Item(<垃圾>) 落空。
void vb6_VariantAssign(vb6_VARIANT* dst, vb6_VARIANT src);
// 账 #305 §B165 (第五十四刀): 值档(BSTR + x86 那一台的 DECIMAL)所有权的一对出口。
// Clear/Copy/Assign 与宿主环(Host_* 与内建 Collection 五站)都问它们 —— 「谁分配谁释放」
// 只许住这一处; 对象那一档不在这对里(加减引用归 Copy/Clear 自己问 vb6_ComIsDispatchable)。
void vb6_VariantValueFree(vb6_VARIANT* v);
void vb6_VariantValueCopy(vb6_VARIANT* dst, const vb6_VARIANT* src);

#ifdef __cplusplus
}
#endif
