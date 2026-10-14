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
static inline vb6_VARIANT vb6_VariantArray(void* _arr) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = (vb6_vartype)vb6_SaVtOfArrayPtr(_arr); v.parray = (struct vb6_SafeArray1D*)_arr; return v;
}

// Fix 024: Variant containing a Byte (VT_UI1)
static inline vb6_VARIANT vb6_VariantByte(uint8_t val) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtByte; v.bVal = val; return v;
}

// Fix <vbeclipse>: vb6_ReleaseObject 的反向操作, 定义在 vb6com/ (与 vb6_ReleaseObject
// 共用 vb6_ComIsDispatchable 判据)。此处前向声明, 避免 vb6rtl_variant.h 反向依赖 vb6com.h。
void vb6_ComAddRefDispatch(void* p);

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
    struct vb6_SafeArray1D*: vb6_VariantArray, \
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
    struct vb6_SafeArray1D*: vb6_VariantArray, \
    vb6_VARIANT:          vb6_VariantIdentity, \
    default:              vb6_VariantObject \
)((x))
#endif

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

/* 账 #305 §B145: 槽 → 真 VARIANT 的那张 vt 表，**只答「数值住在哪一格」**，一律借用语义
 * (不复制 BSTR、不 AddRef)。需要带权副本的出站那一族自己补下面那一步。 */
void vb6_VariantPackBorrow(vb6_VARIANT v, VARIANT *out);

/* 账 #305 §B145: Variant 参与 + - * / 的唯一运行期出口(定义在 vb6rtl_com.c)。
 * 两侧都装箱成真 VARIANT 后交 oleaut32 的 VarAdd/VarSub/VarMul/VarDiv —— VB6 本人的
 * Variant 算术引擎，于是 Decimal 的 96 位精确、Double 的小数、Null 传播、Date 加天数
 * 与它自己的舍入/越界政策一起继承，发码层不再替运行期猜档。 */
vb6_VARIANT vb6_VarAddV(vb6_VARIANT a, vb6_VARIANT b);
vb6_VARIANT vb6_VarSubV(vb6_VARIANT a, vb6_VARIANT b);
vb6_VARIANT vb6_VarMulV(vb6_VARIANT a, vb6_VARIANT b);
vb6_VARIANT vb6_VarDivV(vb6_VARIANT a, vb6_VARIANT b);
// Fix 029: Variant → SafeArray extraction (variant holding array).
// 用于调用点反向强制: callee 期望 vb6_SafeArray1D* 但实参是 vb6_VARIANT.
struct vb6_SafeArray1D* vb6_VariantToSafeArray1D(vb6_VARIANT v);
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
