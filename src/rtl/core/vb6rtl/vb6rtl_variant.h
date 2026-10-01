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
static inline vb6_VARIANT vb6_VariantArray(void* _arr) {
    vb6_VARIANT v; memset(&v, 0, sizeof(v));
    v.vt = vb6_vtArray | vb6_vtVariant; v.parray = (struct vb6_SafeArray1D*)_arr; return v;
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
    unsigned int:         vb6_VariantLong, \
    long:                 vb6_VariantLong, \
    unsigned long:        vb6_VariantLong, \
    long long:            vb6_VariantLongLong, \
    unsigned long long:   vb6_VariantLong, \
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
    unsigned int:         vb6_VariantLong, \
    long:                 vb6_VariantLong, \
    unsigned long:        vb6_VariantLong, \
    long long:            vb6_VariantLongLong, \
    unsigned long long:   vb6_VariantLong, \
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

#ifdef __cplusplus
}
#endif
