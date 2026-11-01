#pragma once
// vb6rtl_builtin.h - 内置函数与系统函数声明（字符串/数学/日期/财务/文件等）
// 由 vb6rtl.h 伞头 include；生成代码不要直接 include 本文件
#include "vb6rtl_base.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 内置函数 (最小子集)
// ============================================================

// 字符串函数
int32_t vb6_Len(BSTR s);
int32_t vb6_LenB_BSTR(BSTR s);  // Fix 048: LenB for BSTR — byte length of string
// Fix 048: LenB — _Generic macro: BSTR → byte length, UDT → sizeof
// VB6 LenB() works for both strings (byte length) and UDTs (structure size)
#define vb6_LenB(x) _Generic((x), \
    BSTR: vb6_LenB_BSTR(x), \
    const BSTR: vb6_LenB_BSTR(x), \
    default: ((int32_t)sizeof(x)) \
)
BSTR vb6_Left(BSTR s, int32_t n);
BSTR vb6_Right(BSTR s, int32_t n);
BSTR vb6_Mid(BSTR s, int32_t start, int32_t len);
int32_t vb6_InStr(int32_t start, BSTR haystack, BSTR needle);
// <vbeclipse>: vbTextCompare 共用核 (定义在 vb6rtl_string.c, InStr/InStrRev/Replace/
// Split/Filter/StrComp 共用)。compare: 0 = vbBinaryCompare, 非 0 = vbTextCompare。
// 位置都是 0 基; 找不到返回 -1 (InStr 系列在外面换算 1 基与 0)。
int32_t vb6_TextFind(BSTR haystack, BSTR needle, int32_t from0, int32_t compare);
int32_t vb6_TextFindRev(BSTR haystack, BSTR needle, int32_t from0, int32_t compare);
int32_t vb6_TextMatchAt(BSTR haystack, int32_t at0, BSTR needle, int32_t compare);
// <vbeclipse>: 带 compare 的 InStr (codegen 在实参是 4 个、或 3 个而首参是字符串时用它)
int32_t vb6_InStrC(int32_t start, BSTR haystack, BSTR needle, int32_t compare);
int32_t vb6_InStrVarC(int32_t start, BSTR haystack, vb6_VARIANT needle, int32_t compare);
// Fix 093a: InStrB — 字节版 InStr. 实参可为 Byte() 一维数组 (vb6_SafeArray1D*)
// 或 BSTR; 返回 1 基字节位置 (0=未找到). 此前 RTL 无此符号 → LNK2019.
int32_t vb6_InStrB(int32_t start, void* haystack, void* needle);
// Fix <vbeclipse>: B 族字节串函数 (实现 vb6rtl_string.c / vb6rtl_file.c)。
// 载体口径与 AscB/ChrB 一致: 一个 wchar 存一个字节 (0-255), "按字节"="按 wchar 计数"。
// InputB 在 vb6rtl_file.c (与 InputString 同处, 读原始字节不经文本解码)。
// 两个 Input* 形参序都是 (count, filenumber), 与 cgen 发射的 VB6 源码序一致。
BSTR vb6_LeftB(BSTR s, int32_t n);
BSTR vb6_RightB(BSTR s, int32_t n);
BSTR vb6_MidB(BSTR s, int32_t start, int32_t len);
BSTR vb6_InputB(int32_t count, int32_t filenumber);
BSTR vb6_UCase(BSTR s);
BSTR vb6_LCase(BSTR s);
BSTR vb6_Trim(BSTR s);
BSTR vb6_LTrim(BSTR s);
// Fix 160w: vb6_LTrim 实参是 vb6_VARIANT 时 (如 Common.c:530 `LTrim(SplitToArray(..)[i])`
// — 下标读出的元素是 Variant → 生成代码裸发 vb6_LTrim(vb6_VARIANT) → C2440
// "无法从 vb6_VARIANT 转换到 BSTR"). 同款 _Generic 分派 (Fix 158s 先例):
// Variant → vb6_LTrimVar (先 vb6_VariantToString 解包), 其余原样. 宏在自身展开
// 期间被禁用, 故 default: 指向真函数; vb6rtl_string.c 的 vb6_LTrim 定义处 #undef.
BSTR vb6_RTrim(BSTR s);
BSTR vb6_Chr(int32_t code);
int32_t vb6_Asc(BSTR s);
double vb6_Val(BSTR s);
// 账 #416: 严格转换族 (CByte/CInt/CLng/CSng/CDbl/CCur/... 的字符串实参) 的那一句解析 ——
// 政策交给 oleaut32, 解析不出就报 13; Val 的宽松口径与它刻意不同, 别合。
double vb6_NumVal(BSTR s);
// 账 #418 (§B192): 「认不认作一个数」与「认不出报几号」是两问，拆成两层 ⇒ 全仓那条
// oleaut32 数值解析站点仍只有一处（住在 vb6_NumTryParse 里），而 IsNumeric 问的是前者。
int32_t vb6_NumTryParse(BSTR s, double* out);
int32_t vb6_IsNumericText(BSTR s);   // = vb6_NumTryParse 的成功本身，**不许报错**
int16_t vb6_BoolVal(BSTR s);      // 账 #416: Boolean 那一档问的是 VarBoolFromStr ("True" 算话)
BSTR vb6_Str(int32_t n);
BSTR vb6_Format(vb6_VARIANT expr, BSTR fmt, int32_t firstDayOfWeek, int32_t firstWeekOfYear);
// 账 #428 (§B202): 这枚出口被内置表点名，cgen 把它**直接发进用户 TU**。头里没声明时 C 只当它「返回 int 的隐式声明」—— 两台都照编照链，所以这一格从形状上看不见: x86 把截断后的低 32 位当指针用恰好还通，x64 交出去的就是废指针 ⇒ vb6_BSTR_AssignMove 一解引用就 AV(0xC0000005)。判据里那枚 VarType 针把它在两台上都钉住: 缺声明时答 3 (VT_I4)，有了才答 8 (VT_BSTR)。
BSTR vb6_FormatDateTime(double dateSerial, int32_t namedFormat);

#ifndef __cplusplus
#define vb6_LTrim(x) _Generic((x), \
    vb6_VARIANT: vb6_LTrimVar, \
    default: vb6_LTrim)((x))
#endif
BSTR vb6_LTrimVar(vb6_VARIANT s);

// 消息框
// 账 #490 §B178: 原型收满 VB6 的五参形 (prompt, buttons, title, helpfile, context),
//   与 vb6_InputBox 同款 —— 那一位早就是七参全签名。改前是三参, 而发码对
//   `MsgBox(msg, Style, Title, Help, Ctxt)` (语料 10 处) 递五枚 ⇒ 本机 cl 不报的静默多递。
//   helpfile/context 今天照收、运行期不用 (MessageBoxW 没有 Help 按钮, F1 也不接)。
int32_t vb6_MsgBox(BSTR prompt, int32_t buttons, BSTR title, BSTR helpfile, int32_t context);
// P7.7: single-arg convenience (MsgBox "text" -> vb6_MsgBox1(text))
static inline int32_t vb6_MsgBox1(BSTR prompt) {
    return vb6_MsgBox(prompt, 0, NULL, NULL, 0);
}

// 数值函数
double vb6_Abs(double x);
int32_t vb6_Sgn(double x);
double vb6_Sqr(double x);
double vb6_Round(double x, int32_t decimals);
// Fix 117: VB6 默认数值→字符串 (最短往返表示; Single 最多 7 位有效数字)
BSTR vb6_NumToBSTRDefault(double v, int isSingle);
float vb6_Rnd(int32_t seed);

// 转换函数
int16_t vb6_CInt(double x);
int32_t vb6_CLng(double x);
double vb6_CDbl(double x);
BSTR vb6_CStr(vb6_VARIANT x);
// M22: typed CStr overloads
BSTR vb6_CStrLong(int32_t x);
BSTR vb6_CStrLongFromVariant(vb6_VARIANT x);   // Fix 158q: _Generic 兜底的 Variant 解包入口
// Fix 084m: LongLong → String (64 位, 不截断; vb6_Format 不认 VT_I8, 故独立实现)
BSTR vb6_CStrLongLong(int64_t x);
// ai/032: 无符号整型 (ULong/ULongLong) → String, 按**无符号十进制**输出。
// 为什么另开两条而不复用 vb6_CStrLong/LongLong: 那两条是有符号口径 (%d/%lld),
// CStr(&HFFFFFFFF As ULong) 会打成 "-1"。与 Fix 084m 另开 LongLong 同一处理由。
BSTR vb6_CStrULong(uint32_t x);
BSTR vb6_CStrULongLong(uint64_t x);
BSTR vb6_CStrDbl(double x);
BSTR vb6_CStrDblFromBSTR(BSTR s);      // vbeclipse: BSTR 实参 (后期绑定 COM 数值属性读) → vb6_Val 解析
BSTR vb6_CStrDblFromVariant(vb6_VARIANT v);  // Fix 158q 同款: _Generic 的 Variant 解包入口
// Fix 117c: Single 专用 (VT_R4, 7 位有效数字 + 最短往返)
BSTR vb6_CStrSingle(float x);
BSTR vb6_CStrBool(int16_t x);
BSTR vb6_CStrByte(uint8_t x);
BSTR vb6_CStrDate(double x);
// P8.4: Variant版转换函数
// Fix 090d: 前置声明 (定义在下方"类型转换"区, 供 inline V 变体引用)
int16_t vb6_CBool(double v);
uint8_t vb6_CByte(double v);
float vb6_CSng(double v);
// 账 #416: 严格转换族读 Variant 里那串文本时问的那一句。以前这一档直接借
// vb6_VariantToDouble —— 它对 vtBSTR 走 vb6_Val(宽松)，于是同一个 `CInt(x)` 对
// 「x 是 String 变量」报 13、对「x 是装着同一串文本的 Variant」静默给 0，两半答两种。
// 数值/布尔之外的档仍归 vb6_VariantToDouble 那一枚权威，这里只把文本那一档改问严格解析。
static inline double vb6_VariantConvNum(vb6_VARIANT v) {
    return (v.vt == vb6_vtBSTR) ? vb6_NumVal(v.bstrVal) : vb6_VariantToDouble(v);
}
static inline int16_t vb6_CIntV(vb6_VARIANT v) { return vb6_CInt(vb6_VariantConvNum(v)); }
static inline int32_t vb6_CLngV(vb6_VARIANT v) { return vb6_CLng(vb6_VariantConvNum(v)); }
static inline double vb6_CDblV(vb6_VARIANT v) { return vb6_VariantConvNum(v); }
static inline uint8_t vb6_CByteV(vb6_VARIANT v) { return vb6_CByte(vb6_VariantConvNum(v)); }
static inline float vb6_CSngV(vb6_VARIANT v) { return (float)vb6_VariantConvNum(v); }
static inline int16_t vb6_CBoolV(vb6_VARIANT v) {
    if (v.vt == vb6_vtBoolean) return v.boolVal;
    if (v.vt == vb6_vtBSTR) return vb6_BoolVal(v.bstrVal);
    return vb6_CBool(vb6_VariantToDouble(v));
}

// ai/032: VB.NET 的无符号/窄整型转换函数 (阶梯 8/16/32/64):
//   CSByte → SByte(8 有符号) · CUInt → UInteger(16 无符号)
//   CULng  → ULong(32 无符号) · CULngLng → ULongLong(64 无符号)
// 与 CInt/CLng/CByte 同口径: 收 double, 由调用点的 _Generic/NumVal 分诊 BSTR/Variant,
// 并沿用 vb6_OvfChk 的 VB6 语义 (越界是 Error 6, 不是静默截断)。
int8_t   vb6_CSByte(double x);
uint16_t vb6_CUInt(double x);
uint32_t vb6_CULng(double x);
uint64_t vb6_CULngLng(double x);
static inline int8_t   vb6_CSByteV(vb6_VARIANT v) { return vb6_CSByte(vb6_VariantConvNum(v)); }
static inline uint16_t vb6_CUIntV(vb6_VARIANT v)  { return vb6_CUInt(vb6_VariantConvNum(v)); }
static inline uint32_t vb6_CULngV(vb6_VARIANT v)  { return vb6_CULng(vb6_VariantConvNum(v)); }
static inline uint64_t vb6_CULngLngV(vb6_VARIANT v) { return vb6_CULngLng(vb6_VariantConvNum(v)); }
// CLngLng — 阶梯里有符号的 64 位那一档 (CInt/CLng 的 32 位往上数第三级)。
// 与 CLng 同口径: 收 double, 半值远离零 (vb6_FltToLng), 越界按 VB6 语义报 Error 6。
// 用途 (ComCtlsDemo ProgressBar.ctl 的 `#If Win64` 分支): 给
// ITaskBarList3::SetProgressValue(HWND, ULONGLONG, ULONGLONG) 送 64 位值 ——
// VB6 没有 LongLong 类型, 只能用转换函数把值凑成 64 位整数送进 vtable 调用。
int64_t vb6_CLngLng(double x);
static inline int64_t vb6_CLngLngV(vb6_VARIANT v) { return vb6_CLngLng(vb6_VariantConvNum(v)); }
// twinBASIC 形 CLngPtr (doc.twinbasic.vb6.pro .../Conversion/CLngPtr.html): 转 LongPtr
// (x86 = Long / x64 = LongLong), 越界 Error 6 —— 与上面四档同族 (FltToLng + OvfChk)。
intptr_t vb6_CLngPtr(double x);
static inline intptr_t vb6_CLngPtrV(vb6_VARIANT v) { return vb6_CLngPtr(vb6_VariantConvNum(v)); }
// CLngPtr 的 void* 实参档 (StrPtr / 裸 COM 结果的指针值): 精确转指针宽度整数,
// 不给这一档会被 default 的 double 形参拒收 (C2440)。
static inline intptr_t vb6_CLngPtrFromPtr(void* p) { return (intptr_t)(uintptr_t)p; }

// 类型检查
int32_t vb6_IsNumeric(vb6_VARIANT v);
int32_t vb6_IsNull(vb6_VARIANT v);
int32_t vb6_IsEmpty(vb6_VARIANT v);
int32_t vb6_IsObject(vb6_VARIANT v);
int32_t vb6_IsArray(vb6_VARIANT v);
int32_t vb6_IsDate(vb6_VARIANT v);
int32_t vb6_IsError(vb6_VARIANT v);
// P21-09: CVErr — create VT_ERROR Variant
vb6_VARIANT vb6_CVErr(int32_t errorNumber);

BSTR vb6_TypeName(vb6_VARIANT v);
int32_t vb6_VarType(vb6_VARIANT v);

// 内置对象
void vb6_Debug_Print(BSTR s);
void vb6_Debug_PrintInt(int32_t n);
void vb6_Debug_PrintDouble(double d);

// Debug.Print 变参版 (cgen保留接口, 暂未使用)
// 内部使用stdarg, fmt字符: s=BSTR, d=int32_t, f=double
void vb6_DebugOutputFmt(const char* fmt, ...);

// 简化版: 单BSTR输出+换行
void vb6_DebugPrintStr(BSTR s);

// Debug.Print 分项输出: 逐参数输出不换行, 最后统一换行
void vb6_DebugWriteBSTR(BSTR s);       // 输出BSTR片段(不换行)
void vb6_DebugWriteLong(int32_t n);    // 输出整数片段(不换行)
void vb6_DebugWriteDouble(double d);   // 输出浮点片段(不换行)
void vb6_DebugWriteNewline(void);      // 输出换行

// 控制台输出: 宽字符走控制台原生路径 (WriteConsoleW, 与 chcp 无关); 重定向时写 UTF-8 字节。
// 中文 cmd 默认代码页 936, 而 RTL 内部是 UTF-16 —— 直接把宽字符交给 CRT 的窄路径会在
// 非 CJK 代码页/重定向下把中文丢成 '?'。所有"给人看的"控制台输出都该走这两个函数。
void vb6_ConWriteOutW(const wchar_t* s, int len);
void vb6_ConWriteErrW(const wchar_t* s, int len);

// Fix 161: 控制台输入原语 (stdin 侧) —— 与上面的输出侧对称:
//   真控制台(键盘) → ReadConsoleW  (本就是 UTF-16)
//   管道/重定向     → ReadFile 字节 → MultiByteToWideChar(GetConsoleCP() ?: ACP)
// 原 vb6_InputBox 用裸 fgetws(stdin) 会把管道字节当 UTF-16 解析 ⇒ ASCII/中文
// 全部读错(不是显示乱码)。VB6 侧无对应关键字, 供宿主/夹具 Declare 或 RTL 内部调用。
int32_t vb6_ConInputIsConsole(void);  // 1=控制台键盘, 0=管道/重定向
BSTR vb6_ConReadLine(void);           // 读一行(不含行尾); 无输入时返回空 BSTR
BSTR vb6_ConReadKey(void);            // 读一个字符(取行首字符)

// Fix 161: Console 对象 (对齐 twinBASIC) —— VB6 侧 `Console.ReadLine()` 等
void vb6_Console_Write(BSTR s);
void vb6_Console_WriteLine(BSTR s);
BSTR vb6_Console_ReadLine(void);
BSTR vb6_Console_ReadKey(void);

// 整除
int32_t vb6_IntDiv(int32_t a, int32_t b);

// Task #44: VB6 '/' 与 Mod 的字面量 0 除数 → 运行期错误 11
// (常量折叠 C2124 规避 + On Error Resume Next 语义, 见 vb6rtl.c 实现)
double vb6_Num_Div(double a, double b);
int32_t vb6_Num_Mod(int32_t a, int32_t b);

// ai/032: 无符号整除/取余 (ULong=32 位 · ULongLong=64 位)。vb6_IntDiv / vb6_Num_Mod
// 的形参是 int32_t, 无符号操作数会被先按有符号解释 (实测 `&H80000000 As ULong \ 2`
// 得 3221225472 而非 1073741824), 所以这两档必须另走无符号版。除零同样是错误 11。
uint32_t vb6_IntDivULong(uint32_t a, uint32_t b);
uint32_t vb6_Num_ModULong(uint32_t a, uint32_t b);
uint64_t vb6_IntDivULongLong(uint64_t a, uint64_t b);
uint64_t vb6_Num_ModULongLong(uint64_t a, uint64_t b);

// ai/032 rev2: 有符号 64 位整除/取余。`\` / `Mod` 的档位由 CCodeGen::intDivResultType
// 一处给出: 32 位异符号 (如 ULong \ Long) 按 VB.NET 的提升规则要升到 64 位**有符号**
// 再除, 而 vb6_IntDiv 只有 int32_t 形参 —— 少了这两条就只能退回 32 位截断。
// 除零同样是运行期错误 11。
int64_t vb6_IntDivLongLong(int64_t a, int64_t b);
int64_t vb6_Num_ModLongLong(int64_t a, int64_t b);

// C3 扩展 (ai/032): `<<` / `>>` 移位。cgen 一律包成这三个 inline 之一, 而不是直接
// 发 C 的 `a << b`, 因为 C 的移位在三处与 VB.NET 语义不同 —— 其中两处是 UB:
//
//   1) 移位计数必须屏蔽。VB.NET 把计数屏蔽到操作数位宽内, C 对 计数 >= 位宽 是
//      **未定义行为** (变量计数时优化器可变出任意结果)。这里统一按 63 取模 ——
//      取 63 而不是"按左操作数位宽"是有意的: cgen 把结果抬到至少 32 位 (见
//      cgen_expr_binary.cpp 的 castBack), 于是 `1 << 20` 拿到 1048576 而不是被
//      C3 的 16 位 Integer 截成 0, 与 VB.NET 的实测答案一致。
//   2) 有符号左移溢出 (`&H80000000 << 1`) 在 C 里是 UB。这里按 uint64 位模式左移,
//      位模式确定, 再由 cgen 的截断回目标类型 —— 结果与 VB.NET 的 wrap 一致。
//   3) `>>` 的有符号性。VB.NET 对有符号整型是**算术**右移 (补符号位), 对无符号整型
//      (Byte/UInteger/ULong/ULongLong) 是**逻辑**右移 (补 0)。C 的 >> 对有符号数是
//      实现定义, 所以分成 vb6_Shr (有符号) / vb6_ShrU (无符号) 两支, 由 cgen 按
//      左操作数的有符号性选。
static inline int64_t vb6_Shl(int64_t v, int64_t n) {
    return (int64_t)((uint64_t)v << (uint64_t)(n & 63));
}
static inline int64_t vb6_Shr(int64_t v, int64_t n) {
    return v >> (n & 63);   /* 有符号: 算术右移 (MSVC/Clang 均按符号扩展) */
}
static inline int64_t vb6_ShrU(uint64_t v, int64_t n) {
    return (int64_t)(v >> (uint64_t)(n & 63));   /* 无符号: 逻辑右移 */
}

// 幂运算
double vb6_Pow(double base, double exp);

// 对象操作
void* vb6_NewObject(const wchar_t* className);
int32_t vb6_TypeOf(void* obj, const wchar_t* typeName);
void* vb6_DictAccess(void* obj, const wchar_t* key);

// 字符串函数 (补充)
BSTR vb6_Replace(BSTR expr, BSTR find, BSTR rep, int32_t start, int32_t count, int32_t compare);
BSTR vb6_Space(int32_t n);
BSTR vb6_String(int32_t n, int32_t charCode);
int32_t vb6_StrComp(BSTR s1, BSTR s2, int32_t compare);
BSTR vb6_StrReverse(BSTR s);
int32_t vb6_InStrRev(BSTR haystack, BSTR needle, int32_t start, int32_t compare);
BSTR vb6_LCase_str(BSTR s);  // LCase$别名
BSTR vb6_UCase_str(BSTR s);
int16_t vb6_Like(BSTR source, BSTR pattern);  // Like运算符
// <vbeclipse>: 带比较模式的 Like (mode: 0=vbBinaryCompare, 非0=文本)。
// `Option Compare Text` 模块里 codegen 发改写调用点到这里。
int16_t vb6_LikeC(BSTR source, BSTR pattern, int32_t mode);
// <vbeclipse>: 文本比较的唯一口径 (CompareStringW 区域语言序, 见 vb6rtl_string.c);
// vb6_StrComp 的 vbTextCompare 形与 Option Compare Text 模块的运算符共用它。
int vb6_TextCmp(const wchar_t* a, const wchar_t* b);

// ============================================================
// P14.2.2: 系统函数
// ============================================================
BSTR vb6_Dir(BSTR pathname, int32_t attributes);
BSTR vb6_CurDir(BSTR drive);
int32_t vb6_Shell(BSTR pathname, int32_t windowstyle);
BSTR vb6_Environ(BSTR envstring);
BSTR vb6_Command(void);










// 数学函数 (补充)
double vb6_Sin(double x);
double vb6_Cos(double x);
double vb6_Tan(double x);
double vb6_Atn(double x);
double vb6_Log(double x);
double vb6_Exp(double x);
double vb6_Fix(double x);
double vb6_Int(double x);
void vb6_Randomize(double seed);
float vb6_Rnd_Full(int32_t seed);

// 日期时间函数
double vb6_Now(void);
double vb6_Date(void);
double vb6_Time(void);
void vb6_DateSet(BSTR dateStr);
void vb6_TimeSet(BSTR timeStr);
int32_t vb6_Year(double date);
int32_t vb6_Month(double date);
int32_t vb6_Day(double date);
int32_t vb6_Hour(double time);
int32_t vb6_Minute(double time);
int32_t vb6_Second(double time);

// P14.2.4: DateAdd/DateDiff/DatePart/DateSerial
double vb6_DateSerial(int32_t year, int32_t month, int32_t day);
// P14.2.4: DateAdd/DateDiff/DatePart
double vb6_DateAdd(BSTR interval, double number, double date);
int64_t vb6_DateDiff(BSTR interval, double date1, double date2, int32_t firstDayOfWeek, int32_t firstWeekOfYear);
int32_t vb6_DatePart(BSTR interval, double date, int32_t firstDayOfWeek, int32_t firstWeekOfYear);

// P21-B: Weekday/DateValue/TimeSerial/TimeValue
int32_t vb6_Weekday(double date, int32_t firstDayOfWeek);
// 账 #427 (§B201): 「这一年里第几周」这一问的唯一出口 (定义在 vb6rtl_date.c)。
// 读者两位: vb6_DatePart 的「ww」臂与 vb6_Format 的「ww」档 —— 所以它不再是 date.c 的内部静态函数。
int32_t vb6_WeekOfYear(int32_t serial, int32_t fdow, int32_t fwoy);
// 账 #429 (§B203) 第一百零一刀: locale 的那几问在 RTL 里只许一处答案 (定义在 vb6rtl_format.c)。
// 读者跨文件 —— vb6rtl_system.c 的 vb6_FormatDateTime、vb6rtl_conv.c 的 vb6_CStrDate、Format 家族自己，
// 所以它们不是 file-static。「general date 什么时候该带长时间」这一条也只住在最后那条出口里。
void vb6_fmtLocaleShortDate(const SYSTEMTIME* st, wchar_t* out, size_t cap);
void vb6_fmtLocaleLongDate(const SYSTEMTIME* st, wchar_t* out, size_t cap);
void vb6_fmtLocaleLongTime(const SYSTEMTIME* st, wchar_t* out, size_t cap);
void vb6_fmtLocaleGeneralDate(const SYSTEMTIME* st, wchar_t* out, size_t cap);
double vb6_DateValue(BSTR dateStr);
double vb6_TimeSerial(int32_t hour, int32_t minute, int32_t second);
double vb6_TimeValue(BSTR timeStr);
// 账 #417 (§B191): 文本→日期序列的解析只有 vb6rtl_date.c 里那一条 VarDateFromStr 站点,
// 「认不出算几号」由下面两位主人各按自己的档答 (全是 13, 但**认不认纯数字串**两档不同):
//   vb6_DateParse = DateValue / TimeValue / CStr(日期形文本) 那一档, 只认日期时间写法;
//   vb6_DateVal   = CDate 那一档 = 日期写法 ∪ 严格数值 (vb6_NumVal), 且保留时间小数。
// 实测口径见 .build/k234_cdate.vbs / k235_datescan.vbs / k236_datefail.vbs。
double vb6_DateParse(BSTR s);
double vb6_DateVal(BSTR s);
// 账 #418 (§B192): 同一句问「像不像」而不是「折出几号」—— IsDate 的文本档问这一枚，
// 它**不报错**（判据不是一次转换），收下集合就是那条唯一解析站点的成功集。
int32_t vb6_DateCanParse(BSTR s);

// P14.3.4: App全局对象属性
BSTR vb6_App_Path(void);     // App.Path - EXE所在目录
BSTR vb6_App_EXEName(void);  // App.EXEName - EXE文件名(不含扩展名)
intptr_t vb6_App_hInstance(void); // App.hInstance - 模块实例句柄 (Fix 179: 必须是指针宽度)
BSTR vb6_App_HelpFile(void); // App.HelpFile - 帮助文件名(无App COM对象, 返回空串)
// 账 #305 第320刀 (账 #592): App.Title 的运行期真值只有一格 (vb6rtl_system.c) ——
// 生成的入口点在 vb6_Init 之后把 .vbp 的 Title 播种进来; `App.Title = v` (VB6 允许运行期改)
// 也只走这一枚 setter。setter **借入** (内部复制一份, 调用方不必交出所有权)。
void vb6_SetAppTitle(const wchar_t* title);
BSTR vb6_App_Title(void);                  // App.Title 读取 (返回副本; 未播种时退 EXEName)
const wchar_t* vb6_AppTitleBorrow(void);   // 内部: MsgBox 默认题头 (借用, 勿释放; 未播种 = NULL)
// 账 #305 第二百六十四刀: App.PrevInstance 的定义在 vb6rtl_system.c, 此前**任何头里都没有声明**
// ⇒ 生成代码那发走 C4013 隐式声明。这一枚恰好返回 int32_t, 隐式「extern 返回 int」同型,
// 所以两架构都出不了错 —— 但它是 §B202 那条规则该抓而没抓到的活样本, 补上以自证规则问对了问题。
int32_t vb6_App_PrevInstance(void); // App.PrevInstance - 本程序的另一个实例是否在跑

// P18-C: Clipboard 对象
void   vb6_Clipboard_SetText(BSTR text);
BSTR   vb6_Clipboard_GetText(void);
void   vb6_Clipboard_Clear(void);
void   vb6_Clipboard_SetData(void* pPicture, int32_t format);  // format: 0=auto, 2=vbCFBitmap, 14=vbCFEMetafile
int32_t vb6_Clipboard_GetFormat(int32_t format);  // 1=vbCFText, 2=vbCFBitmap, etc.

// P18-C: Screen 对象
int32_t vb6_Screen_Width(void);      // Screen.Width (twips)
int32_t vb6_Screen_Height(void);     // Screen.Height (twips)
int32_t vb6_Screen_MouseX(void);     // Mouse position X (twips)
int32_t vb6_Screen_MouseY(void);     // Mouse position Y (twips)
void*  vb6_Screen_ActiveControl(void);  // Active control HWND
void*  vb6_Screen_ActiveForm(void);     // Active form HWND
int32_t vb6_Screen_TwipsPerPixelX(void);
int32_t vb6_Screen_TwipsPerPixelY(void);
// Fix 161f: Screen.MousePointer 全局读/写 (非控件级)
int32_t vb6_Screen_MousePointer(void);
void    vb6_Screen_SetMousePointer(int32_t pointer);

// P18-C: Printer 对象
void   vb6_Printer_Print(BSTR text);
void   vb6_Printer_EndDoc(void);
void   vb6_Printer_NewPage(void);
int32_t vb6_Printer_Width(void);
int32_t vb6_Printer_Height(void);
int32_t vb6_Printer_CurrentX(void);
int32_t vb6_Printer_CurrentY(void);
void   vb6_Printer_SetCurrentX(int32_t x);
void   vb6_Printer_SetCurrentY(int32_t y);
// Task #39: Printer / Printers 内置全局对象 (cDlg.cls ShowFont / IsPrinter)
// vb6_Printer_Object: 默认打印机 DC 指针作哨兵 (懒创建, 无打印机时 NULL →
//   `Printer Is Nothing` 判真, 调用方自然跳过 hDC 取用);
// vb6_Printer_hDC: HDC 句柄 (Printer.hDC, ChooseFont.hDC 等 API 用);
// vb6_Printers_Collection: 空 RTL Collection (vb6_ForEach_Init 原生支持,
//   For Each 循环零次; 真实 EnumPrinters 枚举属后续增强)。
void*  vb6_Printer_Object(void);

// Fix <vbeclipse>: vb6_Screen_Object — Screen 全局对象哨兵 (NULL)。
// .Width/.MouseX 等已登记成员在 MemberAccess 分支直接发 vb6_Screen_Xxx();
// 未登记成员赋值落 NULL 槽语义, 无害。
void*  vb6_Screen_Object(void);
extern void* vb6_Screen_MouseIcon;         // Fix <vbeclipse>: Screen.MouseIcon 读写槽
void*  vb6_Printer_hDC(void);
void*  vb6_Printers_Collection(void);

// P18-C: Forms 集合
int32_t vb6_Forms_Count(void);
void*  vb6_Forms_Item(int32_t index);  // 0-based
void   vb6_Forms_Register(void* hwnd, const char* name);   // 窗体创建时注册（账 #257: 带上 VB 模块名，宿主模型的 .Name 从此有数可问）
void   vb6_Forms_Unregister(void* hwnd); // 窗体销毁时注销
void   vb6_Forms_LoopDepth(int delta);   // Fix 188: 消息循环进出 (最后一个窗体卸载才投 WM_QUIT)
void*  vb6_Forms_GetActive(void);        // Fix 146: 当前活动窗体 (Screen.ActiveForm)

// P14.2.4: IIf / InputBox
BSTR vb6_IIfBSTR(int32_t cond, BSTR truepart, BSTR falsepart);
int32_t vb6_IIfLong(int32_t cond, int32_t truepart, int32_t falsepart);
double vb6_IIfDouble(int32_t cond, double truepart, double falsepart);
vb6_VARIANT vb6_IIfVariant(int32_t cond, vb6_VARIANT truepart, vb6_VARIANT falsepart);
BSTR vb6_InputBox(BSTR prompt, BSTR title, BSTR defaultstr, int32_t xpos, int32_t ypos, BSTR helpfile, int32_t context);

// 类型转换 (补充)
int16_t vb6_CBool(double v);
uint8_t vb6_CByte(double v);
// 溢出检查收窄 (VB6 Error 6): 值越界时 vb6_RaiseError(6, "Overflow"),
// 可被 On Error 捕获; 范围内原样窄化。ai/009 §5.10 P3。
uint8_t vb6_ChkByte(int32_t v);
int16_t vb6_ChkInt(int32_t v);
int32_t vb6_ChkLong(int64_t v);
// 账 #248: 浮点交给整数目标时**先取整**这一件事只有一个出口 (CLng/CInt 与隐式赋值同源)。
// C 的参数转换是截断，发码把 `vb6_ChkLong(<double>)` 直接递进去就丢了 VB6 的取整。
int64_t vb6_FltToLng(double x);
float vb6_CSng(double v);
double vb6_CDate(vb6_VARIANT v);
BSTR vb6_Hex(int32_t n);
BSTR vb6_Oct(int32_t n);

// P18-A: 兼容性填平 — 新增RTL函数
double vb6_CCur(double v);             // Fix 126: CCur 值语义 (4 位小数)
vb6_VARIANT vb6_CDec(vb6_VARIANT v);     // P20-07: CDec返回真实DECIMAL (vt=14)
long vb6_RGB(int32_t r, int32_t g, int32_t b);  // RGB: OLE color
long vb6_QBColor(int32_t n);        // QBColor: 16-color lookup
double vb6_FileDateTime(BSTR pathname); // FileDateTime: 文件修改时间→VB6 date serial
int32_t vb6_FileLen(BSTR pathname);    // FileLen: 文件大小(字节)
// <vbeclipse>: GetAttr/SetAttr 只有定义没有声明 (vb6rtl_misc.c:99/107) ⇒ 生成的 C 走
// 隐式 int 声明, 每个用到它们的工程都吃一条 cl C4013; 补上原型消除该族警告。
int32_t vb6_GetAttr(BSTR pathname);    // GetAttr: 文件属性 (GetFileAttributesW)
void    vb6_SetAttr(BSTR pathname, int32_t attributes); // SetAttr: 设置文件属性
void   vb6_SendKeys(BSTR keys, int32_t wait);    // SendKeys: 发送按键
void   vb6_AppActivate(BSTR title, int32_t wait); // AppActivate: 激活窗口(标题或数字PID)
void   vb6_AppActivateByPid(int32_t pid, int32_t wait); // AppActivate: 按进程ID激活窗口
void   vb6_MidSet(BSTR* target, int32_t start, int32_t len, BSTR replacement); // Mid$ statement赋值

// P18-D: 兼容性填平 — 字符串/指针/格式化函数
int32_t vb6_AscW(BSTR s);           // AscW: Unicode code point
BSTR   vb6_ChrW(int32_t code);      // ChrW: Unicode character
int32_t vb6_AscB(BSTR s);           // AscB: first byte value
BSTR   vb6_ChrB(int32_t code);      // ChrB: single-byte string
double  vb6_Timer(void);            // Timer: seconds since midnight (fractional)
BSTR   vb6_StrConv(BSTR text, int32_t conversion, int32_t localeID); // StrConv
struct vb6_SafeArray1D; // forward declaration
// StrConv 字节数组语义 (twinbasic 兼容):
//   vb6_StringToByteArray    : String -> Byte()，复制 BSTR 原始 UTF-16LE 内存字节
//   vb6_StrConvToByteArray   : StrConv(s, 128/64) -> Byte()
//                               128 (vbFromUnicode): Unicode -> 系统 ANSI/GBK 字节
//                               64  (vbUnicode)    : 原始 BSTR 内存按 ANSI 解码后重建 UTF-16LE 字节
//   vb6_StrConvFromByteArray : 反向, Byte() -> BSTR (供 s = StrConv(ba, vbUnicode))
struct vb6_SafeArray1D* vb6_StringToByteArray(BSTR text);
struct vb6_SafeArray1D* vb6_StrConvToByteArray(BSTR text, int32_t conversion, int32_t localeID);
BSTR   vb6_StrConvFromByteArray(struct vb6_SafeArray1D* arr, int32_t conversion, int32_t localeID);
// Fix 140: Byte() -> BSTR (UTF-16LE 字节直读) 与 Byte() 数据指针.
//   vb6_ByteArrayToString(arr): 把 arr->data 当作 UTF-16LE 宽字符串还原为 BSTR.
//     与 vb6_StringToByteArray 互逆 (原始内存字节, 非 ANSI 转码).
//   vb6_SafeArrayDataPtr(arr) : StrPtr(Byte()) 语义 — 返回 arr->data (非结构体指针).
BSTR   vb6_ByteArrayToString(struct vb6_SafeArray1D* arr);
void*  vb6_SafeArrayDataPtr(struct vb6_SafeArray1D* arr);
struct vb6_SafeArray1D* vb6_Filter(struct vb6_SafeArray1D* source, BSTR match, int32_t include, int32_t compare);
void*    vb6_StrPtr(BSTR s);          // StrPtr: address of string data
// 账 #305 §B180 (第六十九刀): StrPtr 收**整枚槽**那一档 —— vt 已是串时**借**槽里那枚 BSTR，
// 不再像 `vb6_StrPtr(vb6_VariantToString(v))` 那样每调用现造一枚没人回收的临时串。
// 发码侧由 arg_variant (撤回解包) 与 builtin_fixup (改 callee 名) 两站同归一处新问句决定，
// 只有「实参本身是槽」时才走这一枚 (理由与读数见 §B180 —— 无条件按位号答槽会打坏 1679 处)。
void*    vb6_StrPtrVar(vb6_VARIANT v);
uintptr_t vb6_ObjPtr(void* obj);       // ObjPtr: address of object
BSTR   vb6_LSet(BSTR str, int32_t length);  // LSet: left-justify
BSTR   vb6_RSet(BSTR str, int32_t length);  // RSet: right-justify
// 账 #118: 定长串赋值 (`Dim s As String * N: s = 值`) 的收口函数 —— LSet 同款
// (不足右侧补空格/超长截右), 顺带释放入参。给发码侧"值本身是自有临时串"那一侧用:
// 定长串赋值必须先 fit 成一只**新** BSTR 再 move 进目标, 裸 LSet 会把 RHS 临时串漏掉。
BSTR   vb6_LSetFree(BSTR str, int32_t length);
BSTR   vb6_WeekdayName(int32_t weekday, int32_t abbreviate, int32_t firstDayOfWeek);
BSTR   vb6_MonthName(int32_t month, int32_t abbreviate);
BSTR   vb6_FormatCurrency(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits);
BSTR   vb6_FormatNumber(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits);
BSTR   vb6_FormatPercent(double value, int32_t numDigits, int32_t incLeading, int32_t useParens, int32_t groupDigits);

// P18-E: 兼容性填平 — 金融函数+文件锁定+Partition
double vb6_SLN(double cost, double salvage, double life);
double vb6_SYD(double cost, double salvage, double life, double period);
double vb6_DDB(double cost, double salvage, double life, double period, double factor);
double vb6_FV(double rate, double nper, double pmt, double pv, int32_t type);
double vb6_PV(double rate, double nper, double pmt, double fv, int32_t type);
double vb6_Pmt(double rate, double nper, double pv, double fv, int32_t type);
double vb6_IPmt(double rate, double per, double nper, double pv, double fv, int32_t type);
double vb6_PPmt(double rate, double per, double nper, double pv, double fv, int32_t type);
double vb6_RATE(double nper, double pmt, double pv, double fv, int32_t type, double guess);
double vb6_NPV(double rate, struct vb6_SafeArray1D* values);
void   vb6_Lock(int32_t filenum, int64_t start, int64_t end);
void   vb6_Unlock(int32_t filenum, int64_t start, int64_t end);
void   vb6_Reset(void);
BSTR   vb6_Partition(int64_t number, int64_t start, int64_t stop, int64_t interval);

// ============================================================
// Fix 158q: CStrLong/CLng/CDbl/CInt/CCur 的 Variant·BSTR 实参分派
// ============================================================
// 生成代码里会出现 vb6_CStrLong(<vb6_VARIANT>) (With 后端 COM 属性
// vb6_VariantFromComResult(vb6_ComGetProp(...)) 被当 Long 转 BSTR → C2440
// "无法从 vb6_VARIANT 转换到 int32_t"), 以及 vb6_CLng/vb6_CDbl/vb6_CInt/
// vb6_CCur 收 vb6_VARIANT 或 BSTR 实参 (UDT 单元格 .Text → C2440 BSTR→double).
// 用 _Generic (MSVC 2019 16.8+, /std:c11) 在**调用点**分派:
//   vb6_VARIANT → 既有 vb6_*V 解包版; BSTR → 新增 vb6_*BSTR 解析版; 其余原样.
// 位置必须在上面所有 vb6_CInt/CLng/CDbl/CCur 声明**之后**: 宏一旦定义, 其后
// 出现同名声明/定义都会被改写 (头内 79~81 行的 CIntV/CLngV 内联体因此已在
// 宏之前解析, 走的是真函数, 不会自我展开). 各 .c 的定义处用 #undef 自行关闭
// (vb6rtl_conv.c / vb6rtl_misc.c), 同 vb6_LenB (Fix 048) 的既有先例.
int32_t vb6_CLngBSTR(BSTR s);
double vb6_CDblBSTR(BSTR s);
int16_t vb6_CIntBSTR(BSTR s);
double vb6_CCurBSTR(BSTR s);
// ComCtlsDemo: 后期绑定 COM 的 **Date** 属性 (DTPicker / MonthView 的 .MinDate/.MaxDate)
// 在「写进 Text (BSTR)」语境下被 codegen 读成 vb6_ComGetStringProp (BSTR) 再套
// vb6_CStrDate, 而本函数原签名只收 double → C2440 "无法从 wchar_t* 转换为 double"
// (PPDTPickerGeneral.c 208/209、PPMonthViewGeneral.c 207/208)。
// 与上面 vb6_CStrDbl 的 BSTR 档同款: 补一个 BSTR 入口, 由 _Generic 按实参 C 类型分派。
// BSTR 是日期串, 故这里走 VarDateFromStr 解析 (同 vb6_DateValue 那条口径),
// 而不是 vb6_CStrDblFromBSTR 的 vb6_Val —— 后者是数值解析, 对日期串答 0。
BSTR vb6_CStrDateFromBSTR(BSTR s);
// ai/032: 新增四档转换函数的 BSTR 入口 —— 同 Fix 158r 的 CLngBSTR.
// 账 #416: 这一族的解析口径归 vb6_NumVal 那一位主人 (十进制 / "&H"&O" 前缀 / 区域分隔符
// 都由它答, 解析不出报 13); cgen 那侧的表达式实参走同一枚函数, 两条路不再各写一份。
uint64_t vb6_CULngLngBSTR(BSTR s);
uint32_t vb6_CULngBSTR(BSTR s);
uint16_t vb6_CUIntBSTR(BSTR s);
int8_t   vb6_CSByteBSTR(BSTR s);
intptr_t vb6_CLngPtrBSTR(BSTR s);   // twinBASIC 形 CLngPtr 的字符串入口 (同族口径)
// CLngLng (阶梯里有符号 64 位那一档) 的字符串入口 —— 与上面四档同口径, 走 vb6_Val。
int64_t  vb6_CLngLngBSTR(BSTR s);
// ai/032: CStrULong/CStrULongLong 的 Variant 解包入口 (与 Fix 158q 同款)
BSTR vb6_CStrULongFromVariant(vb6_VARIANT x);
BSTR vb6_CStrULongLongFromVariant(vb6_VARIANT x);
static inline double vb6_CCurV(vb6_VARIANT v) { return vb6_CCur(vb6_VariantConvNum(v)); }
// Fix 160w: CLng(指针) 分派目标 — 截断为 int32_t (VB6 Long = 地址低半).
// (更名 vb6_CLngFromPtr: 原 vb6_CLngPtr 这个名字让给下方 CLngPtr 转换函数 ——
//  与 VB 内置同名; 旧名只在本头两处出现, 一处定义一处 _Generic。)
static inline int32_t vb6_CLngFromPtr(void* p) { return (int32_t)(intptr_t)p; }
#ifndef __cplusplus
#define vb6_CStrLong(x) _Generic((x), \
    vb6_VARIANT: vb6_CStrLongFromVariant, \
    default: vb6_CStrLong)((x))
// vbeclipse: CStrDbl 补 BSTR/Variant 分派 (默认分支不变, 老调用点行为一致).
// 背景: 后期绑定 COM 数值属性在拼接语境下被 codegen 读成 vb6_ComGetStringProp
// (BSTR), 而 vb6_CStrDbl 原签名只收 double → C2440 (ucPerspective.c 3333/3384/
// 3453/3504).
#define vb6_CStrDbl(x) _Generic((x), \
    BSTR: vb6_CStrDblFromBSTR, \
    vb6_VARIANT: vb6_CStrDblFromVariant, \
    default: vb6_CStrDbl)((x))
// ComCtlsDemo: CStrDate 补 BSTR 分派 (默认分支不变, 既有 double 调用点行为一字不改).
// 背景同 vb6_CStrDbl: 后期绑定 COM 的 Date 属性在 BSTR 语境下被读成
// vb6_ComGetStringProp (BSTR) 再套 vb6_CStrDate → C2440. 这一档此前是
// CStr 家族里唯一没有 BSTR 入口的 (Long/Byte 因只有 warning 而漏过).
#define vb6_CStrDate(x) _Generic((x), \
    BSTR: vb6_CStrDateFromBSTR, \
    default: vb6_CStrDate)((x))
#define vb6_CLng(x) _Generic((x), \
    vb6_VARIANT: vb6_CLngV, \
    BSTR: vb6_CLngBSTR, \
    void*: vb6_CLngFromPtr, \
    default: vb6_CLng)((x))
#define vb6_CDbl(x) _Generic((x), \
    vb6_VARIANT: vb6_CDblV, \
    BSTR: vb6_CDblBSTR, \
    default: vb6_CDbl)((x))
#define vb6_CInt(x) _Generic((x), \
    vb6_VARIANT: vb6_CIntV, \
    BSTR: vb6_CIntBSTR, \
    default: vb6_CInt)((x))
#define vb6_CCur(x) _Generic((x), \
    vb6_VARIANT: vb6_CCurV, \
    BSTR: vb6_CCurBSTR, \
    default: vb6_CCur)((x))
// ai/032: CStrULong/CStrULongLong 与四档新转换函数的 Variant·BSTR 分派。
// 与上面 vb6_CLng 一族同款: 宏只改**调用点**, 各 .c 的定义处自带 #undef。
#define vb6_CStrULong(x) _Generic((x), \
    vb6_VARIANT: vb6_CStrULongFromVariant, \
    default: vb6_CStrULong)((x))
#define vb6_CStrULongLong(x) _Generic((x), \
    vb6_VARIANT: vb6_CStrULongLongFromVariant, \
    default: vb6_CStrULongLong)((x))
#define vb6_CSByte(x) _Generic((x), \
    vb6_VARIANT: vb6_CSByteV, \
    BSTR: vb6_CSByteBSTR, \
    default: vb6_CSByte)((x))
#define vb6_CUInt(x) _Generic((x), \
    vb6_VARIANT: vb6_CUIntV, \
    BSTR: vb6_CUIntBSTR, \
    default: vb6_CUInt)((x))
#define vb6_CULng(x) _Generic((x), \
    vb6_VARIANT: vb6_CULngV, \
    BSTR: vb6_CULngBSTR, \
    default: vb6_CULng)((x))
#define vb6_CULngLng(x) _Generic((x), \
    vb6_VARIANT: vb6_CULngLngV, \
    BSTR: vb6_CULngLngBSTR, \
    default: vb6_CULngLng)((x))
#define vb6_CLngLng(x) _Generic((x), \
    vb6_VARIANT: vb6_CLngLngV, \
    BSTR: vb6_CLngLngBSTR, \
    default: vb6_CLngLng)((x))
// twinBASIC 形 CLngPtr: 三档分派 + void* 档 (指针值精确转, 见 vb6_CLngPtrFromPtr)。
// 返回 intptr_t, 宽度随目标架构 (x86 4 / x64 8)。
#define vb6_CLngPtr(x) _Generic((x), \
    vb6_VARIANT: vb6_CLngPtrV, \
    BSTR: vb6_CLngPtrBSTR, \
    void*: vb6_CLngPtrFromPtr, \
    default: vb6_CLngPtr)((x))
#endif

// Fix 158s: vb6_InStr 的 needle 实参是 vb6_VARIANT 时 (For 循环里对每个 Buffer 判
// 定包含关系, Value 形参是 Variant) → 生成代码裸发 vb6_InStr(1, Buffer, Value),
// VBFlexGrid.c 一系 C2440 "无法从 vb6_VARIANT 转换到 BSTR"。同款 _Generic 分派:
// Variant → vb6_InStrVar (先 vb6_VariantToString 解包), 其余原样。两侧返回类型同为
// int32_t, 分派自洽。宏在自身展开期间被禁用 (与 vb6_CLng 同技巧), 故 default: 指向
// 真函数, 无递归; vb6rtl_string.c 的定义处自带 #undef。
int32_t vb6_InStrVar(int32_t start, BSTR haystack, vb6_VARIANT needle);
#ifndef __cplusplus
#define vb6_InStr(start, haystack, needle) _Generic((needle), \
    vb6_VARIANT: vb6_InStrVar, \
    default: vb6_InStr)((start), (haystack), (needle))
#endif

// <vbeclipse>: 4 参形的同款分派 (Fix 158s 的机制, 只是多一个 compare 实参) ——
// InStr 的"字符串优先三参形"与四参形都发到这里, needle 是 Variant 时先解包。
#ifndef __cplusplus
#define vb6_InStrC(start, haystack, needle, compare) _Generic((needle), \
    vb6_VARIANT: vb6_InStrVarC, \
    default: vb6_InStrC)((start), (haystack), (needle), (compare))
#endif

#ifdef __cplusplus
}
#endif
