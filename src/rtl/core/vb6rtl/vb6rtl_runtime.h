#pragma once
// vb6rtl_runtime.h - 运行时初始化/退出、ParamArray 支持、Option Compare、运行期补充
// 必须排在 array / builtin 之后（依赖 SafeArray 与 DebugPrintStr 声明）
#include "vb6rtl_base.h"

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 运行时初始化/退出
// ============================================================

// ============================================================
// ParamArray runtime support (P14.1.5)
// ============================================================

SAFEARRAY* vb6_PA_Create(int32_t count);
void vb6_PA_Destroy(SAFEARRAY* psa);
void vb6_PA_SetVariant(SAFEARRAY* psa, int32_t index, VARIANT* pv);
void vb6_PA_SetLong(SAFEARRAY* psa, int32_t index, int32_t val);
void vb6_PA_SetLongPtr(SAFEARRAY* psa, int32_t index, intptr_t val);  /* Fix 082: x64-safe VarPtr parameter */
void vb6_PA_SetDouble(SAFEARRAY* psa, int32_t index, double val);
void vb6_PA_SetBSTR(SAFEARRAY* psa, int32_t index, BSTR val);
vb6_VARIANT vb6_PA_GetVariant(SAFEARRAY* psa, int32_t index);  /* Fix rev37: 返回 vb6_VARIANT 供 _Generic 装箱 */
int32_t vb6_PA_GetLong(SAFEARRAY* psa, int32_t index);
double vb6_PA_GetDouble(SAFEARRAY* psa, int32_t index);
BSTR vb6_PA_GetBSTR(SAFEARRAY* psa, int32_t index);
int32_t vb6_IsMissing(SAFEARRAY* psa);
int32_t vb6_PA_UBound(SAFEARRAY* psa);
int32_t vb6_PA_LBound(SAFEARRAY* psa);

void vb6_Init(void);
// comctl32 通用控件引导 (ProgressBar/StatusBar/Toolbar/ListView/TreeView 依赖)。
// 由 vb6_Init 自己调用; 这里额外声明以便其它 RTL 单元直接使用。
void vb6_ComCtl_Init(void);
void vb6_Exit(void);
void vb6_End(void);
void vb6_Beep(void);

// P18-C: Option Compare (Text/Binary)
extern int g_vb6_optionCompareText;  // 0=Binary(default), 1=Text
int vb6_StrCmp(const wchar_t* a, const wchar_t* b);  // respects Option Compare
// <vbeclipse>: Option Compare 是**按模块**的编译期属性, 所以 codegen 在 Text 模块里
// 直接把字符串比较改发到这条恒文本入口 (而不是靠进程唯一的全局去切语义)。
int vb6_StrCmpT(const wchar_t* a, const wchar_t* b);  // 恒 vbTextCompare 口径

// Fix 048: Missing runtime functions that were generating C4013 warnings

// vb6_SA_Destroy — generic SafeArray destroyer (used by COM class destructor codegen)
// 账 #305 §B220 (第二百二十四刀): "generic" 之前只对 1D 成立 —— 多维定长字段的载体是
// vb6_SafeArrayND* (创建在 cgen_com.cpp 的出厂圈), 硬投成 vb6_SafeArray1D* 在 x86 上把
// data 读成 totalElements (free(0x8) = AV), 在 x64 上则因 data 偏移恰好相同而"看着没事"
// (ND String 那一形静默漏掉全部元素 BSTR)。秩不能看声明: `Dim x()` 再 `ReDim x(1 To 2, 3)`
// 的存储是 1D* 而载体已是 ND ⇒ 只能问运行期 (vb6_SaIs1D 就是为这件事写的)。
// 读数: ai/022 第二百二十三刀。
static inline void vb6_SA_Destroy(void* arr) {
    if (!arr) return;
    if (vb6_SaIs1D(arr)) vb6_SafeArrayDestroy1D((vb6_SafeArray1D*)arr);
    else vb6_SafeArrayDestroyND((vb6_SafeArrayND*)arr);
}

// vb6_DebugAssert — Debug.Assert (no-op in compiled mode)
static inline void vb6_DebugAssert(int32_t cond) {
    (void)cond;  /* no-op: Debug.Assert only active in IDE */
}

// vb6_DebugPrint — Debug.Print (no-op in compiled mode, use DebugPrintStr for actual output)
static inline void vb6_DebugPrint(BSTR s) {
    vb6_DebugPrintStr(s);
}

// Fix <vbeclipse>: LoadRes* 实装 (2026-10-06) — 资源从用户 .res (ResFile32, 链接时
// 已随 exe 进资源段) 用 FindResource/LoadResource 直读。三个函数的实参在 VB6 里
// 本来就是 Variant (LoadResString(101) 数字 id / LoadResData("BIN1","CUSTOM")
// 字符串名都能出现), 所以形参统一 vb6_VARIANT, cgen 侧用 vb6_VariantFromValue
// 包装 (见 cgen_expr_call_builtin_fixup.inc)。找不到资源按 VB6 抛错误 326。
vb6_VARIANT vb6_LoadResData(vb6_VARIANT resourceId, vb6_VARIANT resourceType);
vb6_VARIANT vb6_LoadResPicture(vb6_VARIANT resourceId, vb6_VARIANT resourceType);
vb6_VARIANT vb6_LoadResString(vb6_VARIANT resourceId);

#ifdef __cplusplus
}
#endif
