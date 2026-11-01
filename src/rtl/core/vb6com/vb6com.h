#pragma once
// vb6com.h - VB6 COM互操作运行时 (P6)
// 独立于vb6rtl.h, 避免VARIANT定义冲突
// 此文件使用Windows原生VARIANT/IDispatch等类型

#include <windows.h>
#include <oleauto.h>
#include <wchar.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

// CreateObject(progId) - 通过ProgID创建COM对象，返回IDispatch*
void* vb6_CreateObject(const wchar_t* progId);

// typelib appobject 全局单例 (VB6 全局对象): 首次按 ProgID CoCreate 并缓存,
// 之后每次调用复用同一实例 (VB6 GlobalMultiUse/appobject 的进程级单例语义 ——
// 同一实例上的状态如 GlobalSettings.Icons 累积的位图在多次访问间保持)。
// 失败返回 NULL (错误已由 vb6_CreateObject raise)。
void* vb6_ComGlobalInstance(const wchar_t* progId);

// GetObject(pathName, progId) - 获取运行中的COM对象
void* vb6_GetObject(const wchar_t* pathName, const wchar_t* progId);

// ============================================================
// Fix 160: 免注册 COM (vbp ComLib= 声明的组件 DLL)
// ============================================================
// cgen 在入口点烘焙的组件表 — 每个元素对应一个 ComLib= DLL 里的一个 coclass.
// vb6_CreateObject 按 ProgID 命中后改走 LoadLibrary+DllGetClassObject, 完全绕开
// 注册表 (与 Object= 控件的免注册路径同款机制). 未命中 → 原有注册表路径.
typedef struct Vb6ComLib {
    const wchar_t* progId;       /* "MyLib.MyClass" */
    const wchar_t* clsidStr;     /* "{XXXXXXXX-...}" — 来自 TYPEATTR->guid, 免注册可得 */
    const wchar_t* coclassName;  /* "MyClass" — progid 缺失时的末段名兜底匹配 */
    const wchar_t* fileName;     /* "bin\mylib.dll" — 相对 exe, 不依赖 CWD */
    ITypeInfo*     dispIface;    /* 免注册默认调度接口 ITypeInfo (vb6_ComLibRegister 加载, 无注册表);
                                     非 NULL 时 vb6_getDispid 经它解析 DISPID, 绕过对象自身
                                     的 GetIDsOfNames (Chilkat 等 typelib 驱动 IDispatch 无需注册 TLB) */
} Vb6ComLib;

/* 免注册 COM 对象 → 默认调度接口 ITypeInfo 查找 (vb6_getDispid 用来解析 DISPID).
 * 命中返回 AddRef 过的 ITypeInfo*, 调用方应 Release; 未命中返回 NULL. */
ITypeInfo* vb6_ComLibLookupTypeInfo(void* disp);

// 注册组件表 (cgen 生成, 入口点调用一次). count 超过 256 截断并打 stderr 警告.
void vb6_ComLibRegister(const Vb6ComLib* libs, int count);

// IsNothing(obj) - 检查对象引用是否为Nothing
int32_t vb6_IsNothing(void* obj);

// ReleaseObject(&ptr) - 释放COM对象引用并置NULL
void vb6_ReleaseObject(void** objPtr);

// 接收者是否真是 COM 对象 (首槽 vtable 的前 7 槽指向可执行内存)。
// 伪装的接收者 (UDT/类实例的地址) 一旦解引用 lpVtbl 就是 AV, 而不是 VB6 的
// "对象不支持此属性或方法" —— 后期绑定与 Release 路径都先用它把这类挡掉。
int32_t vb6_ComIsDispatchable(const void* disp);

// 账 #305 第二百六十二刀 (丙): 晚绑定出口漏掉的「非 IDispatch 接收者」计数。
// 七个出口 (ComCall / ComCallByDispid / ComGetProp / ComGetPropArg / ComSetProp /
// ComSetPropArg / ComSetRef) 现在都会把这种接收者**报一行** (stderr 标记
// C3_LATEBIND_MISS) 并计数; 行为不变 (仍按原路返回空)。
// 判据「非 IDispatch 进晚绑定出口 = 0」按 stderr 里那条标记的条数读。
int32_t vb6_ComLateBindMissCount(void);

// Fix <vbeclipse>: vb6_ReleaseObject 的反向操作 (AddRef), 判据相同 (vb6_ComIsDispatchable)。
// 供 vb6rtl_variant.h 的 vb6_VariantObject / vb6rtl.c 的 vb6_VariantCopy 使用 ——
// Variant 既然在 Clear 时 Release, 构造/拷贝时就必须自己持有一份引用。
void vb6_ComAddRefDispatch(void* p);

// COM后期绑定 (P6.2)
// 返回VARIANT* (Windows VARIANT), 调用方需vb6_ComVarClear释放
void* vb6_ComCall(void* disp, const wchar_t* methodName,
                  void* args, int32_t argc);
void* vb6_ComCallByDispid(void* disp, int32_t dispid,
                         void* args, int32_t argc);
void* vb6_ComGetProp(void* disp, const wchar_t* propName);
void vb6_ComSetProp(void* disp, const wchar_t* propName, void* value);
void vb6_ComSetPropArg(void* disp, const wchar_t* propName, void** args, int32_t argc, void* value);
void vb6_ComSetRef(void* disp, const wchar_t* propName, void* objRef);

// COM VARIANT封装/解封 (P6.2 cgen使用)
// vb6rtl.h通过void*声明这些函数，避免VARIANT类型冲突

// 将BSTR封装为VARIANT (返回堆分配的VARIANT*, 需vb6_ComVarClear释放)
void* vb6_ComPackBSTR(const wchar_t* bstr);

// 将int32_t封装为VARIANT
void* vb6_ComPackInt(int32_t val);


// 将VB6 Boolean (int32_t: -1=True, 0=False) 封装为VARIANT VT_BOOL
void* vb6_ComPackBool(int32_t val);
// 将double封装为VARIANT
void* vb6_ComPackDouble(double val);

// 将void*(IDispatch*)封装为VARIANT (用于对象参数)
void* vb6_ComPackObject(void* obj);

// Fix 160-com-byref: 将对象变量的地址封装为 VT_DISPATCH|VT_BYREF VARIANT,
// 用于 COM 方法的 ByRef 对象出参 (如 Chilkat ObjectOf2/ArrayOf2 的 obj/array 形参)。
// ppObj 指向调用方的 void* 变量; Invoke 写入新指针后, 调用方变量即持有新对象。
void* vb6_ComPackObjectRef(void** ppObj);

// 账 #305 §B224 第二半 (第二百三十九刀): 标量实参按引用打包 (VT_BYREF|型)。
// 真 VB6 晚绑定对带型变量递的就是这一形 (实测 x As Long 被 `Bump(ByRef v As Long)` 写回成 11);
// 按 C 存储宽度分五档 —— Integer/Boolean 都是 2 字节, 不许按 I4 打。
void* vb6_ComPackIntRef(int32_t* p);
void* vb6_ComPackIntegerRef(int16_t* p);
void* vb6_ComPackBoolRef(int16_t* p);
void* vb6_ComPackSingleRef(float* p);
void* vb6_ComPackDoubleRef(double* p);

// Fix 104: 将"省略的实参"封装为 VARIANT (VT_ERROR + DISP_E_PARAMNOTFOUND).
// VB6 遇到 obj.Method a, , c 这类省略实参时, 对该形参位置传的正是这个值 ——
// 接收方 (IDispatch 实现) 据此把"未提供"与"显式传 0/空串"区分开.
// 直接丢弃省略实参会让后续实参前移 (a, , c 被当成 a, c), 即参数错位.
void* vb6_ComPackMissing(void);

// 从VARIANT*解封BSTR (返回BSTR, 需vb6_BSTR_Free释放)
wchar_t* vb6_ComUnpackBSTR(void* variant);

// 从VARIANT*解封int32_t
int32_t vb6_ComUnpackInt(void* variant);

// 从VARIANT*解封double
double vb6_ComUnpackDouble(void* variant);

// 从VARIANT*解封对象(void*/IDispatch*)
void* vb6_ComUnpackObject(void* variant);

// 释放ComCall/ComGetProp返回的VARIANT* (值类型安全: 释清BSTR/数字等)
void vb6_ComVarClear(void* variant);

// 仅释放VARIANT结构体, 不清除内容 (对象引用已转移给调用方)
void vb6_ComVarFree(void* variant);

// 一体化COM辅助函数 (内部处理临时VARIANT清理, cgen直接使用)

// COM方法调用→对象 (内部UnpackObject+VarFree)
void* vb6_ComCallObject(void* disp, const wchar_t* methodName,
                        void* args, int32_t argc);
// COM**默认成员**调用→对象: `coll(idx)` 的 VB6 语义是 DISPID_VALUE(0), 不是按名
// "Item"。名字表里没有 "Item" 时 (Codejock 的 ReportColumns/Actions 一族默认成员
// 叫 Column/Action) 按名找一律 DISP_E_MEMBERNOTFOUND ⇒ 交回 NULL。
void* vb6_ComCallObjectDefault(void* disp, void* args, int32_t argc);
// COM方法调用→BSTR (内部UnpackBSTR+VarClear)
wchar_t* vb6_ComCallBSTR(void* disp, const wchar_t* methodName,
                         void* args, int32_t argc);
// COM方法调用→int32_t (内部UnpackInt+VarClear)
int32_t vb6_ComCallInt(void* disp, const wchar_t* methodName,
                       void* args, int32_t argc);
// COM方法调用→double (内部UnpackDouble+VarClear)
double vb6_ComCallDouble(void* disp, const wchar_t* methodName,
                         void* args, int32_t argc);

// COM属性Get→BSTR (内部UnpackBSTR+VarClear)
wchar_t* vb6_ComGetStringProp(void* disp, const wchar_t* propName);
// COM属性Get→int32_t (内部UnpackInt+VarClear)
int32_t vb6_ComGetIntProp(void* disp, const wchar_t* propName);
// COM属性Get→intptr_t (LongPtr: 句柄/指针; 兼容 32/64 位整数变体)
intptr_t vb6_ComGetLongPtrProp(void* disp, const wchar_t* propName);
// COM属性Get→double (内部UnpackDouble+VarClear)
double vb6_ComGetDoubleProp(void* disp, const wchar_t* propName);
// COM属性Get→对象 (内部UnpackObject+VarFree)
void* vb6_ComGetObjectProp(void* disp, const wchar_t* propName);
// `Set <接口类型变量> = <对象>` 的类型检查 (VB6: 不符 = 错误 13):
// 是本宿主托管的 OCX 控件时按控件类名比对该接口名, 不符返回 NULL;
// 其余对象原样返回 (普通 COM 对象 / 工程类 / 宿主包装器都不受影响)。
void* vb6_ComObjectTypeCast(void* obj, const wchar_t* typeName);

// 参数化属性Get (带参数, 如Dictionary.Item(key))
// 内部使用DISPATCH_METHOD|DISPATCH_PROPERTYGET组合标志
void* vb6_ComGetPropArg(void* disp, const wchar_t* propName,
                        void* args, int32_t argc);
wchar_t* vb6_ComGetPropertyString(void* disp, const wchar_t* propName,
                                   void* args, int32_t argc);
int32_t vb6_ComGetPropertyInt(void* disp, const wchar_t* propName,
                               void* args, int32_t argc);
double vb6_ComGetPropertyDouble(void* disp, const wchar_t* propName,
                                 void* args, int32_t argc);
void* vb6_ComGetPropertyObject(void* disp, const wchar_t* propName,
                                void* args, int32_t argc);
void* vb6_ComGetPropertyVariant(void* disp, const wchar_t* propName,
                                 void* args, int32_t argc);

// ============================================================
// P6.3: COM前期绑定运行时支持 (vtable直接调用)
// ============================================================

// QueryInterface获取指定接口指针 (用于前期绑定初始化)
// iidStr: IID字符串 (如 "{2A0A3E20-...}")
// 返回: 请求的接口指针, 失败返回NULL
void* vb6_ComQI(void* obj, const char* iidStr);

// 创建前期绑定COM对象: CreateObject + QI
// progId: ProgID, iidStr: 目标接口IID
// 返回: 接口指针 (前期绑定类型), 失败返回NULL
void* vb6_ComCreateTyped(const wchar_t* progId, const char* iidStr);

// 释放前期绑定COM对象 (与ReleaseObject类似, 但针对接口指针)
void vb6_ComReleaseTyped(void** objPtr);

// vtable直接调用辅助函数 (P6.3 cgen使用)
// obj: 接口指针, vtIndex: vtable偏移量
// 参数通过va_list传递, 返回值从VARIANT解封

// vtable调用→void (无返回值/丢弃)
void vb6_ComVtableCallVoid(void* obj, int32_t vtIndex, ...);

// vtable调用→BSTR
wchar_t* vb6_ComVtableGetBSTR(void* obj, int32_t vtIndex, ...);

// vtable调用→int32_t
int32_t vb6_ComVtableGetInt(void* obj, int32_t vtIndex, ...);

// vtable调用→double
double vb6_ComVtableGetDouble(void* obj, int32_t vtIndex, ...);

// vtable调用→对象(void*)
void* vb6_ComVtableGetObject(void* obj, int32_t vtIndex, ...);

// vtable调用→通用void* (VARIANT结果)
void* vb6_ComVtableGetVoid(void* obj, int32_t vtIndex, ...);

// COM初始化/退出 (由vb6_Init/vb6_Exit调用)
// P13.21: IConnectionPointContainer / Advise support (WithEvents)
// Connect a COM object's event source to a sink (IDispatch-based callback)
int vb6_ComAdvise(void* obj, const char* riidStr, void* sink, int* adviseCookie);
int vb6_ComUnadvise(void* obj, const char* riidStr, int adviseCookie);

// P13.22: Create a generic IDispatch event sink
// Maps DISPID (event IDs) to callback functions
// sourceIid: optional source interface IID (for dispinterface event sinks)
// handler: consumer object passed to each callback as first argument
void* vb6_CreateEventSink(const int* dispids, void** callbacks, int count, const IID* sourceIid, void* handler);
void vb6_FreeEventSink(void* sink);
void vb6_ComInit(void);
void vb6_ComExit(void);


// P22-11: For Each COM collection (IEnumVARIANT)
void* vb6_ForEach_Init(void* disp);       // Call _NewEnum, return IEnumVARIANT*
// 账 #257 第十五层: 集合是**工程类实例**时 (不是 IDispatch) 用它的 NewEnum 接口直接建枚举
void* vb6_ForEach_InitFromEnumIface(void* pEnumRaw);

// For Each Form In Forms: RTL 窗体注册表 (vb6_Forms_Register) 的合成枚举器。
// 元素 = 窗体 HWND 的宿主包装 (真 IDispatch, .Controls 走宿主 Controls 应答)。
void* vb6_ForEach_InitForms(void);
int32_t vb6_ForEach_Next(void* enumPtr, VARIANT* outVar);  // Next element, returns 1=ok 0=done
void vb6_ForEach_Release(void* enumPtr);  // Release IEnumVARIANT

#ifdef __cplusplus
}
#endif
