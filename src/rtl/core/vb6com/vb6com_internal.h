// vb6com_internal.h - vb6com 模块内部共享声明 (P6 COM 互操作)
// 仅供 RTL 内部各 .c 使用; 生成代码只 include vb6com.h
//
// 拆分说明: vb6com.c 按 COM 调用层次拆为多个编译单元, 各单元由 MSVC 独立
// 编译成 .obj 再链接, 文件级 static 不跨文件可见 —— 因此把跨族共享的符号
// 集中声明在此, 定义仍留在 vb6com.c。

#ifndef VB6C3_VB6COM_INTERNAL_H
#define VB6C3_VB6COM_INTERNAL_H

#include "vb6com.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

// P24-08: COM错误传播 — 将HRESULT/EXCEPINFO转换为VB6运行时错误
// VB6行为: COM方法失败时自动触发Err.Raise, 可被On Error捕获
// (定义在 vb6rtl.c; vb6com 系列不 include vb6rtl.h, 以避免 VARIANT 冲突)
extern void vb6_RaiseError(int32_t errNum, void* description);
extern int32_t vb6_err_resume_next;

// 账 #173: 同一处口径下漏登记的第二个跨族符号。
// vb6_VariantToDouble 返回 double —— x86 的 cdecl 把返回值放在 **x87 栈 ST0**,
// 调用方必须 `fstp` 弹出; 没有原型时 MSVC 只给 C4013 "假设外部返回 int",
// 于是调用方按 EAX 取值、**ST0 永不弹出** ⇒ 每调一次漏一层 x87, 八层之后栈满,
// 下一次浮点运算得到无效操作 QNaN (0x7FF8000000000000)。症状因此离真因极远:
// 实测在 Charts 2020 (x86) 里表现为不相干的 `CByte((Abs(Opacity)/100)*255)`
// 收到 NaN ⇒ run-time error 6 "Overflow", 且**任何多插入的调用都会让它换位置**
// (Heisenbug)。x64 的 double 返回值走 XMM0, 没有 x87 栈 ⇒ 同一份代码在 x64
// 永远不现形, 而 CI 的门只跑默认架构。
// 形参写 Windows VARIANT 是安全的: vb6_VARIANT 与它**逐平台同布局** (rev11 口径,
// 见 vb6rtl_variant.h 里那段注释), 按值传结构体的 ABI 只取决于尺寸与偏移。
extern double vb6_VariantToDouble(VARIANT v);

// --- 跨族共享的内部辅助 (定义在 vb6com.c) ---
// COM错误->VB6错误转换
// hr: Invoke返回的HRESULT / excep: EXCEPINFO结构 / context: 调用上下文
void vb6_ComCheckError(HRESULT hr, EXCEPINFO* excep, const wchar_t* context);

// VB6标准错误号 -> 中文描述 (查不到返回NULL)
// 供手写COM对象(如 vb6com_collection.c)填 EXCEPINFO.bstrDescription
const wchar_t* vb6_StdErrorDesc(int32_t errNum);

// IUnknown 的 C 风格 vtable 辅助 (避免 C++ IUnknown 方法调用问题)
HRESULT vb6_UnknownQI(void* obj, REFIID riid, void** ppv);
ULONG vb6_UnknownRelease(void* obj);

// ============================================================
// Fix 103: 内建 Collection 的跨编译单元共享定义
// 实现分在两个编译单元 —— vb6com_collection.c (容器 + IDispatch) 与
// vb6com_collection_enum.c (_NewEnum 的 IEnumVARIANT), 结构体与枚举器
// vtable 必须两边可见。
// ============================================================
typedef struct vb6_CollEntry {
    BSTR    key;      // NULL 表示该元素无键
    VARIANT value;
} vb6_CollEntry;

// Collection 实例。lpVtbl 必须是首成员 (COM ABI)。
typedef struct vb6_BuiltinCollection {
    IDispatchVtbl* lpVtbl;
    LONG           refCount;
    int32_t        count;
    int32_t        capacity;
    vb6_CollEntry* entries;
} vb6_BuiltinCollection;

// _NewEnum 返回的枚举器。持有集合引用, 按 0-based 下标向前推进。
typedef struct vb6_CollEnum {
    IEnumVARIANTVtbl*      lpVtbl;   // 必须是首成员
    LONG                   refCount;
    vb6_BuiltinCollection* coll;
    int32_t                index;
} vb6_CollEnum;

// 枚举器 vtable (定义在 vb6com_collection_enum.c; coll_doNewEnum 用它装配实例)
extern IEnumVARIANTVtbl g_collEnumVtbl;

// ============================================================
// Fix 160-com: 免注册分派辅助声明 (定义在 vb6com_invoke.c)
// ============================================================
// 免注册 GetIDsOfNames (For Each 等直接路径)
HRESULT vb6_ComGetIdsOfNames(IDispatch* pDisp, const wchar_t* name, DISPID* pDispid);
// 免注册 Invoke 总入口 (ITypeInfo 驱动, 无 TI 映射时落回 IDispatch::Invoke)
HRESULT vb6_ComInvoke(IDispatch* pDisp, DISPID dispid, WORD wFlags,
                      DISPPARAMS* dp, VARIANT* result, EXCEPINFO* excep, UINT* argErr);

#endif // VB6C3_VB6COM_INTERNAL_H
