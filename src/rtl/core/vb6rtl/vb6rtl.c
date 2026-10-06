// vb6rtl.c - VB6运行时库最小实现
// 仅支持 hello.bas 等简单程序运行
// 2026-09-17 按家族拆分为 13 个编译单元（纯搬移，逐行未改）。
// 本文件保留：头部 include + 运行时初始化/退出 + 内存/类支持 + Variant 转换与比较 + 错误处理。
// 拆出文件与原行区间：
//   vb6rtl_string.c        40~64, 65~283, 1606~1728, 1729~1789, 4667~4698
//   vb6rtl_format.c        284~893
//   vb6rtl_conv.c          894~902, 903~918, 919~953, 1272~1327, 1328~1439, 2680~2707, 3149~3182
//   vb6rtl_misc.c          954~1271, 2512~2567
//   vb6rtl_system.c        1790~1950, 1951~1993, 1994~2065, 2066~2133, 2134~2208, 2209~2236
//   vb6rtl_compat.c        2237~2511, 2568~2679
//   vb6rtl_date.c          2708~2827, 2828~3044, 3045~3148
//   vb6rtl_array.c         3183~3373, 3374~3617
//   vb6rtl_file.c          3618~3918, 4115~4181, 4595~4666
//   vb6rtl_paramarray.c    4182~4430
//   vb6rtl_financial.c     4431~4594
//   vb6rtl_com.c           4699~4793, 4955~5327
//   vb6rtl_registry.c      4794~4954


#include "vb6rtl.h"
#include "vb6rtl_crash.h"   /* 账 #181: vb6_CrashTraceClaim */
#include <intrin.h>   // _ReturnAddress (未处理错误定位)
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
// COM互操作前向声明 (实现在vb6com.c中，避免vb6_VARIANT类型冲突)
// ============================================================
extern void* vb6_CreateObject(const wchar_t* progId);
extern void* vb6_GetObject(const wchar_t* pathName, const wchar_t* progId);
extern int32_t vb6_IsNothing(void* obj);
extern void vb6_ReleaseObject(void** objPtr);
extern void* vb6_ComCall(void* disp, const wchar_t* methodName, void* args, int32_t argc);
extern void* vb6_ComCallByDispid(void* disp, int32_t dispid, void* args, int32_t argc);
extern void* vb6_ComGetProp(void* disp, const wchar_t* propName);
extern void vb6_ComSetProp(void* disp, const wchar_t* propName, void* value);
extern void vb6_ComSetRef(void* disp, const wchar_t* propName, void* objRef);
extern void vb6_ComInit(void);
extern void vb6_ComExit(void);

// ============================================================
// 整除 / 幂运算
// ============================================================

int32_t vb6_IntDiv(int32_t a, int32_t b) {
    // VB6 语义: \ 除 0 → 运行期错误 11 ("Division by zero"), 与 '/' 同源
    // (此前 TODO "raise error" 一直没接, 变量除数静默返回 0)。
    if (b == 0) { vb6_ErrRaiseNumber(11); return 0; }
    // VB6 \ 运算符: 截断到整数 (C的整数除法对正负数的行为与VB6一致)
    return a / b;
}

double vb6_Pow(double base, double exp) {
    return pow(base, exp);
}

// ============================================================
// 运行时初始化/退出
// ============================================================

// Fix 172: 崩溃回溯钩子 —— 生成代码的 AV 只有 WER 里一个"故障偏移", 定位不到调用者。
// 装一个 VEH, 访问违例时把异常地址与栈帧**按 RVA** 打到 stderr (与 exe 首选基址
// 0x140000000 相加即可用 PDB 解析到生成的 .c 行号, 见 .temp/sym3.ps1)。
// 仅在 C3_COM_TRACE / C3_CRASH_TRACE 环境变量存在时安装, 只打印不改流程
// (继续 EXCEPTION_CONTINUE_SEARCH, WER/退出码不变)。
#ifdef _WIN32
static LONG CALLBACK vb6_CrashTraceVEH(PEXCEPTION_POINTERS ep) {
    DWORD code = ep->ExceptionRecord->ExceptionCode;
    /* 账 #181: 记轨迹本身会再触发异常 (栈溢出时 fprintf 拿 CRT 锁 /
     * CaptureStackBackTrace 再读同一批页) ⇒ 每进程只记第一次。 */
    if (!vb6_CrashTraceClaim(VB6_CRASH_CLAIM_STDERR)) return EXCEPTION_CONTINUE_SEARCH;
    if (code == EXCEPTION_ACCESS_VIOLATION || code == EXCEPTION_IN_PAGE_ERROR
        || code == EXCEPTION_ILLEGAL_INSTRUCTION || code == EXCEPTION_INT_DIVIDE_BY_ZERO
        || code == EXCEPTION_STACK_OVERFLOW) {
        HMODULE hSelf = GetModuleHandleA(NULL);
        void* frames[24];
        USHORT n = (USHORT)CaptureStackBackTrace(0, 24, frames, NULL);
        fprintf(stderr, "[C3_CRASH] code=0x%lx at rva=0x%lx base=%p frames=%u\n",
                (unsigned long)code,
                (unsigned long)((char*)ep->ExceptionRecord->ExceptionAddress - (char*)hSelf),
                (void*)hSelf, (unsigned)n);
        if (code == EXCEPTION_ACCESS_VIOLATION) {
            /* Fix 187 诊断: ExceptionInformation[0] 的取值是 0=读 / 1=写 / **8=执行(DEP)** ——
             * 8 表示 CPU 跳到了一个不可执行的地址(常见为 0, 即 call NULL), 与"读写越界"
             * 是完全不同的故障形态, 打成 write=8 会把人引向"缺空指针检查的出参"。 */
            fprintf(stderr, "[C3_CRASH]   av %s target=0x%llx\n",
                    ep->ExceptionRecord->ExceptionInformation[0] == 8 ? "EXECUTE(DEP)" :
                    ep->ExceptionRecord->ExceptionInformation[0] ? "write" : "read",
                    (unsigned long long)ep->ExceptionRecord->ExceptionInformation[1]);
        }
        /* 账 #182: 帧不落在本 exe 镜像里时**照实标出来**。VEH 里 CaptureStackBackTrace
         * 拿到的是派发链自己（ntdll 若干帧 + 处理器的返回地址），旧写法把它们一律减掉
         * 本模块基址打成 "rva=0x4376..."，看着像本模块的符号、实际是隔壁 DLL 的地址。 */
        PIMAGE_DOS_HEADER exeDos = (PIMAGE_DOS_HEADER)hSelf;
        PIMAGE_NT_HEADERS exeNt = (PIMAGE_NT_HEADERS)((char*)hSelf + exeDos->e_lfanew);
        char* exeLo = (char*)hSelf;
        char* exeHi = exeLo + exeNt->OptionalHeader.SizeOfImage;
        for (USHORT i = 0; i < n; i++) {
            if ((char*)frames[i] >= exeLo && (char*)frames[i] < exeHi) {
                fprintf(stderr, "[C3_CRASH] #%u rva=0x%lx\n", (unsigned)i,
                        (unsigned long)((char*)frames[i] - exeLo));
            } else {
                fprintf(stderr, "[C3_CRASH] #%u %p (outside exe)\n", (unsigned)i, frames[i]);
            }
        }
#if defined(_M_IX86) || defined(_M_X64)
        /* x86/x64 上 CaptureStackBackTrace 都只返回 VEH/异常派发链, 应用侧调用者全丢
         * (账 #182 之前这段只编 x86 ⇒ x64 的崩溃现场一条应用帧都没有)。
         * 追加一次栈指针线性扫描, 打印落在本模块镜像内的候选返回地址(按栈深度标 st+N)。 */
        {
#ifdef _M_IX86
            DWORD* sp = (DWORD*)ep->ContextRecord->Esp;
#else
            ULONG_PTR* sp = (ULONG_PTR*)ep->ContextRecord->Rsp;
#endif
            int printed = 0;
            for (int k = 0; k < 512 && printed < 24; k++) {
#ifdef _M_IX86
                DWORD v;
#else
                ULONG_PTR v;
#endif
                __try { v = sp[k]; }
                __except (EXCEPTION_EXECUTE_HANDLER) { break; }
                if ((char*)v >= exeLo && (char*)v < exeHi) {
                    fprintf(stderr, "[C3_CRASH]   st+%d rva=0x%lx\n", k, (unsigned long)((char*)v - exeLo));
                    printed++;
                }
            }
        }
#endif
        fflush(stderr);
    }
    return EXCEPTION_CONTINUE_SEARCH;
}
#endif

void vb6_Init(void) {
    // 初始化随机种子
    srand((unsigned int)time(NULL));
#ifdef _WIN32
    if (GetEnvironmentVariableA("C3_COM_TRACE", NULL, 0) > 0
        || GetEnvironmentVariableA("C3_CRASH_TRACE", NULL, 0) > 0) {
        AddVectoredExceptionHandler(1, vb6_CrashTraceVEH);
    }
#endif
    // P20-44: OLE 拖放 RTL 自测 —— 环境变量 C3_OLEDDB_TEST=1 时, 把
    // IDataObject/IDropTarget 的 DragEnter/DragOver/Drop 走一遍并写结果文件。
    // 无头环境没法真拖 (DoDragDrop 是模态循环), 直接调 IDropTarget 的方法才测得动。
    {
        // 用 W 版取环境变量: 输出路径可能含非 ASCII (fwprintf/_wfopen 全宽链路)
        wchar_t oleOut43[MAX_PATH] = { 0 };
        wchar_t oleFlag43[8] = { 0 };
        if (GetEnvironmentVariableW(L"C3_OLEDDB_TEST", oleFlag43, 8) > 0) {
            if (!GetEnvironmentVariableW(L"C3_OLEDDB_TEST_OUT", oleOut43, MAX_PATH))
                lstrcpyW(oleOut43, L"oledd_test.txt");
            extern int32_t vb6_oleDD_SelfTest(const wchar_t* outPath);
            vb6_oleDD_SelfTest(oleOut43);
        }
    }
    // comctl32 通用控件注册 (ProgressBar/StatusBar/Toolbar/ListView/TreeView)。
    // 必须在任何通用控件 CreateWindowExW 之前 —— 否则类未注册,
    // CreateWindowExW 静默返回 NULL (GetLastError()==1400), 控件凭空消失。
    vb6_ComCtl_Init();
    // 初始化COM库 (实现在vb6com.c中)
    vb6_ComInit();
}

void vb6_Exit(void) {
    // 清理COM库
    vb6_ComExit();
}

void vb6_End(void) {
    vb6_Exit();
    exit(0);
}

void vb6_Beep(void) {
#ifdef _WIN32
    // Beep() requires windows.h, use MessageBeep as fallback
    MessageBeep(0);
#else
    putchar('\a');
    fflush(stdout);
#endif
}
// P18-C: Option Compare
int g_vb6_optionCompareText = 0;  // 0=Binary(default), 1=Text
int vb6_StrCmp(const wchar_t* a, const wchar_t* b) {
    // Fix 173: VB6 里未赋值的 String (UDT 字段 / 未初始化变量) 是 **NULL BSTR**
    // (= vbNullString), 与 L"" 比较相等; 直接把 NULL 喂给 wcscmp → 读 0x0
    // → 0xC0000005。踩到点: VBFlexGrid.ctl:20425 `If Not .Format = vbNullString`
    // (GetTextDisplay 里 Format 从未赋值的列 → .Format==NULL)。
    if (!a) a = L"";
    if (!b) b = L"";
    if (g_vb6_optionCompareText) return vb6_TextCmp(a, b);
    return wcscmp(a, b);
}

// <vbeclipse>: 文本模式的**显式**比较入口。VB6 的 Option Compare 是**按模块**的编译期
// 属性, 而 g_vb6_optionCompareText 是进程唯一的全局 —— 一个 Text 模块会把同进程里
// Binary 模块的 `=`/`<>`/Select Case 语义一起带跑。因此 codegen 按"当前模块的
// Option Compare"直接选函数名 (Text 模块发 vb6_StrCmpT), 这个全局保留给
// 旧的运行期路径, 不再是判定入口。
int vb6_StrCmpT(const wchar_t* a, const wchar_t* b) {
    return vb6_TextCmp(a, b);
}

// ============================================================
// 类支持: 实例分配/释放
// ============================================================

void* vb6_Alloc(size_t size) {
    void* p = calloc(1, size);  // calloc 自动清零 = VB6默认值初始化
    return p;
}

void vb6_Free(void* ptr) {
    free(ptr);
}

int32_t vb6_VariantToLong(vb6_VARIANT v) {
    switch (v.vt) {
        case vb6_vtBoolean: return v.boolVal ? -1 : 0;
        case vb6_vtByte:    return (int32_t)v.bVal;
        case vb6_vtInteger: return (int32_t)v.iVal;
        case vb6_vtLong:    return v.lVal;
        case VT_I8:         return (int32_t)v.llVal;  /* Task #44: VT_I8 进 Long 按 C 截断语义 */
        case vb6_vtSingle:  return (int32_t)round(v.fltVal);
        case vb6_vtDate:    return (int32_t)round(v.dblVal);  /* <vbeclipse>: VT_DATE 的数值面就是序列号 */
        case vb6_vtDouble:  return (int32_t)round(v.dblVal);
        case vb6_vtCurrency:return (int32_t)(v.cyVal / 10000);
        case vb6_vtBSTR:    return (int32_t)vb6_Val(v.bstrVal);
        default:            return 0;
    }
}

// Variant → LongPtr (指针/句柄语义). 数值提取语义与 vb6_VariantToLong 一致,
// 但返回 intptr_t, 避免 x64 下把 64 位句柄/指针截断为 32 位.
intptr_t vb6_VariantToLongPtr(vb6_VARIANT v) {
    switch (v.vt) {
        case vb6_vtBoolean: return v.boolVal ? -1 : 0;
        case vb6_vtByte:    return (intptr_t)v.bVal;
        case vb6_vtInteger: return (intptr_t)v.iVal;
        case vb6_vtLong:    return (intptr_t)v.lVal;
        case vb6_vtSingle:  return (intptr_t)round(v.fltVal);
        case vb6_vtDate:    return (intptr_t)round(v.dblVal);  /* <vbeclipse>: 同上 */
        case vb6_vtDouble:  return (intptr_t)round(v.dblVal);
        case vb6_vtCurrency:return (intptr_t)(v.cyVal / 10000);
        case vb6_vtBSTR:    return (intptr_t)vb6_Val(v.bstrVal);
        case VT_I8:         return (intptr_t)v.llVal;
        case VT_UI8: {      /* 位模式按无符号解释 */
            uint64_t uv;
            memcpy(&uv, &v.llVal, sizeof(uv));
            return (intptr_t)uv;
        }
        default:            return 0;
    }
}

// Fix 093a: Variant → Boolean (VB6 CBool 语义, True = -1). BSTR 先按
// "True"/"False" 文本判断, 其余按数值 != 0 (VB6 CBool 对数字非零即 True).
int16_t vb6_VariantToBool(vb6_VARIANT v) {
    switch (v.vt) {
        case vb6_vtBoolean:  return v.boolVal ? -1 : 0;
        case vb6_vtByte:     return v.bVal ? -1 : 0;
        case vb6_vtInteger:  return v.iVal ? -1 : 0;
        case vb6_vtLong:     return v.lVal ? -1 : 0;
        case VT_I8:          return v.llVal ? -1 : 0;  /* Task #44: LongLong/LongPtr 64 位 */
        case vb6_vtSingle:   return v.fltVal != 0.0f ? -1 : 0;
        case vb6_vtDate:     return v.dblVal != 0.0 ? -1 : 0;  /* <vbeclipse>: 同上 */
        case vb6_vtDouble:   return v.dblVal != 0.0 ? -1 : 0;
        case vb6_vtCurrency: return v.cyVal ? -1 : 0;
        case vb6_vtDispatch: return v.pdispVal ? -1 : 0;
        case vb6_vtBSTR:
            if (!v.bstrVal) return 0;
            if (_wcsicmp(v.bstrVal, L"true") == 0) return -1;
            if (_wcsicmp(v.bstrVal, L"false") == 0) return 0;
            return vb6_Val(v.bstrVal) != 0.0 ? -1 : 0;
        default:             return 0;  // Empty / Null / Error / 数组
    }
}

double vb6_VariantToDouble(vb6_VARIANT v) {
    switch (v.vt) {
        case vb6_vtBoolean: return v.boolVal ? -1.0 : 0.0;
        case vb6_vtByte:    return (double)v.bVal;
        case vb6_vtInteger: return (double)v.iVal;
        case vb6_vtLong:    return (double)v.lVal;
        case VT_I8:         return (double)v.llVal;  /* Task #44: LongLong/LongPtr 64 位 */
        case vb6_vtSingle:  return (double)v.fltVal;
        case vb6_vtDate:    return v.dblVal;  /* <vbeclipse>: 同上 */
        case vb6_vtDouble:  return v.dblVal;
        case vb6_vtCurrency:return (double)v.cyVal / 10000.0;
        case vb6_vtBSTR:    return vb6_Val(v.bstrVal);
        default:            return 0.0;
    }
}

BSTR vb6_VariantToString(vb6_VARIANT v) {
    return vb6_CStr(v);
}

// Fix 029: 从 Variant 中提取 SafeArray1D* (当 Variant 持有数组时).
// 用于调用点反向强制: callee 期望 vb6_SafeArray1D* 但实参是 vb6_VARIANT.
struct vb6_SafeArray1D* vb6_VariantToSafeArray1D(vb6_VARIANT v) {
    if ((v.vt & vb6_vtArray) && v.parray) {
        return v.parray;
    }
    /* Variant 不持有数组时返回 NULL (与 VB6 行为一致; 调用方需 NULL 检查) */
    return NULL;
}

// Fix 029: vb6_VariantToObject 的右值兼容版本.
// vb6_VariantToObject 接受 vb6_VARIANT* (要求实参左值), 而调用点包装的实参
// 经常是函数返回值 (vb6_VariantArrayGet(...) 等) 无法取址. 这里提供按值版本.
void* vb6_VariantToObjectVal(vb6_VARIANT v) {
    if (getenv("C3_IV_TRACE")) {
        fprintf(stderr, "[V2O] vt=%d pdisp=%p\n", (int)v.vt, v.pdispVal);
        fflush(stderr);
    }
    if (v.vt == vb6_vtDispatch) return v.pdispVal;
    return NULL;
}

// P8.4: Variant清理 - 释放内含BSTR等资源
void vb6_VariantClear(vb6_VARIANT* v) {
    if (!v) return;
    // 释放BSTR
    if (v->vt == vb6_vtBSTR && v->bstrVal) {
        vb6_BSTR_Free(v->bstrVal);
        v->bstrVal = NULL;
    }
    // 释放IDispatch指针
    if (v->vt == vb6_vtDispatch && v->pdispVal) {
        vb6_ReleaseObject(&v->pdispVal);
        v->pdispVal = NULL;
    }
#ifndef _WIN64
    // Fix <vbeclipse> rev11: x86 的 DECIMAL 由 pdecVal 指向外部缓冲, 需释放
    if (v->vt == vb6_vtDecimal && v->pdecVal) {
        CoTaskMemFree(v->pdecVal);
        v->pdecVal = NULL;
    }
#endif
    v->vt = vb6_vtEmpty;
}

// P8.4: Variant深拷贝 - 复制BSTR等需要独立所有权的资源
void vb6_VariantCopy(vb6_VARIANT* dst, const vb6_VARIANT* src) {
    if (!dst || !src) return;
    *dst = *src;  // 浅拷贝
    // BSTR需要深拷贝
    if (src->vt == vb6_vtBSTR && src->bstrVal) {
        dst->bstrVal = vb6_BSTR_FromBSTR(src->bstrVal);
    }
    // Dispatch需要AddRef
    if (src->vt == vb6_vtDispatch && src->pdispVal) {
        // Fix <vbeclipse>: 原先注释自认 "COM AddRef would go here; just copy pointer" ——
        // 但 vb6_VariantClear 会对 vb6_vtDispatch 调 vb6_ReleaseObject, 所以浅拷贝 + 双方各
        // Clear 一次 = 过度释放 (0xC0000374)。与 vb6_VariantObject 同口径补 AddRef,
        // 使「每个持有 Variant 的槽位各自持有一份引用」成立。
        dst->pdispVal = src->pdispVal;
        vb6_ComAddRefDispatch(dst->pdispVal);
    }
}

// Fix <vbeclipse> rev16: Variant **槽位赋值**("接管"语义) —— 先释放 dst 旧内容, 再把
// src 的值深拷贝进去。与 vb6_VariantCopy 的唯一区别就是那句 Clear; BSTR 另分配 /
// Dispatch AddRef / x86 DECIMAL 另分配全部复用同一口径, 保证"每个槽各自持有一份
// 所有权"成立。
//
// 为什么必须深拷贝: 原先 Variant 数组元素赋值发的是
//   `slot = vb6_VariantFromValue((*Item))`
// 而 vb6_VariantFromValue 对 vb6_VARIANT 走 _Generic 的 vb6_VariantIdentity ——
// 纯结构体浅拷贝, 槽与调用方实参**共用同一只 BSTR / 同一个 Dispatch**。调用方
// (ByRef Variant 形参的宿主) 下一次 vb6_VariantClear(&v) 就把它 free 掉, 随后
// 同尺寸分配又复用那个地址 ⇒ 每个槽都读出"最后一次写入的值"。
// 实测 tests/ve_list (工程内 List.cls 形态): idx=0/1/2 全答 "Gamma", 按 key 也全答
// "Gamma"; 真工程 play78: Folder.Views.Item(0) 取回 1 字符垃圾 →
// m_Views.Item(<垃圾>) 落空 → NULL 解引用 (av read 0x4 @ vb6_View_prop_get_ViewId)。
// 对照: 同一份代码里的 m_Keys (String 数组) 走 vb6_BSTR_Assign 深拷贝, 所以它一直是对的。
void vb6_VariantAssign(vb6_VARIANT* dst, vb6_VARIANT src) {
    if (!dst) return;
    vb6_VariantClear(dst);        /* 释放旧值 (BSTR / Dispatch / x86 DECIMAL) */
    vb6_VariantCopy(dst, &src);   /* *dst = src + BSTR 另分配 / Dispatch AddRef */
#ifndef _WIN64
    /* x86 的 DECIMAL 由 pdecVal 指向外部缓冲: 上面 Copy 只搬了指针, 这里必须另分配
       一只, 否则 src 的宿主释放时会把 dst 手里的同一块 free 掉 (二次 free)。 */
    if (src.vt == vb6_vtDecimal && src.pdecVal) {
        DECIMAL* p = (DECIMAL*)CoTaskMemAlloc(sizeof(DECIMAL));
        if (p) { *p = *src.pdecVal; dst->pdecVal = p; }
        else   { dst->vt = vb6_vtEmpty; dst->pdecVal = NULL; }
    }
#endif
}

// ============================================================
// 错误处理 (MVP: 全局标志 + setjmp/longjmp)
// ============================================================

// 全局Err对象
typedef struct vb6_ErrObject {
    int32_t number;
    BSTR description;
    BSTR source;
    // Fix <vbeclipse> 2026-10-06: Err.LastDllError 快照 —— DLL 调用返回那一刻的
    // GetLastError(), 由 vb6_ErrSetLastDllError 在调用点捕获; 访问 Err.LastDllError
    // 返回此快照, 而非访问那一刻的 GetLastError() (中间 RTL/打印会重置 last-error)。
    int32_t lastDllError;
} vb6_ErrObject;

static vb6_ErrObject vb6_err = {0};

// 全局错误处理状态 (由cgen生成的代码直接使用)
int32_t vb6_err_resume_next = 0;
int32_t vb6_err_jmp_active = 0;
void* vb6_err_handler_label = NULL;

// 错误跳转缓冲区 (支持On Error GoTo label)
#include <setjmp.h>
jmp_buf* vb6_error_jmp_ptr = NULL;
int32_t vb6_error_jmp_set = 0;

// P14.1.2: Resume恢复点跟踪
int32_t vb6_err_resume_point = 0;
int32_t vb6_err_resume_next_point = 0;
int32_t vb6_err_dispatch = 0;
int32_t vb6_err_in_handler = 0;

// P12.3: On Error嵌套栈 — 保存/恢复错误处理状态
typedef struct vb6_ErrFrame {
    jmp_buf* jmp_ptr;
    int32_t jmp_set;
    int32_t jmp_active;
    int32_t resume_next;
    int32_t resume_point;          // P14.1.2
    int32_t resume_next_point;     // P14.1.2
    int32_t in_handler;            // P14.1.2
} vb6_ErrFrame;

static vb6_ErrFrame vb6_err_stack[VB6_ERR_STACK_SIZE];
static int32_t vb6_err_stack_top = 0;

void vb6_SaveErrState(void) {
    if (vb6_err_stack_top < VB6_ERR_STACK_SIZE) {
        vb6_err_stack[vb6_err_stack_top].jmp_ptr = vb6_error_jmp_ptr;
        vb6_err_stack[vb6_err_stack_top].jmp_set = vb6_error_jmp_set;
        vb6_err_stack[vb6_err_stack_top].jmp_active = vb6_err_jmp_active;
        vb6_err_stack[vb6_err_stack_top].resume_next = vb6_err_resume_next;
        vb6_err_stack[vb6_err_stack_top].resume_point = vb6_err_resume_point;
        vb6_err_stack[vb6_err_stack_top].resume_next_point = vb6_err_resume_next_point;
        vb6_err_stack[vb6_err_stack_top].in_handler = vb6_err_in_handler;
        vb6_err_stack_top++;
    }
}

void vb6_RestoreErrState(void) {
    if (vb6_err_stack_top > 0) {
        vb6_err_stack_top--;
        vb6_error_jmp_ptr = vb6_err_stack[vb6_err_stack_top].jmp_ptr;
        vb6_error_jmp_set = vb6_err_stack[vb6_err_stack_top].jmp_set;
        vb6_err_jmp_active = vb6_err_stack[vb6_err_stack_top].jmp_active;
        vb6_err_resume_next = vb6_err_stack[vb6_err_stack_top].resume_next;
        vb6_err_resume_point = vb6_err_stack[vb6_err_stack_top].resume_point;
        vb6_err_resume_next_point = vb6_err_stack[vb6_err_stack_top].resume_next_point;
        vb6_err_in_handler = vb6_err_stack[vb6_err_stack_top].in_handler;
    }
}

int32_t vb6_ErrNumber(void) { return vb6_err.number; }
BSTR vb6_ErrDescription(void) { return vb6_err.description; }
// Fix <vbeclipse> 2026-10-06: Err.Clear 清空 **全部** 属性 —— 含 LastDllError。
// VB6/VBA 文档 "Clear Method (Err Object)" 的属性表逐项列出 Clear 后的取值:
//   Description ""  HelpContext 0  HelpFile ""  LastDLLError 0  Number 0  Source ""
// 且 Clear 会被 Resume / Exit Sub|Function|Property / On Error 语句**自动**调用
// (本 RTL 里 Resume 已走 vb6_ErrClear, 见 cgen_jumps.cpp)。漏掉 lastDllError 会让
// 快照跨过一次 Clear 存活, 与 VB6 不符。
void vb6_ErrClear(void) {
    vb6_err.number = 0; vb6_err.description = NULL; vb6_err.source = NULL;
    vb6_err.lastDllError = 0;
}

// P21-27: Erl — 出错行号 (声明见 vb6rtl_class_com.h)
// 行号嵌入机制未实现 (cgen 不生成 VB 行号标签), 按 VB6 语义返回 0 —— VB6 中源码
// 不带行号时 Erl 同样返回 0。此前只有声明没有定义: 任何用到 Erl 的工程都会在
// 链接期报 LNK2019: unresolved external symbol vb6_Erl (CI vbman demo 实证)。
int32_t vb6_Erl(void) { return 0; }

BSTR vb6_ErrSource(void) { return vb6_err.source; }

// Fix <vbeclipse> 2026-10-06: Err.LastDllError 快照存取 (见 vb6_ErrObject.lastDllError)
int32_t vb6_ErrLastDllError(void) { return vb6_err.lastDllError; }
void vb6_ErrSetLastDllError(int32_t code) { vb6_err.lastDllError = code; }

void vb6_ErrRaise(int32_t errNum, BSTR source, BSTR description) {
    vb6_err.number = errNum;
    vb6_err.source = source;
    vb6_err.description = description;
    if (vb6_err_resume_next) return;
    if (vb6_err_jmp_active && vb6_error_jmp_set && vb6_error_jmp_ptr) {
        vb6_err_in_handler = 1;
        longjmp(*vb6_error_jmp_ptr, errNum);
    }
    // 未处理错误: 显示消息并退出
    // 定位调用点用: C3_COM_TRACE 下打印返回地址 (RVA 可对 -g 产出的 .map 符号化)
    if (GetEnvironmentVariableA("C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_ERR] unhandled raise %d from %p\n",
                (int)errNum, _ReturnAddress());
        fflush(stderr);
    }
    fwprintf(stderr, L"Unhandled error %d", (int)errNum);
    if (description) fwprintf(stderr, L": %s", description);
    fwprintf(stderr, L"\n");
    ExitProcess(errNum);
}

void vb6_ErrRaiseNumber(int32_t errNum) {
    vb6_ErrRaise(errNum, NULL, NULL);
}

// Task #44 (SSTabEx): VB6 '/' 与 Mod 的零除数语义是运行期错误 11 ("Division by zero"),
// 不是 IEEE inf。除数为字面量 0 时 cgen 强制走这两个 helper:
//  (a) 两侧都是常量时 MSVC 会常量折叠 → C2124 被零除 (frmTest.c:1569 InIde
//      "Debug.Print 1 / 0" 实证); 换成函数调用即不可折叠。
//  (b) 走 vb6_ErrRaise: On Error Resume Next 下静默置 Err.Number=11 —— InIde 的
//      "除零探测错误处理"技巧依赖此行为; 无错误处理时按 VB6 弹框/退出。
// 账 (除零补全): cgen 现在对 **变量除数** 也统一走 vb6_Num_Div / vb6_Num_Mod /
// vb6_IntDiv (三者都在 b==0 时 vb6_ErrRaiseNumber(11)), 此前变量除数落裸 C
// 除法得 IEEE inf / 静默 0 的缺口已闭合。
double vb6_Num_Div(double a, double b) {
    if (b == 0.0) { vb6_ErrRaiseNumber(11); return 0.0; }
    return a / b;
}

int32_t vb6_Num_Mod(int32_t a, int32_t b) {
    if (b == 0) { vb6_ErrRaiseNumber(11); return 0; }
    return a % b;
}
// ============================================================
// P24-Bug2: Variant比较函数
// 简化VB6语义: 两端都是字符串→字符串比较, 否则→Double数值比较
// ============================================================

static double vb6_VarToDouble_internal(vb6_VARIANT* v) {
    if (!v) return 0.0;
    switch ((vb6_vartype)v->vt) {
        case (vb6_vartype)VT_I2: return (double)v->iVal;
        case (vb6_vartype)VT_I4: return (double)v->lVal;
        case (vb6_vartype)VT_R4: return (double)v->fltVal;
        case (vb6_vartype)VT_R8: return v->dblVal;
        case (vb6_vartype)VT_BOOL: return v->boolVal ? -1.0 : 0.0;
        case (vb6_vartype)VT_BSTR: {
            if (!v->bstrVal) return 0.0;
            return wcstod(v->bstrVal, NULL);
        }
        default: return 0.0;
    }
}

static int32_t vb6_VarIsString(vb6_VARIANT* v) {
    return v && v->vt == (vb6_vartype)VT_BSTR;
}

int32_t vb6_VarCmpEq(vb6_VARIANT* a, vb6_VARIANT* b) {
    if (vb6_VarIsString(a) && vb6_VarIsString(b)) {
        // Fix 092v: 字符串 Variant 比较统一返回 VB6 Boolean (-1/0), 与数值分支一致.
        int eq = (a->bstrVal && b->bstrVal) ? (wcscmp(a->bstrVal, b->bstrVal) == 0) : (a->bstrVal == b->bstrVal);
        return eq ? -1 : 0;
    }
    double da = vb6_VarToDouble_internal(a);
    double db = vb6_VarToDouble_internal(b);
    return (da == db) ? -1 : 0;
}
int32_t vb6_VarCmpNe(vb6_VARIANT* a, vb6_VARIANT* b) {
    if (vb6_VarIsString(a) && vb6_VarIsString(b)) {
        int ne = (a->bstrVal && b->bstrVal) ? (wcscmp(a->bstrVal, b->bstrVal) != 0) : (a->bstrVal != b->bstrVal);
        return ne ? -1 : 0;
    }
    double da = vb6_VarToDouble_internal(a);
    double db = vb6_VarToDouble_internal(b);
    return (da != db) ? -1 : 0;
}
int32_t vb6_VarCmpLt(vb6_VARIANT* a, vb6_VARIANT* b) {
    if (vb6_VarIsString(a) && vb6_VarIsString(b)) {
        int lt = (a->bstrVal && b->bstrVal) ? (wcscmp(a->bstrVal, b->bstrVal) < 0) : 0;
        return lt ? -1 : 0;
    }
    return (vb6_VarToDouble_internal(a) < vb6_VarToDouble_internal(b)) ? -1 : 0;
}
int32_t vb6_VarCmpGt(vb6_VARIANT* a, vb6_VARIANT* b) {
    if (vb6_VarIsString(a) && vb6_VarIsString(b)) {
        int gt = (a->bstrVal && b->bstrVal) ? (wcscmp(a->bstrVal, b->bstrVal) > 0) : 0;
        return gt ? -1 : 0;
    }
    return (vb6_VarToDouble_internal(a) > vb6_VarToDouble_internal(b)) ? -1 : 0;
}
int32_t vb6_VarCmpLe(vb6_VARIANT* a, vb6_VARIANT* b) {
    if (vb6_VarIsString(a) && vb6_VarIsString(b)) {
        int le = (a->bstrVal && b->bstrVal) ? (wcscmp(a->bstrVal, b->bstrVal) <= 0) : 0;
        return le ? -1 : 0;
    }
    return (vb6_VarToDouble_internal(a) <= vb6_VarToDouble_internal(b)) ? -1 : 0;
}
int32_t vb6_VarCmpGe(vb6_VARIANT* a, vb6_VARIANT* b) {
    if (vb6_VarIsString(a) && vb6_VarIsString(b)) {
        int ge = (a->bstrVal && b->bstrVal) ? (wcscmp(a->bstrVal, b->bstrVal) >= 0) : 0;
        return ge ? -1 : 0;
    }
    return (vb6_VarToDouble_internal(a) >= vb6_VarToDouble_internal(b)) ? -1 : 0;
}

// Variant vs Long (常见场景: If v > 0 Then)
int32_t vb6_VarCmpLongEq(vb6_VARIANT* a, int32_t b) { vb6_VARIANT vb; memset(&vb, 0, sizeof(vb)); vb.vt = (vb6_vartype)VT_I4; vb.lVal = b; return vb6_VarCmpEq(a, &vb); }
int32_t vb6_VarCmpLongNe(vb6_VARIANT* a, int32_t b) { vb6_VARIANT vb; memset(&vb, 0, sizeof(vb)); vb.vt = (vb6_vartype)VT_I4; vb.lVal = b; return vb6_VarCmpNe(a, &vb); }
int32_t vb6_VarCmpLongLt(vb6_VARIANT* a, int32_t b) { vb6_VARIANT vb; memset(&vb, 0, sizeof(vb)); vb.vt = (vb6_vartype)VT_I4; vb.lVal = b; return vb6_VarCmpLt(a, &vb); }
int32_t vb6_VarCmpLongGt(vb6_VARIANT* a, int32_t b) { vb6_VARIANT vb; memset(&vb, 0, sizeof(vb)); vb.vt = (vb6_vartype)VT_I4; vb.lVal = b; return vb6_VarCmpGt(a, &vb); }
int32_t vb6_VarCmpLongLe(vb6_VARIANT* a, int32_t b) { vb6_VARIANT vb; memset(&vb, 0, sizeof(vb)); vb.vt = (vb6_vartype)VT_I4; vb.lVal = b; return vb6_VarCmpLe(a, &vb); }
int32_t vb6_VarCmpLongGe(vb6_VARIANT* a, int32_t b) { vb6_VARIANT vb; memset(&vb, 0, sizeof(vb)); vb.vt = (vb6_vartype)VT_I4; vb.lVal = b; return vb6_VarCmpGe(a, &vb); }


void vb6_RaiseError(int32_t errNum, BSTR description) {
    vb6_err.number = errNum;
    vb6_err.description = description;
    if (vb6_err_resume_next) {
        // On Error Resume Next: 忽略错误, 继续执行
        return;
    }
    if (vb6_err_jmp_active && vb6_error_jmp_set && vb6_error_jmp_ptr) {
        // On Error GoTo label: longjmp 跳到 setjmp 点
        vb6_err_in_handler = 1;  // P14.1.2: 标记进入错误处理器
        longjmp(*vb6_error_jmp_ptr, errNum);
    }
    // 未设置错误处理: 只要有**可写的 std 句柄** (控制台 / 管道 / 文件) 就写 stderr, 只有句柄都不可用
    // (从资源管理器双击起的 GUI 程序) 才弹 VB6 风格对话框。
    // 判据不能用 GetConsoleWindow(): GUI 子系统 exe 从 cmd 里起时 std 句柄**就是**那份控制台,
    // 而它返回 NULL —— 会把本该打到 stderr 的错误变成模态框; 更不能只看"是不是控制台": 批处理/套件
    // 把 stdout/stderr 重定向成管道时, 弹框没人点确定 ⇒ 进程永远卡住 (CI 上表现为 5s 超时被杀)。
    {
        HANDLE hErr = GetStdHandle(STD_ERROR_HANDLE);
        HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);
        int canWrite = 0;
        if (hErr && hErr != INVALID_HANDLE_VALUE && GetFileType(hErr) != FILE_TYPE_UNKNOWN) canWrite = 1;
        else if (hOut && hOut != INVALID_HANDLE_VALUE && GetFileType(hOut) != FILE_TYPE_UNKNOWN) canWrite = 1;
        if (canWrite) {
            wchar_t line[600];
            int n = swprintf(line, 600, L"Unhandled VB6 Error #%d: %ls\n", errNum,
                             description ? description : L"(no description)");
            if (n > 0) vb6_ConWriteErrW(line, n);
        } else {
            // 无处可写 (双击启动的 GUI 程序): 弹 VB6 风格错误对话框
            wchar_t msg[512];
            swprintf(msg, 512, L"Run-time error '%d':\n%ls",
                     errNum, description ? description : L"(no description)");
            MessageBoxW(NULL, msg, L"VB6 Runtime Error", MB_ICONERROR | MB_OK);
        }
    }
    exit(errNum);
}

// 账 #181: 见 vb6rtl_runtime.h 里那条注释 —— 三个出口每进程各记一次。
static volatile long vb6_crashClaim[4] = { 0, 0, 0, 0 };

int vb6_CrashTraceClaim(int slot) {
    if (slot < 0 || slot >= 4) return 1;           /* 未知槽位不拦 (宁可多记一份) */
    return InterlockedCompareExchange(&vb6_crashClaim[slot], 1, 0) == 0;
}
