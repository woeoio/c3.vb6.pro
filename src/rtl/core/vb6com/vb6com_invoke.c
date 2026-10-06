// vb6com_invoke.c - vb6com 模块拆分: DISPID 缓存 + 方法调用 + 属性读写
// 由 vb6com.c 按 COM 调用层次拆分而来 (纯搬移, 零行为改动)

#include "vb6com.h"
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "vb6com_internal.h"
#include "vb6forms.h"   /* Fix 112: 宿主对象分派 */


// ============================================================
// COM后期绑定: IDispatch::Invoke
// ============================================================

// DISPID缓存 (MVP: 线性查找)
typedef struct vb6_DispidCacheEntry {
    void*   obj;
    wchar_t name[64];
    DISPID  dispid;
} vb6_DispidCacheEntry;

#define VB6_DISPID_CACHE_SIZE 256
static vb6_DispidCacheEntry vb6_dispid_cache[VB6_DISPID_CACHE_SIZE];
static int32_t vb6_dispid_cache_count = 0;

// Fix 191: 按名后期绑定的接收者必须真的是 COM 对象, 否则 `pDisp->lpVtbl->...`
// 就是一次野指针解引用 —— 进程直接 0xC0000005 死掉, 而不是像 VB6 那样把
// "对象不支持此属性或方法" 交给 On Error 处理。
// 触发实例 (VBFlexGridDemo, 左键点网格 → WM_SETFOCUS → VTableHandle.ActivateIPAO):
//   VTableIPAO(0 To 9) As LongPtr 被按 Variant 载体分配 (见 cgen_expr_array.cpp
//   的 mapSaElemType 缺 LongPtr 分支), 于是 VarPtr(VTableIPAO(0)) 指向的是
//   vb6_VARIANT 头部; 伪 vtable 的第 5 槽读出来是 vt=VT_I4=3, 再 +0x28 →
//   读 0x2b 崩溃。工程里所有手写 vtable 子类化 (VTableControl/VTablePPB/
//   VTableEnumVARIANT…) 都走同一条路, 一个坏接收者就能杀掉整个进程。
// 判据 (刻意不用"vtable 必须在模块镜像内": VTableHandle.bas 的伪 vtable 是
// `VTableIPAO() As LongPtr` 数组, 载体在堆上, 那样会把合法对象一并拒掉):
//   对象指针可读 → 首槽 vtable 可读且容得下前 7 槽 → 槽 0/1/2/5/6 指向可执行内存。
// 放在 DISPID 缓存查找之后, 命中缓存的调用不受影响。
static int32_t vb6_ComIsCodePtr(const void* p) {
    MEMORY_BASIC_INFORMATION mbi;
    if (!p || VirtualQuery(p, &mbi, sizeof(mbi)) == 0) return 0;
    return (mbi.Protect & (PAGE_EXECUTE | PAGE_EXECUTE_READ
                         | PAGE_EXECUTE_READWRITE | PAGE_EXECUTE_WRITECOPY)) != 0;
}

int32_t vb6_ComIsDispatchable(const void* disp) {
    if (!disp) return 0;
    // Fix <vbeclipse>: 窗口句柄 (HWND) 是内核索引, 不是指向 IDispatch 对象的指针。
    // 宿主 Controls 枚举 (`For Each ctrl In Controls`) 把子控件的**裸 HWND** 放进
    // VT_DISPATCH 槽; 下面 `*(void**)disp` 会把 HWND 当指针读堆内存, 若那块内存
    // 恰好看着像 vtable 就误判"可分派" → 调伪 AddRef → 跳到垃圾地址崩
    // (ucTabStrip.UserControl_Resize 实测, PC=野值)。真窗口句柄一律非 COM 分派对象,
    // 由宿主模型 (IsHostObject/IsWindow) 单独应答, 这里直接判否。
    if (IsWindow((HWND)disp)) return 0;
    if (IsBadReadPtr(disp, sizeof(void*))) return 0;
    void* vtbl = *(void**)disp;
    /* IDispatch: QueryInterface/AddRef/Release/GetTypeInfoCount/
     * GetTypeInfo/GetIDsOfNames/Invoke — 前 7 槽必须在可执行页里 */
    if (!vtbl || IsBadReadPtr(vtbl, 7 * sizeof(void*))) return 0;
    void** slot = (void**)vtbl;
    return vb6_ComIsCodePtr(slot[0]) && vb6_ComIsCodePtr(slot[1])
        && vb6_ComIsCodePtr(slot[2]) && vb6_ComIsCodePtr(slot[5])
        && vb6_ComIsCodePtr(slot[6]);
}

static DISPID vb6_getDispid(IDispatch* pDisp, const wchar_t* name) {
    // 先查缓存
    for (int32_t i = 0; i < vb6_dispid_cache_count; i++) {
        if (vb6_dispid_cache[i].obj == (void*)pDisp &&
            wcscmp(vb6_dispid_cache[i].name, name) == 0) {
            return vb6_dispid_cache[i].dispid;
        }
    }
    if (!vb6_ComIsDispatchable(pDisp)) {
        if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
            fwprintf(stderr, L"[C3_COM] getDispid(\"%ls\"): NOT dispatchable (rejected by vb6_ComIsDispatchable)\n", name);
            fflush(stderr);
        }
        return DISPID_UNKNOWN;
    }

    /* Fix 160-com: 免注册 typelib 驱动 IDispatch (Chilkat 等) 在 TLB 未注册时,
     * 自身 GetIDsOfNames 返回 TYPE_E_LIBNOTREGISTERED。改从 C3 预载的 ITypeInfo
     * (LoadTypeLibEx REGKIND_NONE, 不写注册表) 解析 DISPID —— 真正免注册。
     * 对象→ITypeInfo 映射在 comLibCreateLocal 创建时登记; 未命中返回 NULL,
     * 落回下方常规 GetIDsOfNames, 行为不变。 */
    {
        ITypeInfo* ti160 = vb6_ComLibLookupTypeInfo((void*)pDisp);
        if (ti160) {
            LPOLESTR tiNames160[1] = { (LPOLESTR)name };
            MEMBERID memid160 = 0;
            HRESULT hrTI = ti160->lpVtbl->GetIDsOfNames(ti160, tiNames160, 1, &memid160);
            ti160->lpVtbl->Release(ti160);
            if (SUCCEEDED(hrTI)) {
                DISPID dispid = (DISPID)memid160;
                if (vb6_dispid_cache_count < VB6_DISPID_CACHE_SIZE) {
                    vb6_dispid_cache[vb6_dispid_cache_count].obj = (void*)pDisp;
                    wcsncpy(vb6_dispid_cache[vb6_dispid_cache_count].name, name, 63);
                    vb6_dispid_cache[vb6_dispid_cache_count].name[63] = L'\0';
                    vb6_dispid_cache[vb6_dispid_cache_count].dispid = dispid;
                    vb6_dispid_cache_count++;
                }
                return dispid;
            }
        }
    }

    // 调用GetIDsOfNames.
    // 不少COM对象(如 Chilkat ActiveX)的 IDispatch::GetIDsOfNames 只认
    // LOCALE_SYSTEM_DEFAULT(0x800) 或 LOCALE_NEUTRAL(0), 对 LOCALE_USER_DEFAULT(0x400)
    // 返回 DISP_E_UNKNOWNNAME —— 表现就是属性/方法全部 "not found". 依次在多个
    // LCID 上重试, 任一命中即用, 避免这类对象静默失效. 顺序: 先 USER(标准), 再
    // SYSTEM / NEUTRAL, 兼容标准对象与严格对象.
    DISPID dispid;
    LPOLESTR names[1] = { (LPOLESTR)name };
    static const LCID tryLcids[] = { LOCALE_USER_DEFAULT, LOCALE_SYSTEM_DEFAULT, LOCALE_NEUTRAL, 0 };
    HRESULT hr = DISP_E_UNKNOWNNAME;
    for (int li = 0; li < 4; li++) {
        hr = pDisp->lpVtbl->GetIDsOfNames(pDisp, &IID_NULL, names, 1,
                                          tryLcids[li], &dispid);
        if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
            fwprintf(stderr, L"[C3_COM] getDispid(\"%ls\") lcid=0x%04lX hr=0x%08lX dispid=%ld\n",
                     name, (unsigned long)tryLcids[li], (unsigned long)hr, (long)dispid);
            fflush(stderr);
        }
        if (SUCCEEDED(hr)) {
            if (li != 0 && GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
                fwprintf(stderr, L"[C3_COM] GetIDsOfNames(\"%ls\") ok with lcid 0x%04lX (USER failed)\n",
                         name, (unsigned long)tryLcids[li]);
                fflush(stderr);
            }
            break;
        }
    }
    if (FAILED(hr)) return DISPID_UNKNOWN;

    // 存入缓存
    if (vb6_dispid_cache_count < VB6_DISPID_CACHE_SIZE) {
        vb6_dispid_cache[vb6_dispid_cache_count].obj = (void*)pDisp;
        wcsncpy(vb6_dispid_cache[vb6_dispid_cache_count].name, name, 63);
        vb6_dispid_cache[vb6_dispid_cache_count].name[63] = L'\0';
        vb6_dispid_cache[vb6_dispid_cache_count].dispid = dispid;
        vb6_dispid_cache_count++;
    }

    return dispid;
}

// ============================================================
// Fix 160-com: 免注册分派辅助 (ITypeInfo 驱动, 完全不依赖注册表)
// ============================================================

/* 免注册 GetIDsOfNames: 若对象是 C3 ComLib 免注册创建的, 优先用预载的
 * ITypeInfo (LoadTypeLibEx REGKIND_NONE 内存副本) 解析 DISPID, 否则退回
 * 对象自身 GetIDsOfNames。供 For Each 等直接调用路径使用。 */
HRESULT vb6_ComGetIdsOfNames(IDispatch* pDisp, const wchar_t* name, DISPID* pDispid) {
    ITypeInfo* ti = vb6_ComLibLookupTypeInfo((void*)pDisp);
    if (ti) {
        LPOLESTR tn[1] = { (LPOLESTR)name };
        MEMBERID mid = 0;
        HRESULT hr = ti->lpVtbl->GetIDsOfNames(ti, tn, 1, &mid);
        ti->lpVtbl->Release(ti);
        if (SUCCEEDED(hr)) { *pDispid = (DISPID)mid; return hr; }
    }
    LPOLESTR on[1] = { (LPOLESTR)name };
    return pDisp->lpVtbl->GetIDsOfNames(pDisp, &IID_NULL, on, 1, LOCALE_USER_DEFAULT, pDispid);
}

/* 免注册 Invoke 总入口: 若对象是 C3 ComLib 免注册创建的, 走 ITypeInfo::Invoke
 * (内存中 TLB 副本), 绕开对象自身因 TLB 未注册而失败的 IDispatch::Invoke;
 * 普通注册表组件无 TI 映射, 落回 IDispatch::Invoke, 行为不变。 */
HRESULT vb6_ComInvoke(IDispatch* pDisp, DISPID dispid, WORD wFlags,
                      DISPPARAMS* dp, VARIANT* result, EXCEPINFO* excep, UINT* argErr) {
    ITypeInfo* ti = vb6_ComLibLookupTypeInfo((void*)pDisp);
    if (ti) {
        HRESULT hr = ti->lpVtbl->Invoke(ti, (void*)pDisp, dispid, wFlags, dp, result, excep, argErr);
        ti->lpVtbl->Release(ti);
        return hr;
    }
    return pDisp->lpVtbl->Invoke(pDisp, dispid, &IID_NULL, LOCALE_USER_DEFAULT, wFlags, dp, result, excep, argErr);
}

// COM方法调用 (返回VARIANT*, 调用方需vb6_ComVarClear释放)
// args_void: VARIANT*[] (指向已分配VARIANT的指针数组), 每个元素由vb6_ComPackXxx分配
void* vb6_ComCall(void* disp, const wchar_t* methodName,
                  void* args_void, int32_t argc) {
    VARIANT** args = (VARIANT**)args_void;
    if (!disp) return NULL;
    /* Fix 112: 宿主对象 (窗体/控件 HWND, Controls 集合, Font 代理) 不是 IDispatch,
     * 直接解引用 lpVtbl 会 AV. 用 Win32 语义的宿主分派应答. */
    if (vb6_Host_IsHostObject(disp)) {
        char hout[64];   /* vb6_VARIANT (vb6com 单元用 Windows VARIANT, 不见 vb6_VARIANT 类型) */
        vb6_Host_Call(disp, methodName, argc, args_void, hout);
        if (args) {
            for (int32_t hi = 0; hi < argc; hi++) { if (args[hi]) free(args[hi]); }
        }
        VARIANT* hres = (VARIANT*)calloc(1, sizeof(VARIANT));
        vb6_Host_ToWinVariant(&hout, hres);
        vb6_Host_ClearVariant(&hout);
        return (void*)hres;
    }
    /* Fix 191: IUnknown 三法在**所有** COM 接口的 vtable 里固定在槽 0/1/2,
     * 自定义接口 (IOleInPlaceActiveObject / IOleObject 这类无 IDispatch 的) 也一样。
     * 而按名后期绑定要先 GetIDsOfNames —— 那是 IDispatch 的槽 5。对手写伪 vtable
     * (VTableHandle.bas 的 VTableIPAO[] 槽位: QI/AddRef/Release/GetWindow/
     * ContextSensitiveHelp/TranslateAccelerator/…) 就等于把
     * IOleIPAO_TranslateAccelerator 当 GetIDsOfNames 调用, 参数全错位 → AV。
     * VB6 编译 `obj.AddRef` 本来就是 vtbl[1](obj), 这里按同一口径直发。 */
    if (vb6_ComIsDispatchable(disp) &&
        (wcscmp(methodName, L"AddRef") == 0 || wcscmp(methodName, L"Release") == 0)) {
        int32_t isAddRef191 = (wcscmp(methodName, L"AddRef") == 0);
        /* 账 #183: 槽位取自**对象首字里那个 vtable**, 不是对象首字本身。旧写法
         * `((void**)disp)[1]` 读的是 disp+8 —— 那是对象自己的第二个字段, 在
         * VTableHandle.bas 手搭的伪 IPAO 对象上恰好是 RefCount=1 ⇒ `call 1` 跳飞。
         * vb6_ComIsDispatchable 已保证 *(void**)disp 是一张槽位全在可执行页的表,
         * 所以这里直接用 C 的 COM 写法, 与 vb6com.c 里 Release/AddRef 两处同形。 */
        IUnknown* pUnk191 = (IUnknown*)disp;
        LONG n191 = isAddRef191 ? (LONG)pUnk191->lpVtbl->AddRef(pUnk191)
                               : (LONG)pUnk191->lpVtbl->Release(pUnk191);
        VARIANT** uargs191 = (VARIANT**)args_void;
        if (uargs191) {
            for (int32_t ui = 0; ui < argc; ui++) { if (uargs191[ui]) free(uargs191[ui]); }
        }
        VARIANT* ures191 = (VARIANT*)calloc(1, sizeof(VARIANT));
        VariantInit(ures191);
        ures191->vt = VT_I4;
        ures191->lVal = n191;
        return (void*)ures191;
    }
    IDispatch* pDisp = (IDispatch*)disp;

    if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_COM] ComCall fallback to IDispatch: disp=%p isFont=%d method=%ls\n",
                disp, vb6_UC_IsFont(disp), methodName); fflush(stderr);
    }
    DISPID dispid = vb6_getDispid(pDisp, methodName);
    if (dispid == DISPID_UNKNOWN) {
        fwprintf(stderr, L"vb6_ComCall: method \"%ls\" not found\n", methodName);
        return NULL;
    }

    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    dp.cArgs = (UINT)argc;

    VARIANT* result = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(result);

    if (argc > 0) {
        dp.rgvarg = (VARIANTARG*)malloc((size_t)argc * sizeof(VARIANTARG));
        // COM参数逆序, 从VARIANT*数组复制VARIANT值
        for (int32_t i = 0; i < argc; i++) {
            VariantInit(&dp.rgvarg[argc - 1 - i]);
            if (args[i]) {
                dp.rgvarg[argc - 1 - i] = *args[i];
            }
        }
    }

    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    HRESULT hr = vb6_ComInvoke(pDisp, dispid, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, result, &excep, &argErr);

    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_COM] Call '%ls' hr=0x%08lX vt=%d\n",
                methodName, (unsigned long)hr, result ? result->vt : -1); fflush(stderr);

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComCall");
        // On Error Resume Next: continue with NULL result
        if (result) { VariantClear(result); free(result); result = NULL; }
    }

    // 清理参数副本 (dp.rgvarg是浅拷贝, VariantClear会释放内部的BSTR/对象)
    if (dp.rgvarg) {
        for (UINT i = 0; i < dp.cArgs; i++) {
            // Fix 160-com-byref: VT_BYREF 出参指向调用方存储, 引用归调用方所有,
            // 这里不能 VariantClear (否则会Release掉方法新写入的对象, 或释放调用方变量)。
            if (dp.rgvarg[i].vt & VT_BYREF) continue;
            VariantClear(&dp.rgvarg[i]);
        }
        free(dp.rgvarg);
    }

    // 清理打包的VARIANT参数结构体 (内容已通过dp.rgvarg的VariantClear释放)
    // 注意: dp.rgvarg[i] = *args[i] 是浅拷贝, BSTR/对象引用已被上面的VariantClear释放
    // 所以这里只free结构体, 不能再VariantClear (否则double free)
    if (args) {
        for (int32_t i = 0; i < argc; i++) {
            if (args[i]) {
                free(args[i]);  // 仅释放结构体, 内容已在dp.rgvarg清理
            }
        }
    }

    return (void*)result;
}


// P24-10: COM默认成员调用 (按DISPID直接调用, 跳过名称查找)
// 用于后期绑定: Dim obj As Object; obj(args) → DISPID_VALUE(0)
void* vb6_ComCallByDispid(void* disp, int32_t dispid,
                          void* args_void, int32_t argc) {
    VARIANT** args = (VARIANT**)args_void;
    if (!disp) return NULL;
    IDispatch* pDisp = (IDispatch*)disp;

    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    dp.cArgs = (UINT)argc;

    VARIANT* result = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(result);

    if (argc > 0) {
        dp.rgvarg = (VARIANTARG*)malloc((size_t)argc * sizeof(VARIANTARG));
        for (int32_t i = 0; i < argc; i++) {
            VariantInit(&dp.rgvarg[argc - 1 - i]);
            if (args[i]) {
                dp.rgvarg[argc - 1 - i] = *args[i];
            }
        }
    }

    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    HRESULT hr = vb6_ComInvoke(pDisp, (DISPID)dispid, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, result, &excep, &argErr);

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComCallByDispid");
        if (result) { VariantClear(result); free(result); result = NULL; }
    }

    if (dp.rgvarg) {
        for (UINT i = 0; i < dp.cArgs; i++) {
            // Fix 160-com-byref: ByRef 出参指向调用方存储, 不释放其引用。
            if (dp.rgvarg[i].vt & VT_BYREF) continue;
            VariantClear(&dp.rgvarg[i]);
        }
        free(dp.rgvarg);
    }
    if (args) {
        for (int32_t i = 0; i < argc; i++) {
            if (args[i]) { free(args[i]); }
        }
    }

    return (void*)result;
}
// COM属性Get (返回VARIANT*)
void* vb6_ComGetProp(void* disp, const wchar_t* propName) {
    if (!disp) return NULL;
    /* Fix 164z2-dbg: C3_COM_TRACE 入口日志 */
    if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_COM] GetProp entry: disp=%p isFont=%d isHost=%d name=%ls\n",
                disp, vb6_UC_IsFont(disp), vb6_Host_IsHostObject(disp), propName);
        fflush(stderr);
    }
    /* Fix 125: 字体结构体 (非 COM 对象) 直接读字段 — 生成代码把 `With <font>.Name`
       编译成 COM 读, 对结构体做 Invoke 会崩。 */
    if (vb6_UC_IsFont(disp)) {
        int32_t kind = -1;
        void* fp = vb6_UC_FontField(disp, propName, &kind);
        VARIANT* res = (VARIANT*)calloc(1, sizeof(VARIANT));
        if (!res) return NULL;
        if (fp) {
            switch (kind) {
                case 0: res->vt = VT_BSTR; res->bstrVal = SysAllocString(*(BSTR*)fp); break;
                case 1: res->vt = VT_R4;   res->fltVal = *(float*)fp; break;
                case 2: res->vt = VT_I2;   res->iVal = *(int16_t*)fp; break;
                default: res->vt = VT_I4;  res->lVal = *(int32_t*)fp; break;
            }
        }
        return (void*)res;
    }
    /* Fix 168: Extender 结构体同 Fix 125 字体 —— `With UserControl.Extender` 的
       .Width/.Height/.Align 被编译成对普通结构体的 COM 读, Invoke 即跳进 .data. */
    if (vb6_UC_IsExtender(disp)) {
        int32_t ekind = -1;
        void* efp = vb6_UC_ExtenderField(disp, propName, &ekind);
        VARIANT* eres = (VARIANT*)calloc(1, sizeof(VARIANT));
        if (!eres) return NULL;
        if (efp) {
            switch (ekind) {
                case 0: eres->vt = VT_BSTR; eres->bstrVal = SysAllocString(*(BSTR*)efp); break;
                case 1: eres->vt = VT_R4;   eres->fltVal = *(float*)efp; break;
                case 2: eres->vt = VT_I2;   eres->iVal = *(int16_t*)efp; break;
                default: eres->vt = VT_I4;  eres->lVal = *(int32_t*)efp; break;
            }
        }
        return (void*)eres;
    }
    /* Fix 112: 宿主对象分派 (见 vb6_ComCall 注释) */
    if (vb6_Host_IsHostObject(disp)) {
        char hout[64];
        vb6_Host_GetProp(disp, propName, hout);
        VARIANT* hres = (VARIANT*)calloc(1, sizeof(VARIANT));
        vb6_Host_ToWinVariant(&hout, hres);
        vb6_Host_ClearVariant(&hout);
        return (void*)hres;
    }
    /* 账 #182: 这条 "fallback to IDispatch" 以前打在字体/Extender/宿主三个岔口**之前**，
     * 于是三条都不走 IDispatch 的路径也照打不误 —— 用它定位崩溃时会把嫌疑引向
     * 根本没执行过的调用。挪到这里只在真要下 IDispatch 时打印。 */
    if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_COM] GetProp fallback to IDispatch: disp=%p isFont=%d name=%ls\n",
                disp, vb6_UC_IsFont(disp), propName);
        fflush(stderr);
    }
    IDispatch* pDisp = (IDispatch*)disp;

    DISPID dispid = vb6_getDispid(pDisp, propName);
    if (dispid == DISPID_UNKNOWN) {
        fwprintf(stderr, L"vb6_ComGetProp: property \"%ls\" not found\n", propName);
        return NULL;
    }

    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));

    VARIANT* result = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(result);

    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    HRESULT hr = vb6_ComInvoke(pDisp, dispid, (DISPATCH_METHOD | DISPATCH_PROPERTYGET), &dp, result, &excep, &argErr);

    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0)
        fprintf(stderr, "[C3_COM] GetProp '%ls' hr=0x%08lX vt=%d\n",
                propName, (unsigned long)hr, result ? result->vt : -1); fflush(stderr);

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComGetProp");
        if (result) { VariantClear(result); free(result); result = NULL; }
    }

    return (void*)result;
}


// COM属性Get带参数 (参数化属性读取, 如Dictionary.Item(key))
// 使用DISPATCH_METHOD|DISPATCH_PROPERTYGET组合标志
// 返回VARIANT* (调用方需vb6_ComVarClear释放)
void* vb6_ComGetPropArg(void* disp, const wchar_t* propName,
                        void* args_void, int32_t argc) {
    VARIANT** args = (VARIANT**)args_void;
    if (!disp) return NULL;
    IDispatch* pDisp = (IDispatch*)disp;

    DISPID dispid = vb6_getDispid(pDisp, propName);
    if (dispid == DISPID_UNKNOWN) {
        fwprintf(stderr, L"vb6_ComGetPropArg: property \"%ls\" not found\n", propName);
        return NULL;
    }

    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    dp.cArgs = (UINT)argc;

    VARIANT* result = (VARIANT*)calloc(1, sizeof(VARIANT));
    VariantInit(result);

    if (argc > 0) {
        dp.rgvarg = (VARIANTARG*)malloc((size_t)argc * sizeof(VARIANTARG));
        for (int32_t i = 0; i < argc; i++) {
            VariantInit(&dp.rgvarg[argc - 1 - i]);
            if (args[i]) {
                dp.rgvarg[argc - 1 - i] = *args[i];
            }
        }
    }

    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    HRESULT hr = vb6_ComInvoke(pDisp, dispid, DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, result, &excep, &argErr);

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComGetPropArg");
        if (result) { VariantClear(result); free(result); result = NULL; }
    }

    if (dp.rgvarg) {
        for (UINT i = 0; i < dp.cArgs; i++) {
            VariantClear(&dp.rgvarg[i]);
        }
        free(dp.rgvarg);
    }

    return (void*)result;
}

// COM属性Set (值类型)
// COM属性Set (值类型)
void vb6_ComSetProp(void* disp, const wchar_t* propName, void* value_void) {
    /* Fix 125: 字体结构体 (非 COM 对象) 直接写字段 —— 这是 `Property Set X_Font` 里
       `With m_X_Font: .Name = New_Font.Name ...` 的实现路径; 走 Invoke 必崩。 */
    if (disp && vb6_UC_IsFont(disp)) {
        int32_t kind = -1;
        void* fp = vb6_UC_FontField(disp, propName, &kind);
        if (fp && value_void) {
            VARIANT v = *(VARIANT*)value_void;
            switch (kind) {
                case 0: *(BSTR*)fp = (v.vt == VT_BSTR && v.bstrVal)
                                     ? SysAllocString(v.bstrVal) : NULL; break;
                case 1: *(float*)fp = (float)vb6_VariantToDouble(v); break;
                case 2: *(int16_t*)fp = (int16_t)vb6_VariantToDouble(v); break;
                default: *(int32_t*)fp = (int32_t)vb6_VariantToDouble(v); break;
            }
        }
        free(value_void);
        return;
    }
    /* Fix 168: Extender 结构体的 `With ...: .Width = ...` 写字段, 同字体路径 */
    if (disp && vb6_UC_IsExtender(disp)) {
        int32_t ekind = -1;
        void* efp = vb6_UC_ExtenderField(disp, propName, &ekind);
        if (efp && value_void) {
            VARIANT v = *(VARIANT*)value_void;
            switch (ekind) {
                case 0: *(BSTR*)efp = (v.vt == VT_BSTR && v.bstrVal)
                                   ? SysAllocString(v.bstrVal) : NULL; break;
                case 1: *(float*)efp = (float)vb6_VariantToDouble(v); break;
                case 2: *(int16_t*)efp = (int16_t)vb6_VariantToDouble(v); break;
                default: *(int32_t*)efp = (int32_t)vb6_VariantToDouble(v); break;
            }
        }
        free(value_void);
        return;
    }
    /* Fix 112: 宿主对象分派 (见 vb6_ComCall 注释) */
    if (disp && vb6_Host_IsHostObject(disp)) {
        char hin[64];
        vb6_Host_FromWinVariant(value_void, hin);
        vb6_Host_SetProp(disp, propName, hin);
        vb6_Host_ClearVariant(hin);
        return;
    }

    VARIANT value;
    VariantInit(&value);
    // 简化: 假设value是已打包的VARIANT; MVP阶段由cgen直接传递
    if (value_void) {
        value = *(VARIANT*)value_void;
    }
    if (!disp) { free(value_void); return; }
    IDispatch* pDisp = (IDispatch*)disp;

    DISPID dispid = vb6_getDispid(pDisp, propName);
    if (dispid == DISPID_UNKNOWN) {
        free(value_void);
        fwprintf(stderr, L"vb6_ComSetProp: property \"%ls\" not found\n", propName);
        return;
    }

    DISPID putId = DISPID_PROPERTYPUT;
    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    dp.cArgs = 1;
    dp.cNamedArgs = 1;
    dp.rgvarg = &value;
    dp.rgdispidNamedArgs = &putId;

    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    HRESULT hr = vb6_ComInvoke(pDisp, dispid, DISPATCH_PROPERTYPUT, &dp, NULL, &excep, &argErr);

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComSetProp");
    }
    free(value_void);  /* 释放ComPackXxx分配的堆VARIANT结构体 */
}

// P25: COM参数化属性Put (如dic.Item(key) = value)
// propName: 属性名 (如"Item"), args: 索引参数数组, argc: 索引参数个数
// value: 新值(已打包为堆VARIANT*, 同ComSetProp)
void vb6_ComSetPropArg(void* disp, const wchar_t* propName, void** args, int32_t argc, void* value_void) {
    VARIANT varValue;
    VariantInit(&varValue);
    if (value_void) varValue = *(VARIANT*)value_void;
    if (!disp) { free(value_void); return; }
    IDispatch* pDisp = (IDispatch*)disp;

    DISPID dispid = vb6_getDispid(pDisp, propName);
    if (dispid == DISPID_UNKNOWN) {
        free(value_void);
        fwprintf(stderr, L"vb6_ComSetPropArg: property \"%ls\" not found\n", propName);
        return;
    }

    /* 构建rgvarg: [索引参数..., value], 值在最后 */
    int32_t totalArgs = argc + 1;
    VARIANT* rgvarg = (VARIANT*)calloc(totalArgs, sizeof(VARIANT));
    /* 值参数放rgvarg[0] (DISPATCH反序), 索引参数依次放rgvarg[1..argc] */
    rgvarg[0] = varValue;
    for (int32_t i = 0; i < argc; i++) {
        rgvarg[i + 1] = *(VARIANT*)args[i];
    }

    DISPID putId = DISPID_PROPERTYPUT;
    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    dp.cArgs = totalArgs;
    dp.cNamedArgs = 1;
    dp.rgvarg = rgvarg;
    dp.rgdispidNamedArgs = &putId;

    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    HRESULT hr = vb6_ComInvoke(pDisp, dispid, DISPATCH_PROPERTYPUT, &dp, NULL, &excep, &argErr);

    /* 清理: 释放索引参数的堆VARIANT和索引数组, 以及value */
    for (int32_t i = 0; i < argc; i++) free(args[i]);
    free(rgvarg);
    free(value_void);

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComSetPropArg");
    }
}

// COM属性SetRef (对象引用 -- PROPERTYPUTREF)
void vb6_ComSetRef(void* disp, const wchar_t* propName, void* objRef) {
    if (!disp) return;
    IDispatch* pDisp = (IDispatch*)disp;

    // Fix <vbeclipse>: 宿主对象 (原生 vb6_cls_X* UC 实例 / HWND) 不是真 IDispatch,
    // 走下面的 getDispid+Invoke 必 "not found"。与 vb6_ComSetProp (line 526) 同口径:
    // 先把 objRef 包成 VT_DISPATCH 的 Win VARIANT, 转成 vb6_VARIANT 交给 Host_SetProp,
    // 后者落到 vb6_UC_OwnPropSet 的名字桥 (uc_hostmodel_setprop.inc)。
    if (vb6_Host_IsHostObject(disp)) {
        VARIANT winObj; VariantInit(&winObj);
        winObj.vt = VT_DISPATCH; winObj.pdispVal = (IDispatch*)objRef;
        char hin[64];
        vb6_Host_FromWinVariant(&winObj, hin);
        vb6_Host_SetProp(disp, propName, hin);
        vb6_Host_ClearVariant(hin);
        return;
    }

    DISPID dispid = vb6_getDispid(pDisp, propName);
    if (dispid == DISPID_UNKNOWN) {
        fwprintf(stderr, L"vb6_ComSetRef: property \"%ls\" not found\n", propName);
        return;
    }

    VARIANT varObj;
    VariantInit(&varObj);
    varObj.vt = VT_DISPATCH;
    varObj.pdispVal = (IDispatch*)objRef;

    DISPID putId = DISPID_PROPERTYPUT;
    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    dp.cArgs = 1;
    dp.cNamedArgs = 1;
    dp.rgvarg = &varObj;
    dp.rgdispidNamedArgs = &putId;

    EXCEPINFO excep;
    memset(&excep, 0, sizeof(excep));
    UINT argErr = 0;

    // 优先尝试PROPERTYPUTREF, 失败则回退PROPERTYPUT
    HRESULT hr = vb6_ComInvoke(pDisp, dispid, DISPATCH_PROPERTYPUTREF, &dp, NULL, &excep, &argErr);
    if (FAILED(hr)) {
        hr = vb6_ComInvoke(pDisp, dispid, DISPATCH_PROPERTYPUT, &dp, NULL, &excep, &argErr);
    }

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComSetRef");
    }
}
