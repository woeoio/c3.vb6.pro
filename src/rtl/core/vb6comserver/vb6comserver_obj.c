// vb6comserver_obj.c - vb6comserver 模块拆分: IDispatch 包装层 (vb6_ComObject)
// 由 vb6comserver.c 按 COM 接口家族拆分而来 (纯搬移, 零行为改动)

#include <string.h>
#include <stdlib.h>
#define COBJMACROS  /* Enable C COM macros (ITypeLib_Release etc.) */
#include "vb6comserver.h"
#include "vb6comserver_internal.h"
#include "vb6rtl_variant.h"   /* 账 #305 §B144: 两头的槽<->VARIANT 换算出口 */
/* 账 #305 §B185: 服务端进门那趟要自带一枚 VB 错误帧。下面这批跨族 extern 按本仓既有做法就地声明
 * (vb6com 系列不 include vb6rtl.h —— VARIANT 冲突), 定义全在 vb6rtl.c。 */
#include <setjmp.h>
extern jmp_buf* vb6_error_jmp_ptr;
extern int32_t vb6_error_jmp_set;
extern int32_t vb6_err_jmp_active;
extern int32_t vb6_err_resume_next;
extern void vb6_SaveErrState(void);
extern void vb6_RestoreErrState(void);
extern int32_t vb6_ErrNumber(void);
extern BSTR vb6_ErrDescription(void);   /* 借来的指针, 交出去前要自己拷一份 (见 vb6_ComFailExcep) */
extern HRESULT vb6_ComFailExcep(EXCEPINFO* pexcep, int32_t errNum,
                                const wchar_t* desc, const wchar_t* fallback);
extern HRESULT vb6_ComVbHresult(int32_t errNum);  /* §B186: 类工厂那一趟的交回形状 (定义在 vb6com.c) */

/* [PROBE189] 临时探针: 包装器引用计数轨迹 (C3_COM_RC_TRACE=1 时输出到 stderr) */
#include <stdio.h>
static int probe189_on(void) {
    static int v = -1;
    if (v < 0) v = (getenv("C3_COM_RC_TRACE") != NULL);
    return v;
}
#define P189(...) do { if (probe189_on()) { fprintf(stderr, __VA_ARGS__); fflush(stderr); } } while (0)
/* [/PROBE189] */

#ifndef CONNECT_E_NOCONNECTION
#define CONNECT_E_NOCONNECTION 0x80040200
#endif

#ifndef GUIDKIND_DEFAULT_SOURCE_DISP_IID
#define GUIDKIND_DEFAULT_SOURCE_DISP_IID 1
#endif

// ============================================================
// vb6_ComObject - IDispatch wrapper implementation
// ============================================================

// --- IUnknown ---

static HRESULT STDMETHODCALLTYPE ComObj_QueryInterface(vb6_ComObject* self, REFIID riid, void** ppv) {
    int i;
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown_) || IsEqualIID(riid, &IID_IDispatch_)) {
        *ppv = self;
        self->vtable->AddRef(self);
        return S_OK;
    }
    // P12.1: Check Implements interface IIDs
    if (self->desc && self->desc->ifaceCount > 0 && self->desc->ifaceIids) {
        for (i = 0; i < self->desc->ifaceCount; i++) {
            if (IsEqualIID(riid, self->desc->ifaceIids[i])) {
                // ai/022 B16: 新式接口 (Interface 块) 交回**薄指针** —— 首字段是
                // vb6_ivtbl_<I>*, 槽 3 起是契约成员, 与类型库 TKIND_INTERFACE 那一条
                // (cFuncs / oVft / CC_STDCALL) 同形. 这正是 B15 起对外广告的那一档,
                // 所以"广告 == 应答"在接口这一档上成立 (D60-4 的"胖应答瘦"到此收口).
                // 生命周期: 薄指针的引用记在实例自己的 __refcount 上 (B05), 包装器
                // 最后一次 Release 会把底座引用交还 (见 ComObj_Release) → QI 之后立刻
                // Release 包装器 (IClassFactory::CreateInstance 的规范姿势) 也安全.
                if (self->desc->ifaceThinPtr && self->vb6Instance) {
                    void* thin = self->desc->ifaceThinPtr(self->vb6Instance,
                                                           self->desc->ifaceIids[i]);
                    if (thin) {
                        vb6_ivtbl_prefix* vt = *(vb6_ivtbl_prefix**)thin;
                        vt->AddRef(thin);
                        *ppv = thin;
                        return S_OK;
                    }
                    /* 本类没实现这个新式接口 (provider 认 IID): 落到胖路 */
                }
                *ppv = self;  // dispinterface: same IDispatch pointer
                self->vtable->AddRef(self);
                return S_OK;
            }
        }
    }
    // Default dispinterface IID
    if (self->desc && self->desc->defaultIfaceIid && IsEqualIID(riid, self->desc->defaultIfaceIid)) {
        *ppv = self;  // dispinterface = same IDispatch pointer
        self->vtable->AddRef(self);
        return S_OK;
    }
    // IConnectionPointContainer (only if coclass has events)
    if (self->desc && self->desc->sourceIfaceIid && IsEqualIID(riid, &IID_IConnectionPointContainer_)) {
        if (!self->cpc) {
            self->cpc = vb6_CPC_Create(self);
        }
        if (self->cpc) {
            *ppv = self->cpc;
            self->cpc->vtable->AddRef(self->cpc);
            return S_OK;
        }
    }
    // IProvideClassInfo2
    if (self->desc && IsEqualIID(riid, &IID_IProvideClassInfo2_)) {
        if (!self->pci) {
            self->pci = vb6_PCI_Create(self);
        }
        if (self->pci) {
            *ppv = self->pci;
            self->pci->vtable->AddRef(self->pci);
            return S_OK;
        }
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE ComObj_AddRef(vb6_ComObject* self) {
    ULONG count = InterlockedIncrement(&self->refCount);
    P189("[189] ADDREF W=%p cnt=%u\n", (void*)self, (unsigned)count);
    InterlockedIncrement(&g_vb6_cRef);
    return count;
}

/* 账 #305 §B188 (第九十一刀): 下面那条 [C3_COM] 诊断原先**每出厂一枚实例打一行**。
 * 它是一指, 不是流水 —— 判据 (tests\tools\leak_probe.c) 一轮 5000 次 create/release
 * 就是 5000 行, 本机 20 万轮实测 19.3 MB stderr 把现场盖掉。口径照 §B181 的
 * vb6_CrashTraceClaim: **同一枚类只记第一次**, 拿不准的一律放行 (指针没有 desc 的空值,
 * 表满也照打) —— 漏掉一条真诊断比多打十条贵。表是进程内的定长小表, 不回收。 */
#define VB6_TERM_ERR_LOG_MAX 4
static const void* volatile vb6_termErrLogged[VB6_TERM_ERR_LOG_MAX];

static int vb6_TermErrClaim(const void* desc) {
    int i;
    if (!desc) return 1;
    for (i = 0; i < VB6_TERM_ERR_LOG_MAX; i++) {
        if (vb6_termErrLogged[i] == desc) return 0;
        if (vb6_termErrLogged[i] == NULL &&
            InterlockedCompareExchangePointer((void* volatile*)&vb6_termErrLogged[i],
                                              (void*)desc, NULL) == NULL) {
            return 1;
        }
    }
    return 1;
}

static ULONG STDMETHODCALLTYPE ComObj_Release(vb6_ComObject* self) {
    ULONG count = InterlockedDecrement(&self->refCount);
    P189("[189] RELEASE W=%p cnt=%u owns=%d inst=%p\n", (void*)self, (unsigned)count,
         self->ownsInstance, (void*)self->vb6Instance);
    InterlockedDecrement(&g_vb6_cRef);
    if (count == 0) {
        // Fix 188 + ai/022 B16 的并集: 是否销毁实例先看**拥有关系** —— 只有拥有型包装
        // (CoCreateInstance / Set x = New cY) 才动实例; 借用型 (Public 对象字段 getter /
        // 方法返回的工程类实例) 的生命周期归宿主, 只清实例上的 __comObj 回填。
        // 拥有型里, 实现了新式接口的类走 claim 那条: 包装器退掉自己的"底座引用",
        // 实例若有薄引用在世则不销毁 (由最后一个薄引用的 Release 收尾), 否则当场销毁;
        // 没有这道 claim 的类 (存量/无接口) 结果与原来直接 destroyFunc 完全一致。
        //
        // 账 #305 §B188 (第八十刀): 上面这两条腿都会**踏进 VB 代码** (Class_Terminate), 以前
        // 是本文件里三站中唯一没接错误边界的一站 —— 体里没人接住的错走到 vb6_RaiseError,
        // 那里没有帧 ⇒ 打一行 stderr 再 ExitProcess(错误号), 把**正在放手的客户端进程**打死
        // (x64 实测: 一枚 terminate 里只写 Err.Raise 234 的类 ⇒ 探针 exit=234, 后面那枚干净
        // 对象的 CREATE=2 与 DONE 全都没打出来)。
        // 这一趟与 §B185/§B186 那两趟有一处实质差别: **没有一条通道能把号带回去** —— Release
        // 交回的是 ULONG 引用计数, 也没有 EXCEPINFO, 改返回值 = 破坏 COM 契约。所以这里只做
        // 底线: 自带帧 + 失败往 stderr 记一条带号诊断 + 照常把包装器拆干净、照常交回引用计数。
        // 号此刻既不进 Err 也不进客户端 (§B188 出路 (b) 是另一格: 那要先答「谁有权改 Err 的
        // 清除时机」)。帧的登记顺序照 cgen 的 On Error GoTo 协议 ⇒ 体里自己写了 On Error 的
        // 照常由体接住, 走到这里的是真没人管的那一类。
        if (self->ownsInstance) {
            int32_t termErrNum = 0;
            {
                jmp_buf termJmp;
                vb6_SaveErrState();
                if (setjmp(termJmp) != 0) {
                    termErrNum = vb6_ErrNumber();
                } else {
                    vb6_error_jmp_ptr = &termJmp;
                    vb6_error_jmp_set = 1;
                    vb6_err_jmp_active = 1;
                    vb6_err_resume_next = 0;
                    if (self->desc && self->desc->instanceClaimRelease && self->vb6Instance) {
                        self->desc->instanceClaimRelease(self->vb6Instance);
                    } else if (self->desc && self->desc->destroyFunc && self->vb6Instance) {
                        self->desc->destroyFunc(self->vb6Instance);
                    }
                }
                vb6_RestoreErrState();
            }
            if (termErrNum && vb6_TermErrClaim(self->desc)) {
                fprintf(stderr, "[C3_COM] Class_Terminate raised error %d in \"%s\"; the client was kept alive (ledger 305 B188)\n",
                        (int)termErrNum,
                        (self->desc && self->desc->classVariable) ? self->desc->classVariable : "?");
                fflush(stderr);
            }
        } else if (self->vb6Instance) {
            *(void**)self->vb6Instance = NULL;
        }
        self->vb6Instance = NULL;
        // Release PCI
        if (self->pci) {
            self->pci->vtable->Release(self->pci);
            self->pci = NULL;
        }
        // Release CPC
        if (self->cpc) {
            self->cpc->vtable->Release(self->cpc);
            self->cpc = NULL;
        }
        CoTaskMemFree(self);
    }
    return count;
}

// --- IDispatch ---

static HRESULT STDMETHODCALLTYPE ComObj_GetTypeInfoCount(vb6_ComObject* self, UINT* pctinfo) {
    if (!pctinfo) return E_POINTER;
    *pctinfo = 1;  // Provide TypeLib info (needed for early binding)
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE ComObj_GetTypeInfo(vb6_ComObject* self, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) {
    if (!ppTInfo) return E_POINTER;
    if (iTInfo != 0) return DISP_E_BADINDEX;
    *ppTInfo = NULL;
    if (!self->desc) return E_FAIL;
    // Get DLL path, load TypeLib from embedded resource
    wchar_t dllPath[MAX_PATH];
    HMODULE hMod = NULL;
    if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, (LPCWSTR)&g_vb6_cRef, &hMod)) {
        return E_FAIL;
    }
    GetModuleFileNameW(hMod, dllPath, MAX_PATH);
    ITypeLib* pTypeLib = NULL;
    HRESULT hr = LoadTypeLib(dllPath, &pTypeLib);
    if (FAILED(hr) || !pTypeLib) {
        return E_FAIL;
    }
    // Find coclass ITypeInfo by CLSID
    CLSID clsid;
    hr = vb6_CLSIDFromStrA(self->desc->clsidStr, &clsid);
    if (FAILED(hr)) {
        ITypeLib_Release(pTypeLib);
        return E_FAIL;
    }
    hr = ITypeLib_GetTypeInfoOfGuid(pTypeLib, &clsid, ppTInfo);
    ITypeLib_Release(pTypeLib);
    return hr;
}

static HRESULT STDMETHODCALLTYPE ComObj_GetIDsOfNames(vb6_ComObject* self, REFIID riid, 
    LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) 
{
    if (!self->desc || !self->desc->methods) return DISP_E_UNKNOWNNAME;
    
    for (UINT n = 0; n < cNames; n++) {
        rgDispId[n] = DISPID_UNKNOWN;
        for (int i = 0; i < self->desc->methodCount; i++) {
            if (wcscmp(rgszNames[n], self->desc->methods[i].name) == 0) {
                rgDispId[n] = self->desc->methods[i].dispid;
                break;
            }
        }
    }
    
    return (cNames > 0 && rgDispId[0] != DISPID_UNKNOWN) ? S_OK : DISP_E_UNKNOWNNAME;
}

static HRESULT STDMETHODCALLTYPE ComObj_Invoke(vb6_ComObject* self, DISPID dispIdMember, 
    REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, 
    VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
    if (!self->desc || !self->desc->methods) return DISP_E_MEMBERNOTFOUND;
    
    // Find method matching dispid AND invkind (same dispid may have Get/Let variants)
    const vb6_DispMethodDesc* method = NULL;
    for (int i = 0; i < self->desc->methodCount; i++) {
        if (self->desc->methods[i].dispid == dispIdMember) {
            int ik = self->desc->methods[i].invkind;
            int ok = 0;
            if (ik == 1 && (wFlags & DISPATCH_METHOD)) ok = 1;
            else if (ik == 2 && (wFlags & DISPATCH_PROPERTYGET)) ok = 1;
            else if ((ik == 4 || ik == 8) && (wFlags & DISPATCH_PROPERTYPUT)) ok = 1;
            if (ok) { method = &self->desc->methods[i]; break; }
        }
    }
    // Fallback: if no invkind match, try first dispid match
    if (!method) {
        for (int i = 0; i < self->desc->methodCount; i++) {
            if (self->desc->methods[i].dispid == dispIdMember) {
                method = &self->desc->methods[i];
                break;
            }
        }
    }
    if (!method) return DISP_E_MEMBERNOTFOUND;
    P189("[189] INVOKE W=%p cls=%s mem=%ls argc=%d\n", (void*)self,
         (self->desc && self->desc->classVariable) ? self->desc->classVariable : "?",
         method->name ? method->name : L"?", (int)(pDispParams ? pDispParams->cArgs : 0));

    // Collect arguments
    // Note: 调用方(VBScript/Office/另一枚工程类)递来的 vt 与形参的声明档**不一定相同**，
    //   而下面生成的读臂是按声明档取联合体成员的 —— 这一趟负责把每一格先换算到那一档，
    //   见 §B219 那段与 vb6_ComArgNormalize 的声明处。
    // invokeFunc 的实参/返回槽在适配器里按 vb6_VARIANT (x86 24B: vt 4B +
    // pad 4B + union 16B) 布局访问; 前 16 字节与 OLE VARIANT 内容一致。
    typedef struct { uint64_t v[3]; } vb6_VarSlot;
    int argc = pDispParams ? (int)pDispParams->cArgs : 0;
    void** args = NULL;
    void* coercedArgs = NULL;   /* 槽阵列, 每槽 vb6_VarSlot */
    VARIANT* scratch = NULL;    /* 账 #305 §B219: 换算产物住这里, 服务端在收尾那一臂回收 */
    int argErrSlot = -1;        /* 账 #305 §B219: 换算不动的那一枚(声明序下标), -1 = 没有 */

    if (argc > 0) {
        args = (void**)CoTaskMemAlloc(argc * sizeof(void*));
        coercedArgs = (void*)CoTaskMemAlloc(argc * sizeof(vb6_VarSlot));
        scratch = (VARIANT*)CoTaskMemAlloc(argc * sizeof(VARIANT));
        if (!args || !coercedArgs || !scratch) {
            if (args) CoTaskMemFree(args);
            if (coercedArgs) CoTaskMemFree(coercedArgs);
            if (scratch) CoTaskMemFree(scratch);
            return E_OUTOFMEMORY;
        }
        // 适配器按 vb6_VARIANT (24B) 布局访问实参与返回槽, 而 COM 边界的
        // OLE VARIANT 是 16B。每个实参先整复制进 24B 槽, 再按声明档换算。
        // (ai/022 B17 那条"未初始化就被 VariantClear"的堆损坏在这里由下面的
        //  `memset(slot, 0, sizeof(*slot))` 一并解决 —— 旧写法按 VARIANT 尺寸清零,
        //  在这套 24B 槽设计下 x86 会算少, 故合并时只保留这一份。)
        for (int i = 0; i < argc; i++) {
            VARIANT* src = &pDispParams->rgvarg[argc - 1 - i];
            vb6_VarSlot* slot = &((vb6_VarSlot*)coercedArgs)[i];
            /* 账 #305 §B144: 这里原来是 memset + memcpy(槽, src, sizeof(VARIANT))，
             * 主张「槽的头 16 字节与 OLE VARIANT 一致」—— 标量那几档成立，VT_DECIMAL 不
             * 成立(对端把 DECIMAL 放在结构体头，我们 x64 在 offset 8、x86 是指针)。
             * 布局那一问自此只住 RTL 的出口，服务端不再自己数偏移。 */
            vb6_VariantBorrowIntoSlot(src, slot);
            args[i] = slot;
        }
        /* 账 #305 §B219(第二百二十刀): 「往里读」那一族在此多出来的一问 —— 发码那 16 处读臂
         * 按**形参声明档**取联合体成员、从不看标签，于是异档实参一律读错一格。靶工程与合成
         * 夹具两台实测(20 发矩阵在 022 第二百二十刀的账里)：
         *   String 档 ← VT_I4 40 / VT_R8 / VT_BOOL ⇒ 把标量当 BSTR 解 ⇒ 原生 AV，**打死客户端**；
         *   Long 档 ← VT_BSTR "12" ⇒ 答指针低 32 位(实测 1833297976)、← VT_R8 12.75 ⇒ 答 0，静默错；
         *   Long 档 ← VT_BOOL / VT_EMPTY ⇒ 侥幸对(靠本处原先那半张手抄表恰好折了这两档)。
         * 这里原来是那半张表(只折 VT_I2/I1/UI1/UI2/BOOL/EMPTY→I4、R4→R8，异档名单之外的照旧
         * 裸取)，与发码那一半合起来才是「两半各答一半」。现在两半都并到唯一出口
         * vb6_ComArgNormalize 上：档位答案由表里的 paramVts 给(cgen 那一处权威 cDispatchArgVt 答)，
         * 换算与失败那一问住 RTL。换算不动交回 DISP_E_TYPEMISMATCH —— 客户端 vb6_ComCheckError
         * 已经把它折成 VB 13(vb6com.c:256)，与 VB6 本人的「Type mismatch」同号。
         * 不许图省事就地换算槽：槽借用了调用方的 BSTR/Dispatch 指针，VariantChangeType 会替
         * 它们 SysFreeString/Release ⇒ 对端随后 clear 就是双份。 */
        for (int i = 0; i < argc; i++) VariantInit(&scratch[i]);
        if (method->paramVts) {
            for (int i = 0; i < argc && i < (int)method->paramVtCount; i++) {
                VARIANT* src = &pDispParams->rgvarg[argc - 1 - i];
                vb6_VarSlot* slot = &((vb6_VarSlot*)coercedArgs)[i];
                if (FAILED(vb6_ComArgNormalize(src, method->paramVts[i], slot, &scratch[i]))) {
                    argErrSlot = i;
                    break;
                }
            }
        }
    }

    // Call VB6 method
    /* 账 #305 §B185: 这一句是工程类服务端走进 VB 代码的**唯一**入口 (方法、属性、Public 字段
     * 三类都从它走), 而它过去没有任何错误边界: VB 体里没人接住的运行期错误一路走到
     * vb6_RaiseError, 那里没有帧 ⇒ 打一行 stderr 再 exit(错误号), 把**客户端**进程打死
     * (两台实测: 写越界的 Integer 字段 RC=6, 体里一句 Err.Raise 5 RC=5)。
     * 现在这一趟自带一枚帧: 帧只登记在**外层没人装**的时候才轮得到接住 —— 顺序照 cgen 的
     * On Error GoTo 协议 (先 setjmp, 成功后再登记 ptr/set/active 并清 resume_next),
     * 于是体里自己写了 On Error 的那些照常由体接住, 交回 S_OK 与空 EXCEPINFO。
     * 失败时按内置 Collection 早就在用的那一形状交回 (vb6_ComFailExcep), 客户端的
     * vb6_ComCheckError 把它折回成一条 VB 错误 ⇒ 客户的 On Error 拿得到错误号。
     * 注意: 不许在这里 longjmp 之后直接 return —— 下面那条收尾要回收 args/coercedArgs
     * 与 §B162 那枚数组那一臂, 跳过它就是漏窗。 §B219 同规: 换算不动那一发也是先收尾再交回。 */
    int32_t srvErrNum = 0;
    if (argErrSlot < 0 && method->invokeFunc) {
        vb6_VarSlot tmpRet;
        jmp_buf srvJmp;
        memset(&tmpRet, 0, sizeof(tmpRet));
        vb6_SaveErrState();
        if (setjmp(srvJmp) != 0) {
            srvErrNum = vb6_ErrNumber();
        } else {
            vb6_error_jmp_ptr = &srvJmp;
            vb6_error_jmp_set = 1;
            vb6_err_jmp_active = 1;
            vb6_err_resume_next = 0;
            method->invokeFunc(self->vb6Instance, args, argc, &tmpRet);
        }
        vb6_RestoreErrState();
        if (srvErrNum == 0 && pVarResult)
            /* 账 #305 §B144: 同上 —— 出站也改问唯一出口(Decimal 那一档不再错位的
             * 同时，返回的串/对象现在带权交出去，对端 VariantClear 不会释放我们的引用)。 */
            vb6_VariantPackCopyInto(&tmpRet, pVarResult);
    }

    /* 账 #305 §B162 (第五十刀): 数组那一臂是个例外 —— 递进来的真 SAFEARRAY 被换成了**我们自己**
     * 那枚描述符, 槽阵列一 free 它就没人管了。其余各档仍是调用方的资源, 一律不动。 */
    /* 账 #305 §B222 (第二百三十刀): 数组形参的**出站回写** —— 被调方 ReDim 过之后新载体只活在
     * 这一趟的槽里, 不写回去客户端拿到的还是调用前那枚(§B162「回写刻意不做」的另一半)。
     * 只认 VT_BYREF|VT_ARRAY 那一形(数组在 VB6 里只能按引用传)。顺序不能挪: 本趟只读槽、
     * 回写成功后把槽的 vt 补成数组档, 下面那句 SlotRelease 才认得出并回收我们的描述符。 */
    if (argErrSlot < 0 && method->invokeFunc) {
        for (int i = 0; i < argc; i++) {
            vb6_ArraySlotWriteBack(&pDispParams->rgvarg[argc - 1 - i],
                                   &((vb6_VarSlot*)coercedArgs)[i]);
            /* 账 #305 §B223 (第二百三十三刀): ByRef 标量那一族同一时机回写(数值/布尔);
             * 两枚各认各的形(BYREF|ARRAY vs BYREF|标量), 互不重叠。 */
            vb6_ByRefScalarWriteBack(&pDispParams->rgvarg[argc - 1 - i],
                                     &((vb6_VarSlot*)coercedArgs)[i]);
        }
    }
    for (int i = 0; i < argc; i++) {
        vb6_VariantSlotRelease(&((vb6_VarSlot*)coercedArgs)[i]);
    }
    // 槽内是调用方 VARIANT 的副本, 资源仍归调用方, 不得 VariantClear
    /* 账 #305 §B219: 换算过的那几格例外 —— 槽里借的是 scratch 里的**我们自己的**副本，
     * 所以回收顺序是「先让被调方用完，再 clear scratch」；scratch 一定先于槽阵列 free。 */
    if (scratch) {
        for (int i = 0; i < argc; i++) VariantClear(&scratch[i]);
        CoTaskMemFree(scratch);
    }
    if (coercedArgs) CoTaskMemFree(coercedArgs);
    if (args) CoTaskMemFree(args);
    /* 账 #305 §B219: 换算不动那一枚 —— pArgErr 交的是 **rgvarg 的下标**(逆序那一头), 不是声明序。 */
    if (argErrSlot >= 0) {
        if (puArgErr) *puArgErr = (UINT)(argc - 1 - argErrSlot);
        return DISP_E_TYPEMISMATCH;
    }
    /* 账 #305 §B185: 收尾(上面那批回收)之后才交回错误 —— 顺序是硬要求。 */
    if (srvErrNum != 0)
        return vb6_ComFailExcep(pExcepInfo, srvErrNum, vb6_ErrDescription(), L"COM server error");
    return S_OK;
}

// IDispatch vtable instance
static const vb6_IDispatchVtable g_ComObjectVtable = {
    ComObj_QueryInterface,
    ComObj_AddRef,
    ComObj_Release,
    ComObj_GetTypeInfoCount,
    ComObj_GetTypeInfo,
    ComObj_GetIDsOfNames,
    ComObj_Invoke,
};

// Create VB6 COM object
vb6_ComObject* vb6_ComObject_Create(const vb6_CoClassDesc* desc, HRESULT* pFailHr) {
    if (!desc || !desc->factoryFunc) {
        if (pFailHr) *pFailHr = CLASS_E_CLASSNOTAVAILABLE;
        return NULL;
    }
    
    vb6_ComObject* obj = (vb6_ComObject*)CoTaskMemAlloc(sizeof(vb6_ComObject));
    if (!obj) {
        if (pFailHr) *pFailHr = E_OUTOFMEMORY;
        return NULL;
    }
    
    obj->vtable = &g_ComObjectVtable;
    obj->refCount = 1;
    obj->desc = desc;
    obj->vb6Instance = NULL;
    /* 账 #305 §B186 (第七十八刀): 服务端进门的**第二站** —— 下面这一步走进 VB 代码, 而且走的是
     * Class_Initialize (§B185 那一趟是方法/属性/字段那三类)。它过去同样没有错误边界: 体里没人接住
     * 的运行期错误一路走到 vb6_RaiseError, 那里没有帧 ⇒ 打一行 stderr 再 exit(错误号), 把**客户端**
     * 进程打死 (两台实测一枚只写 Err.Raise 7 的类 = RC=7, CREATE 那一行根本没打出来)。
     * 帧的登记顺序照 cgen 的 On Error GoTo 协议 (先 setjmp, 成功后再登记 ptr/set/active 并清
     * resume_next), 于是体里自己写了 On Error 的那些照常由体接住、照旧构造成功。
     * 交回形状只能编进 HRESULT: IClassFactory::CreateInstance 的签名里没有 EXCEPINFO 通道,
     * 所以「造」那一问问 vb6_ComVbHresult (vb6com.c, §B185 已在用同一枚答案), 不在这里抄形状。
     * 回收只到 obj 为止: 类实例指针是 factoryFunc 的返回值, 抛错时永远拿不到, 要回收它得动类工厂
     * 的发码 —— 那是另一问, 已另立 §B187 登记读数。 */
    int32_t srvErrNum = 0;
    {
        jmp_buf srvJmpC;
        vb6_SaveErrState();
        if (setjmp(srvJmpC) != 0) {
            srvErrNum = vb6_ErrNumber();
        } else {
            vb6_error_jmp_ptr = &srvJmpC;
            vb6_error_jmp_set = 1;
            vb6_err_jmp_active = 1;
            vb6_err_resume_next = 0;
            obj->vb6Instance = desc->factoryFunc();  // Call vb6_cls_<Name>_New()
        }
        vb6_RestoreErrState();
    }
    if (srvErrNum) {
        CoTaskMemFree(obj);
        if (pFailHr) *pFailHr = vb6_ComVbHresult(srvErrNum);
        return NULL;
    }
    P189("[189] CREATE W=%p inst=%p cls=%s\n", (void*)obj, (void*)obj->vb6Instance,
         desc->classVariable ? desc->classVariable : "?");
    obj->ownsInstance = 1;  // Fix 188: 类工厂创建的实例归客户端所有
    obj->cpc = NULL;  // Lazy init CPC
    obj->pci = NULL;  // Lazy init PCI
    // Set back-pointer for event support (first field of VB6 class struct = __comObj)
    if (obj->vb6Instance) {
        void** ppComObj = (void**)obj->vb6Instance;
        *ppComObj = obj;
    }
    
    if (!obj->vb6Instance) {
        CoTaskMemFree(obj);
        if (pFailHr) *pFailHr = E_OUTOFMEMORY;
        return NULL;
    }
    
    InterlockedIncrement(&g_vb6_cRef);
    return obj;
}

// ============================================================
// Fix 099: 包装已存在的 VB6 类实例 (Public 对象字段的 COM 暴露)
// ============================================================
// ActiveX DLL 里 `Public Router As cHttpServerRouter` 这类字段的实例, 是在类
// 内部 (Class_Initialize 的 `Set Router = New ...`) 创建的裸结构体指针, 其
// __comObj 仍为 NULL. COM 客户端读该字段 (Property Get) 时必须拿到一个能继续
// 调用方法的 IDispatch —— 直接返回裸指针冒充 VT_DISPATCH, 客户端首次 Invoke
// 就会把结构体当 IDispatch 解引用 vtable → 0xC0000005.
//
// 复用规则: 实例首个字段 __comObj 已指向包装对象时直接 AddRef 复用 (同一实例
// 每次读出返回同一 IDispatch, 引用计数正确), 否则新建包装并回填 __comObj.
// 前置条件: 实例所在类的结构体首字段必须是 __comObj —— 自 ExeComBridge 01 起**所有**类模块
// 的结构体都有这一格 (工厂里那句 me->__comObj = NULL 是无条件发的), 所以本函数不只服务 DLL
// 侧的字段 getter; 但 EXE 里今天既不装箱也没人查登记表 (§B187), 真走到它的仍只有 ActiveX 那一路。
vb6_ComObject* vb6_ComObject_FromInstance(const vb6_CoClassDesc* desc, void* instance) {
    void** ppComObj;
    vb6_ComObject* existing;
    vb6_ComObject* obj;
    if (!desc || !instance) return NULL;

    ppComObj = (void**)instance;
    existing = (vb6_ComObject*)*ppComObj;
    if (existing) {
        existing->vtable->AddRef(existing);
        return existing;
    }

    obj = (vb6_ComObject*)CoTaskMemAlloc(sizeof(vb6_ComObject));
    if (!obj) return NULL;

    obj->vtable = &g_ComObjectVtable;
    obj->refCount = 1;
    obj->desc = desc;
    obj->vb6Instance = instance;
    obj->ownsInstance = 1;  /* Fix 188: 调用方把新实例的所有权交给包装器 */
    P189("[189] FROM-INST W=%p inst=%p cls=%s\n", (void*)obj, (void*)instance,
         desc->classVariable ? desc->classVariable : "?");
    obj->cpc = NULL;
    obj->pci = NULL;
    *ppComObj = obj;  /* 回填 __comObj, 后续复用 */

    InterlockedIncrement(&g_vb6_cRef);
    return obj;
}

// Fix 188: 包装"宿主已拥有"的实例 —— 见 vb6comserver.h 声明. 释放到 0 时只回收
// 包装器并清空 __comObj 回填, 不调用 destroyFunc.
vb6_ComObject* vb6_ComObject_FromBorrowedInstance(const vb6_CoClassDesc* desc, void* instance) {
    void** ppComObj;
    vb6_ComObject* existing;
    vb6_ComObject* obj;
    if (!desc || !instance) return NULL;

    ppComObj = (void**)instance;
    existing = (vb6_ComObject*)*ppComObj;
    if (existing) {
        existing->vtable->AddRef(existing);
        return existing;
    }

    obj = (vb6_ComObject*)CoTaskMemAlloc(sizeof(vb6_ComObject));
    if (!obj) return NULL;

    obj->vtable = &g_ComObjectVtable;
    obj->refCount = 1;
    obj->desc = desc;
    obj->vb6Instance = instance;
    obj->ownsInstance = 0;
    P189("[189] FROM-BORROW W=%p inst=%p cls=%s\n", (void*)obj, (void*)instance,
         desc->classVariable ? desc->classVariable : "?");
    obj->cpc = NULL;
    obj->pci = NULL;
    *ppComObj = obj;

    InterlockedIncrement(&g_vb6_cRef);
    return obj;
}

// Fix 099: 按类变量名 (VB6 模块名, 如 "cHttpServerRouter") 查 coclass 描述.
// 供 dll_entry.c 的 Public 对象字段 getter 使用 (包装实例需要 desc).
// 只命中 coclass 表内的类 (cgen 仅收 MultiUse/SingleUse); 未命中返回 NULL,
// 此时 getter 返回 NULL dispatch — 客户端拿到 Nothing, 不会崩.
const vb6_CoClassDesc* vb6_FindCoClassDesc(const char* classVariable) {
    int i;
    if (!classVariable) return NULL;
    for (i = 0; i < g_vb6_coclassCount; i++) {
        if (g_vb6_coclasses[i].classVariable &&
            _stricmp(g_vb6_coclasses[i].classVariable, classVariable) == 0) {
            return &g_vb6_coclasses[i];
        }
    }
    return NULL;
}

// Fix <vbeclipse> rev30: **工程类实例 → coclass 描述**的反查表。
//
// 为什么需要 (play78 --arch x86, 探针链逐级定位, 全部是直接铁证):
//   `ucFolder.ContainsView(ViewId)` 生成成
//     vb6_ComCall(vb6_ComGetObjectProp(ViewTabs, L"Tabs"), L"Contains", {ViewId}, 1)
//   探针实测: `Tabs` 对象**拿得到** (rev25 的对象返回型桥在工作), `Contains`
//   也**返回了非空指针**, 但那个 Variant 的 **vt=0 (VT_EMPTY)** —— 值没被填。
//   根因: `tabsObj` 是**裸 `vb6_cls_List*`** (原生工程类实例, 首字段 `__comObj` 恒 NULL
//   ⇒ 没有真 vtable)。`vb6_ComCall` 只有两条路: 宿主对象 (HWND/包装器/UC 实例) 或
//   真 `IDispatch`; `vb6_ComIsDispatchable` 读首字段当 vtbl → 判否 →
//   `GetIDsOfNames` 返回 DISPID_UNKNOWN ⇒ **返回值永远是 VT_EMPTY**。
//   ⇒ `Contains` 恒 0 ⇒ `ucPerspective.ShowView` 遍历 5 个 folder 全部 miss
//   (探针: `ContainsView=0` ×5) ⇒ **没有任何 folder 被激活 ⇒ 停靠面板全空**。
//
// 为什么不能靠包装器: `vb6_ComObject_FromInstance` 能把实例包成真 IDispatch, 但
// **调用点拿到的就是裸指针**, 包一次就多一次引用计数与一次释放责任, 而这里只需要
// "按方法名找到 invokeFunc 并调它"。所以直接查原生方法表。
//
// 为什么用注册表而不是遍历 coclass 表猜: 实例指针里**没有**任何 desc 痕迹
// (`__comObj` 是 back-ptr 且为 NULL), 无法反推是哪个类。cgen 在类工厂里登记一次
// 是唯一可靠办法 —— 而且工厂本来就是每个实例必经之处, 零额外成本。
//
// 生命周期 (账 #481 = 第一百八十二刀订正): 实例销毁时**必须**注销 (见 vb6_UnregisterProjectClassInstance,
// 调用点 = 每个类 _Destroy 里 vb6_Free 前一行, 与登记腿同一道 isDll_ 闸)。从前的说法是
// 「表是弱映射, 只按指针查, 不注销也没事」—— 两处站不住:
//   ① 实例释放后地址留在表里, 地址被别的分配复用、而新主人不是工程类实例时, 按旧地址查会答旧类;
//   ② 计数只增不减 ⇒ 一个进程里累计建过 VB6_PROJCLASS_MAX 枚实例之后, 新实例一个都登记不进去。
// 注销只把那一格置空、不动计数 (计数 = 扫描上界的高水位), 登记腿优先复用空出来的格。
//
// 2026-10-15 (账 #305 §B187) 两条订正, 都按读码/真跑核过:
//   ① 这里原来写的是「命中后还要用 vb6_ProjClassVerify 复核 (见 vb6com_invoke.c 的分派
//      处)」—— 那一枚复核函数与那个分派处在整棵树与全部历史里都不存在 (git log -S 对两个
//      名字都只命中引入本表那一笔), 所以它是一个从来没接上的打算, 不是现有机制。今天
//      查表的读者是 0 枚: vb6_FindProjectClassDesc / ...Raw 全仓无调用点 ⇒ 本表只有登记
//      这一腿在跑。2026-10-15 口径已拍 (§B187 第二格): **按工程类型分档** —— ActiveX 那一档
//      (Type=DLL / Type=Control, 后端读 isDll_) 走 (b) = 接读者 + 配注销腿, 其余走 (a) =
//      连登记都不发。订正一条旧解释: 先前写的「EXE 里没有 CoClass 块 ⇒ vb6_FindCoClassDesc
//      恒 NULL ⇒ 这类实例压根没进表」是**错的** —— EXE 也链一张完整 coclass 表 (com_entry.c,
//      ExeComBridge 01), 收集判据是类的 Instancing 而非有没有手写 CoClass 块, 那张表在 EXE 里
//      天天被写。真断点在两条腿都没接: Set v = <工程内对象> 发成**裸指针赋值** (不装箱、不查
//      表), 晚绑定那站 vb6_ComIsDispatchable 判否 ⇒ DISPID_UNKNOWN ⇒ 恒 VT_EMPTY (实测
//      .build/k79_late/ 与 .build/k81_run.bat, 两台逐行相同: v.Twice(21) 空答、Err=0)。
//   ② 「扫到哪儿」这一问以前有两个答案: 登记腿那圈只按 g_vb6_projClassCount 走, 查表腿
//      那圈另外夹住 VB6_PROJCLASS_MAX; 而计数是在夹界检查之前自增、放弃登记那一路不回退
//      ⇒ 计数无上界, 从第 514 枚实例起登记腿就在读数组外面的相邻静态数据 (巧合命中时还会
//      往外面写)。现在两圈都只问 vb6_ProjClassScanLimit() 这一处。
#define VB6_PROJCLASS_MAX 512
static struct { void* inst; const vb6_CoClassDesc* desc; } g_vb6_projClass[VB6_PROJCLASS_MAX];
static volatile LONG g_vb6_projClassCount = 0;

/* 「扫到哪儿」这一问的唯一答案 (账 #305 §B187): 登记腿与查表腿以前各写各的边界, 登记那一遍
 * 漏了夹界, 而计数在夹界检查之前自增且不回退 ⇒ 表满之后每登记一枚就往数组外多读一格。 */
static LONG vb6_ProjClassScanLimit(void) {
    LONG n = g_vb6_projClassCount;
    if (n > VB6_PROJCLASS_MAX) n = VB6_PROJCLASS_MAX;
    if (n < 0) n = 0;
    return n;
}

void vb6_RegisterProjectClassInstance(void* inst, const vb6_CoClassDesc* desc) {
    LONG i, hole = -1;
    if (!inst || !desc) return;
    /* 先线性找: 同一实例可能被重复登记 (DllMain + 工厂各一次), 直接覆盖即可。
     * 顺手记下第一枚空位 —— 注销腿腾出来的格必须能复用, 否则计数只增不减, 累计建过
     * VB6_PROJCLASS_MAX 枚实例之后整张表永久饱和。*/
    for (i = 0; i < vb6_ProjClassScanLimit(); i++) {
        if (g_vb6_projClass[i].inst == inst) {
            g_vb6_projClass[i].desc = desc;
            return;
        }
        if (hole < 0 && g_vb6_projClass[i].inst == NULL) hole = i;
    }
    if (hole >= 0 &&
        InterlockedCompareExchangePointer((PVOID volatile*)&g_vb6_projClass[hole].inst, inst, NULL) == NULL) {
        /* 空洞用 CAS 抢: 两枚并发登记同时看见同一个洞时只有一个占得到, 输的那枚落到下面新增槽
         * (它一个字都没碰这个洞 ⇒ 不会改坏赢家刚写的 desc)。desc 在 inst 之后写, 与新增槽那一支
         * 同形 —— 查表只认 inst, 这个窗口与既往「计数先动、字段后写」是同一个。*/
        g_vb6_projClass[hole].desc = desc;
        return;
    }
    i = InterlockedIncrement(&g_vb6_projClassCount) - 1;
    if (i < 0 || i >= VB6_PROJCLASS_MAX) return;      // 表满: 放弃登记 (退化成原行为)
    g_vb6_projClass[i].inst = inst;
    g_vb6_projClass[i].desc = desc;
}

/* 账 #481 = 第一百八十二刀: 登记腿的配对另一半。实例真被拆掉前撤掉登记, 那一格置空
 * 留给下一次登记复用; 计数是扫描上界的高水位, 这里不动它。 */
void vb6_UnregisterProjectClassInstance(void* inst) {
    LONG i;
    if (!inst) return;
    for (i = 0; i < vb6_ProjClassScanLimit(); i++) {
        if (g_vb6_projClass[i].inst == inst) {
            g_vb6_projClass[i].inst = NULL;
            g_vb6_projClass[i].desc = NULL;
            return;
        }
    }
}

const vb6_CoClassDesc* vb6_FindProjectClassDesc(const void* inst) {
    LONG i;
    if (!inst) return NULL;
    for (i = 0; i < vb6_ProjClassScanLimit(); i++)
        if (g_vb6_projClass[i].inst == inst) return g_vb6_projClass[i].desc;
    return NULL;
}

// 同上的 void* 出口: 给**看不到 vb6_CoClassDesc 完整定义**的 TU 用
// (如 vb6com_invoke.c 里只有前向声明, 直接声明返回该类型会 C2037)。
const void* vb6_FindProjectClassDescRaw(const void* inst) {
    return (const void*)vb6_FindProjectClassDesc(inst);
}

// Fix 099: 从 IDispatch 取回其 VB6 类实例裸指针 (Public 对象字段的 Property Let/Set
// 桥接用: 把客户端传来的对象写进字段). 只认本 RTL 产出的 vb6_ComObject (用自身
// vtable 指针判定), 外部 COM 对象/非对象返回 NULL (字段置空, 与 VB6 传 Nothing 等效).
void* vb6_ComObject_GetInstance(void* pdisp) {
    vb6_ComObject* obj;
    if (!pdisp) return NULL;
    obj = (vb6_ComObject*)pdisp;
    if (obj->vtable != &g_ComObjectVtable) {
        /* Fix <vbeclipse> rev14: 宿主对象回落。
         * cgen 对 `Set x = <Variant 值>` 一律生成
         *   (vb6_cls_X*)vb6_ComObject_GetInstance(vb6_VariantToObjectVal(v))
         * 而 Controls.Item/Add 给出的是 UserControl **宿主 HWND** (不是 vb6_ComObject)。
         * vb6_UC_InstanceOf 只遍历 vb6_ucRecs 比较 hwnd, 对任意指针安全 (含 NULL /
         * 裸 HWND / 非宿主指针), 命中即返回该宿主的 vb6_cls_<UC>* 实例。 */
        { extern void* vb6_UC_InstanceOf(void* hwnd);
          void* inst = vb6_UC_InstanceOf(pdisp); if (inst) return inst; }
        if (getenv("C3_IV_TRACE")) {
            fprintf(stderr, "[IV] REJECT pdisp=%p vtable=%p (expect %p)\n",
                    pdisp, (void*)obj->vtable, (void*)&g_ComObjectVtable);
            fflush(stderr);
        }
        return NULL;
    }
    if (getenv("C3_IV_TRACE")) {
        fprintf(stderr, "[IV] OK pdisp=%p inst=%p\n", pdisp, obj->vb6Instance);
        fflush(stderr);
    }
    return obj->vb6Instance;
}

/* 账 #305 §B187 第五腿: 「这枚值是不是我们那族包装器」—— deref 之前先证明可读,
 * 论证见 vb6comserver.h 声明处。命中 ⇒ 交回身份判据真; 其余一律 0 (原样交回调用点)。 */
int32_t vb6_ComObject_IsBoxed(void* v) {
    if (!v) return 0;
    if (IsWindow((HWND)v)) return 0;
    if (IsBadReadPtr(v, sizeof(void*))) return 0;
    return ((vb6_ComObject*)v)->vtable == &g_ComObjectVtable;
}

/* 账 #305 §B189 第一档: 见 vb6comserver.h 声明处的论证。char*→BSTR 走 CP_ACP 那一条,
 * 与本文件上方 progId / 类名那几处同一口径 (VB 的类模块名允许非 ASCII 标识符)。
 * 认不出 (不是我们的包装器 / desc 缺 / 名为空) 一律 NULL —— 调用方那一路要交回原答案。 */
BSTR vb6_ComObject_ClassNameBStr(void* v) {
    const vb6_ComObject* obj;
    const char* s;
    int wlen;
    BSTR r;
    if (!vb6_ComObject_IsBoxed(v)) return NULL;
    obj = (const vb6_ComObject*)v;
    s = obj->desc ? obj->desc->classVariable : NULL;
    if (!s || !s[0]) return NULL;
    wlen = MultiByteToWideChar(CP_ACP, 0, s, -1, NULL, 0);  /* -1 = 输入以 NUL 结尾 */
    if (wlen <= 1) return NULL;
    r = SysAllocStringLen(NULL, (UINT)wlen - 1);
    if (!r) return NULL;
    MultiByteToWideChar(CP_ACP, 0, s, -1, r, wlen);
    return r;
}

/* 账 #305 新账 #471 第一百七十四刀: 「这枚包装器是哪一枚工程类」的不分配那一档 —— 论证见
 * vb6comserver.h 声明处。三闸与身份判据全走兄弟那位主人 vb6_ComObject_IsBoxed, 这里只把
 * desc 里那根 ANSI 名交回去 (不复制、不分配、不归调用方释放)。 */
const char* vb6_ComObject_ClassNameA(void* v) {
    const vb6_ComObject* obj;
    const char* s;
    if (!vb6_ComObject_IsBoxed(v)) return NULL;
    obj = (const vb6_ComObject*)v;
    s = obj->desc ? obj->desc->classVariable : NULL;
    return (s && s[0]) ? s : NULL;
}

// ============================================================
// ExeComBridge 03: 工程类实例 → COM 调用实参 (VT_DISPATCH VARIANT*)
// ============================================================
// 触发场景: EXE/DLL 里 `New <本工程类>` 直接作为 COM 调用的实参
//   (VBMAN_DEMO: .Router.Reg "Demo", New bHello)
// cgen 原先对它生成通用打包 vb6_ComPackValue(vb6_cls_bHello_New()) —— 裸结构体
// 指针被当成 VT_DISPATCH 传出, 对端 (Dictionary 赋值 / Property Set) 取值或释放
// 时 AddRef, 把结构体首字段当 vtable 解引用 → 0xC0000005.
// 这里先经 vb6_ComObject_FromInstance 包装成真 IDispatch, 再转成 cgen 侧约定的
// VARIANT*. 引用计数: FromInstance 交给我们的 1 个引用直接转移给 VARIANT, 不额外
// AddRef —— VariantClear/Release 时正好归零, 包装器与实例随之释放.
void* vb6_ComPackVB6InstanceRaw(const char* classVariable, void* instance) {
    const vb6_CoClassDesc* desc;
    vb6_ComObject* obj;
    VARIANT* pv;

    if (!instance) return NULL;
    desc = vb6_FindCoClassDesc(classVariable);
    P189("[189] PACKRAW cls=%s inst=%p desc=%p\n", classVariable, instance, (void*)desc);
    // 未进 coclass 表 (Private / 非 MultiUse|SingleUse) 的类没有 IDispatch 方法表,
    // 包装也无从应答 GetIDsOfNames → 返回 NULL (等效 VB6 传 Nothing), 不冒充对象.
    if (!desc) return NULL;

    pv = (VARIANT*)calloc(1, sizeof(VARIANT));
    if (!pv) return NULL;
    obj = vb6_ComObject_FromInstance(desc, instance);
    if (!obj) {
        free(pv);
        return NULL;
    }
    VariantInit(pv);
    pv->vt = VT_DISPATCH;
    pv->pdispVal = (IDispatch*)obj;  /* 引用已由 FromInstance 提供, 所有权转移给 VARIANT */
    return (void*)pv;
}

/* 账 #305 新账 #407 / §B190 第一百七十一刀 = 「视图那一档怎么进 Variant」的唯一出口。
 * 为什么要按**指针**问身份: 视图 (vb6_iface_<I>) 是值结构体、静态类型是接口，实现类只有运行期
 * 才知道，而兄弟 vb6_ComPackVB6InstanceRaw (本文件 :726) 问的是**类名** —— 这一站手里没有名字。
 * 身份那张表本来就有唯一主人 vb6_FindProjectClassDesc (:642)，写者就是每枚工程类 _New 里那句
 * vb6_RegisterProjectClassInstance(me, desc)，所以这里只读、不另立一张。
 * 认不出身份 (没进 coclass 表的类 = Private / 非 MultiUse 与 SingleUse 那两档) 交回 VT_EMPTY，
 * 与兄弟那条「不冒充对象」的口径一字相同 —— 把裸实例指针塞进 dispatch 档就是 §B85 那族野 Release。
 * 所有权: FromBorrowedInstance 造新那一路把 refCount 停在 1 并写进实例的 __comObj (那一枚归实例)，
 * 已包过那一路自带一次 AddRef；vb6_VariantObject 另加一枚，正好与槽清时那一次 Release 相抵。 */
vb6_VARIANT vb6_VariantFromProjectInstance(void* instance) {
    vb6_VARIANT v;
    const vb6_CoClassDesc* desc;
    vb6_ComObject* wrapper;
    memset(&v, 0, sizeof(v));
    v.vt = vb6_vtEmpty;
    if (!instance) return v;
    desc = vb6_FindProjectClassDesc(instance);
    if (!desc) return v;
    wrapper = vb6_ComObject_FromBorrowedInstance(desc, instance);
    if (!wrapper) return v;
    return vb6_VariantObject((void*)wrapper);
}
