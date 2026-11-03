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
/* 账 #305 第三百四十二刀: IsBadReadPtr 在守卫页/页堆下不可靠 (探针本身可能触发
 * 系统侧异常路径), 换 VirtualQuery 的提交+可读档判定。区间可能跨页, 两端都查。 */
static int32_t vb6_ComIsReadable(const void* p, size_t n) {
    if (!p) return 0;
    MEMORY_BASIC_INFORMATION mbi;
    if (VirtualQuery(p, &mbi, sizeof(mbi)) != sizeof(mbi)) return 0;
    if (mbi.State != MEM_COMMIT) return 0;
    DWORD prot = mbi.Protect & 0xFF;
    if (prot == PAGE_NOACCESS) return 0;
    if (n > 1) return vb6_ComIsReadable((const char*)p + n - 1, 1);
    return 1;
}

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
    if (!vb6_ComIsReadable(disp, sizeof(void*))) return 0;
    void* vtbl = *(void**)disp;
    /* IDispatch: QueryInterface/AddRef/Release/GetTypeInfoCount/
     * GetTypeInfo/GetIDsOfNames/Invoke — 前 7 槽必须在可执行页里 */
    if (!vtbl || !vb6_ComIsReadable(vtbl, 7 * sizeof(void*))) return 0;
    void** slot = (void**)vtbl;
    return vb6_ComIsCodePtr(slot[0]) && vb6_ComIsCodePtr(slot[1])
        && vb6_ComIsCodePtr(slot[2]) && vb6_ComIsCodePtr(slot[5])
        && vb6_ComIsCodePtr(slot[6]);
}

/* 账 #305 第二百六十二刀 (丙): 晚绑定出口的「非 IDispatch 接收者」护栏。
 * 七个出口 (ComCall / ComCallByDispid / ComGetProp / ComGetPropArg /
 * ComSetProp / ComSetPropArg / ComSetRef) 以前对「接收者不是真 IDispatch」
 * 各行其是: 有的打一行 "not found" 就走人, 有的 (CallByDispid / NULL 接收者)
 * 一声不吭。静默 = 下回再漏时没人知道漏在哪一枚。
 * 这里收成一位主人: 计数 + 一行 stderr 标记 (C3_LATEBIND_MISS)。
 * **行为不变** (仍按原路返回 NULL / 空): 只把「静默」换成「响亮」。
 * 刻意不在这儿抛错 —— 无处理器的 vb6_RaiseError 会 exit(424) 把进程带走。
 * 判据「非 IDispatch 进晚绑定出口 = 0」按 stderr 里 C3_LATEBIND_MISS 的条数读。
 * 伪接收者 (字体 / Extender / 宿主对象: HWND / 控件集合 / Collection /
 * UC 实例) 由各自的岔口应答, 不算漏 —— 判据是 vb6_ComLateBindIsPseudo。 */
static int32_t g_com_latebind_miss = 0;

/* 账 (438 语义刀): 宿主模型答「成员不存在」时抛 438 之后, 链上的下一枚成员访问
 * (`rtfText.SelLength`) 仍会被生成代码照常求值 —— 接收者是那枚空值 ⇒ 再记一条
 * [C3_LATEBIND_MISS]。VB6 语义是**整条语句在 438 处终止**, 那条后续访问根本不存在。
 * 这里用一位「刚抛过 438」的标记把那一次链尾 miss 吸收掉 (只吸收一枚; 遇到非空
 * 接收者就先清标记, 后续真缺口照记)。 */
static int32_t g_com_438_dangling = 0;

int32_t vb6_ComLateBindMissCount(void) { return g_com_latebind_miss; }

/* 这枚接收者是不是「宿主伪对象」? 不是 IDispatch, 但宿主模型会应答。 */
static int32_t vb6_ComLateBindIsPseudo(const void* disp) {
    if (!disp) return 0;
    if (vb6_UC_IsFont(disp)) return 1;
    if (vb6_UC_IsExtender(disp)) return 1;
    if (vb6_Host_IsHostObject(disp)) return 1;
    return 0;
}

/* 账 (438 语义刀): 刚抛过 438 的那条语句在链上留下的**空接收者**不算漏 —— VB6 里整条
 * 语句在 438 处终止, 链尾那一次成员访问根本不存在。标记只吸收一枚; 走到这里时接收者
 * 不是空档 (链已走完) 就先清标记, 后续真缺口照记。
 * 这一格住在判定谓词里而**不在**护栏里: 护栏是纯报告出口 (void、不带控制流),
 * 见 scripts/check_com_latebind_guard.ps1 的 R3。 */
static int32_t vb6_ComAbsorb438(const void* disp) {
    if (!g_com_438_dangling) return 0;
    g_com_438_dangling = 0;
    return disp ? 0 : 1;
}

/* 护栏的唯一判据: 既不是「刚被 438 吞掉的链尾」, 也不是宿主伪对象, 又不是真 IDispatch。 */
static int32_t vb6_ComLateBindIsMiss(const void* disp) {
    if (vb6_ComAbsorb438(disp)) return 0;
    if (vb6_ComLateBindIsPseudo(disp)) return 0;
    return !vb6_ComIsDispatchable(disp);
}

/* 账 (状态栏/Output 面板刀): `C3_COM_TRACE=1` 时把三条 Rcv 的「成员名 + 接收者身份」打一行。
 * 为什么需要它: 同一个 `x.Text = "..."`
 *   · 走窗口语义 (SetWindowTextW 打在 OCX 的外壳窗上) 与
 *   · 走 OCX 自己的 IDispatch (FlatEdit 内层子窗才会变) 
 * 从外面看都「没报错」, 只有这一行能分出是哪一层接的。同理 `AddPane` 走宿主模型
 * 还是真 IDispatch、`StatusBar` 是不是被宿主臂截走, 都靠它。默认关闭, 零成本。 */
static void vb6_ComTraceRcv(const char* site, void* disp, const wchar_t* name) {
    static int on = -1;
    if (on < 0) on = GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0;
    if (!on) return;
    wchar_t cls[96]; cls[0] = 0;
    int isWin = (disp && IsWindow((HWND)disp)) ? 1 : 0;
    if (isWin) GetClassNameW((HWND)disp, cls, 96);
    fwprintf(stderr, L"[C3_COM_T] %hs '%ls' rcv=%p isWin=%d host=%d cls=%ls\n",
             site, name ? name : L"", disp, isWin,
             (disp && vb6_Host_IsHostObject(disp)) ? 1 : 0, cls);
    fflush(stderr);
}

static void vb6_ComLateBindMiss(const wchar_t* site, const wchar_t* name, void* disp) {
    g_com_latebind_miss++;
    fwprintf(stderr, L"[C3_LATEBIND_MISS #%ld] site=%ls name=\"%ls\" disp=%p kind=%ls\n",
             (long)g_com_latebind_miss, site, name ? name : L"(null)", disp,
             disp ? L"notdispatchable" : L"null");
    fflush(stderr);
}

/* 账 #305 第三百十七刀 (#600 的实测收口): 晚绑定的接收者可能是**另一张接口视图**。
 * 实证 (ComCtlsDemo Command2): Common.bas 的 PictureFromHandle 用
 *   OleCreatePictureIndirect(&PICD, IID_IPicture, ...) 造出 IPicture*, 而 VB6 在
 *   `Set <IPictureDisp 槽> = <IPicture>` 那一站做一次 QI; 我们照抄指针 ⇒ 下游
 *   `.Handle` 晚绑定把 IPicture 的 vtable 当 IDispatch 用, 第 5 槽 GetIDsOfNames
 *   落到 IPicture::get_Type 上, 把 `&IID_NULL` (只读) 当 short* 写 ⇒ 0xC0000005。
 * 独立探针 (本机 .local/k317/probe_pic.c, x64): QI(IID_IPictureDisp) ==
 *   QI(IID_IDispatch) == p+0x10 (换指针!), 在该指针上 GetIDsOfNames("Handle")
 *   得 dispid=0, Invoke 得 65579 = 图标的 hIcon。
 *
 * 修法 = VB6 那条规则本身: 晚绑定前先换到 IDispatch 视图。逐调用一发 QI 太贵,
 * 而「同一 vtable ⇒ 同一布局 ⇒ 同一偏移」是 COM 的硬性质 ⇒ 按 **vtable** 缓存
 * 「本指针到它 IDispatch 视图的字节偏移」: 真 IDispatch 缓存 0 (原路零改动);
 * QI 不认的伪 vtable (VTableHandle 那种手写表) 缓存 -1 (同样原路, 不能把它们
 * 当野指针拒掉); 只有异视图才真付一次 QI。取出的视图 AddRef、用完 Release
 * (IUnknown 三槽在所有视图里同址), 所以调用期间视图一定活着。 */
typedef struct { void* vtbl; int32_t off; } Vb6DispViewEnt;
#define VB6_DISPVIEW_CACHE 64
static Vb6DispViewEnt g_dispView[VB6_DISPVIEW_CACHE];
static int32_t g_dispViewCount = 0;
static int32_t g_dispViewNext = 0;

static int32_t vb6_ComDispViewOffset(void* disp) {
    if (!disp) return -1;
    if (!vb6_ComIsDispatchable(disp)) return -1;
    void* vtbl = *(void**)disp;
    for (int32_t i = 0; i < g_dispViewCount; i++)
        if (g_dispView[i].vtbl == vtbl) return g_dispView[i].off;
    int32_t off = -1;
    void* d = NULL;
    if (SUCCEEDED(((IUnknown*)disp)->lpVtbl->QueryInterface((IUnknown*)disp, &IID_IDispatch, &d)) && d) {
        off = (int32_t)((char*)d - (char*)disp);
        ((IUnknown*)d)->lpVtbl->Release((IUnknown*)d);
        if (off < 0) off = -1;   /* 负偏移理论上不存在; 当「不动」处理 */
    }
    Vb6DispViewEnt ent;
    ent.vtbl = vtbl; ent.off = off;
    if (g_dispViewCount < VB6_DISPVIEW_CACHE) g_dispView[g_dispViewCount++] = ent;
    else { g_dispView[g_dispViewNext] = ent;
           g_dispViewNext = (g_dispViewNext + 1) % VB6_DISPVIEW_CACHE; }
    return off;
}

/* 交回本次调用该用的接收者; *owned 非空 = 借出的一枚引用, 用完交给 ...Out。 */
static void* vb6_ComDispViewIn(void* disp, void** owned) {
    *owned = NULL;
    /* 账 #305 第三百四十二刀: 非可分派接收者 (BSTR/UDT/野指针) 在这里就拦下 ——
     * 原先 DispViewOffset 对它直接 QueryInterface, 间接调用字符串头冒充的
     * vtable 槽 => AV (GUISample AttachPane 链, av read 0x2e0032 实测)。
     * 拦下后原样交回, Rcv 的 getDispid 走 not-found 安全路 (与 ComCall 的
     * C3_LATEBIND_MISS 同口径, 响亮不崩)。 */
    if (disp && !vb6_ComLateBindIsPseudo(disp) && !vb6_ComIsDispatchable(disp)) {
        vb6_ComLateBindMiss(L"DispView", L"(receiver)", disp);
        return disp;
    }
    int32_t off = vb6_ComDispViewOffset(disp);
    if (off <= 0) return disp;
    void* d = (char*)disp + off;
    ((IUnknown*)d)->lpVtbl->AddRef((IUnknown*)d);
    *owned = d;
    return d;
}

static void vb6_ComDispViewOut(void* owned) {
    if (owned) ((IUnknown*)owned)->lpVtbl->Release((IUnknown*)owned);
}

/* 账 #305 第二百八十九刀 (§B164 新形): 「认不出这个名字」那一支的统一清点。
   打包来的实参 (vb6_ComPackVariant 每次一发一枚堆上 VARIANT) 交进这七个出口后就只有本文件
   那几条 clear 路接手 —— 早退不接, 就是整批发漏。空跑靶工程 8 秒 29 发 `Move`, 每发四枚。
   VT_BYREF 那一档的内容指向调用方的存储, 与成功路径 rgvarg 那圈同一个规矩: 不清内容。
   (在册: ComGetPropArg 的**成功**尾巴也从没 free 过结构体, ComSetProp/ComSetPropArg 只 free
    结构体不清内容 —— 那些路今天没有读数, 且要先证调用方不 reuse 那枚数组, 不在这一格动。 */
static void vb6_ComDropPackedArgs(VARIANT** a, int32_t argc) {
    if (!a) return;
    for (int32_t i = 0; i < argc; i++) {
        if (!a[i]) continue;
        if (!(a[i]->vt & VT_BYREF)) VariantClear(a[i]);
        free(a[i]);
    }
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

/* 账 #305 第二百九十刀: 出站实参按**被调方声明的形参类型**收窄。
   VB6 的晚绑定是 ITypeInfo 驱动的: 实参先折到形参档再进 Invoke。OLE 自己那套
   标准实现则在自己的 Invoke 里按 typelib 严格校验 —— 把 `.Height` (VT_R8) 递进
   一个 `typedef LONG` 的形参, 直接 DISP_E_TYPEMISMATCH, 于是 Common.bas 的
   `.Render hDC, x, y, cx, cy, 0, .Height, .Width, -.Height, ByVal 0&`
   一声不响地什么都没画 (ComCtlsDemo 状态条 panel1 那枚图标就是这么丢的)。
   判据 (RTL 外的产物级读数, .local/k294/render_probe): 同一个 IPictureDisp,
   cArgs=10 全 VT_I4 + DISPATCH_METHOD ⇒ S_OK 且真画出 20x20; 三个 himetric
   实参换成 VT_R8 ⇒ 0x80020005; 旗标带上 DISPATCH_PROPERTYGET ⇒ 0x8002000E。
   只收窄**数值变窄**这一档 (R8/R4 → 整型), 已经是整数档的实参一个都不动 ——
   这是本刀"零回归"的根据, 也是不碰 ByRef/BSTR/对象三档的理由。*/
static int vb6_ComIntParamVt(ITypeInfo* ti, const TYPEDESC* td) {
    if (!ti || !td) return 0;
    VARTYPE vt = td->vt;
    if (vt == VT_USERDEFINED) {
        ITypeInfo* ref = NULL;
        if (FAILED(ti->lpVtbl->GetRefTypeInfo(ti, td->hreftype, &ref)) || !ref) return 0;
        int out = 0;
        TYPEATTR* ta = NULL;
        if (SUCCEEDED(ref->lpVtbl->GetTypeAttr(ref, &ta)) && ta) {
            if (ta->typekind == TKIND_ALIAS) {
                VARTYPE av = ta->tdescAlias.vt;
                if (av == VT_I4 || av == VT_I2 || av == VT_INT ||
                    av == VT_UI4 || av == VT_UI2) out = (int)av;
            }
            ref->lpVtbl->ReleaseTypeAttr(ref, ta);
        }
        ref->lpVtbl->Release(ref);
        return out;
    }
    if (vt == VT_I4 || vt == VT_I2 || vt == VT_INT || vt == VT_UI4 || vt == VT_UI2) return (int)vt;
    return 0;
}

/* 每枚对象取一次 GetTypeInfo(0) 并记住 —— 绘制路径一秒能问几十次。 */
#define VB6_TI_CACHE_SIZE 64
static struct { IDispatch* disp; ITypeInfo* ti; } vb6_ti_cache[VB6_TI_CACHE_SIZE];
static int vb6_ti_cache_count = 0;

static ITypeInfo* vb6_ComTypeInfoOf(IDispatch* pDisp) {
    for (int i = 0; i < vb6_ti_cache_count; i++)
        if (vb6_ti_cache[i].disp == pDisp) return vb6_ti_cache[i].ti;
    ITypeInfo* ti = NULL;
    if (FAILED(pDisp->lpVtbl->GetTypeInfo(pDisp, 0, 0, &ti)) || !ti) {
        ti = NULL;
        if (FAILED(pDisp->lpVtbl->GetTypeInfo(pDisp, 0, LOCALE_USER_DEFAULT, &ti)) || !ti) return NULL;
    }
    if (vb6_ti_cache_count < VB6_TI_CACHE_SIZE) {
        vb6_ti_cache[vb6_ti_cache_count].disp = pDisp;
        vb6_ti_cache[vb6_ti_cache_count].ti = ti;
        vb6_ti_cache_count++;
    }
    return ti;
}

/* 账 #305 第291刀 (b): 旗标按被调方声明的 INVOKEKIND 走, 不按"有没有实参"猜。
   第一版写成 `argc > 0 ⇒ 只发 METHOD`, 当场打红两只: `test_com_default_prop` 与
   b433-10 —— Scripting.Dictionary 的默认成员 Item 是**带实参的 propget**,
   只发 METHOD 它答"找不到成员"。老行为 (METHOD|PROPERTYGET 一发齐递) 对
   propget 一直是对的, 只有真函数 (INVOKE_FUNC, 例如 IPictureDisp::Render) 才
   必须单发 METHOD。认不出成员的 (无 typeinfo / 无 FUNCDESC) 一律退回老行为,
   这样这条规则只碰"能证明是函数"的那一发。 */
static WORD vb6_ComFlagsFor(IDispatch* pDisp, DISPID dispid, int32_t argc) {
    WORD dflt = DISPATCH_METHOD | DISPATCH_PROPERTYGET;   /* 猜不出就按老行为两可齐递 */
    if (argc <= 0) return dflt;
    ITypeInfo* ti = vb6_ComTypeInfoOf(pDisp);
    if (!ti) return dflt;
    TYPEATTR* ta = NULL;
    if (FAILED(ti->lpVtbl->GetTypeAttr(ti, &ta)) || !ta) return dflt;
    WORD fl = 0;
    for (UINT f = 0; f < ta->cFuncs; f++) {
        FUNCDESC* fd = NULL;
        if (FAILED(ti->lpVtbl->GetFuncDesc(ti, f, &fd)) || !fd) continue;
        if (fd->memid == dispid) {
            if (fd->invkind == INVOKE_FUNC) fl = DISPATCH_METHOD;
            ti->lpVtbl->ReleaseFuncDesc(ti, fd);
            break;
        }
        ti->lpVtbl->ReleaseFuncDesc(ti, fd);
    }
    ti->lpVtbl->ReleaseTypeAttr(ti, ta);
    return fl ? fl : dflt;
}

static void vb6_ComNarrowArgs(IDispatch* pDisp, DISPID dispid, VARIANT** args, int32_t argc) {
    if (!pDisp || !args || argc <= 0) return;
    int widen = 0;
    for (int32_t i = 0; i < argc; i++)
        if (args[i] && (args[i]->vt == VT_R8 || args[i]->vt == VT_R4)) { widen = 1; break; }
    if (!widen) return;                       /* 没有浮点实参 ⇒ 这条路径零代价 */
    ITypeInfo* ti = vb6_ComTypeInfoOf(pDisp);
    if (!ti) return;
    TYPEATTR* ta = NULL;
    if (FAILED(ti->lpVtbl->GetTypeAttr(ti, &ta)) || !ta) return;
    for (UINT f = 0; f < ta->cFuncs; f++) {
        FUNCDESC* fd = NULL;
        if (FAILED(ti->lpVtbl->GetFuncDesc(ti, f, &fd)) || !fd) continue;
        if (fd->memid != dispid || !fd->lprgelemdescParam) {
            ti->lpVtbl->ReleaseFuncDesc(ti, fd);
            continue;
        }
        /* rgvarg[j] = 形参 (cParams-1-j); args[i] 落在 rgvarg[argc-1-i]
           ⇒ args[i] 对的是形参 cParams-argc+i (cParams==argc 时就是 i)。
           反过来写 (argc-1-i) 会整列错位: 第 7 个实参去对了第 4 个形参, 于是
           该收窄的三个 himetric 实参一个都没被认出来。 */
        for (int32_t i = 0; i < argc; i++) {
            int k = (int)(fd->cParams - argc + i);
            if (k < 0 || k >= (int)fd->cParams) continue;
            VARIANT* a = args[i];
            if (!a || (a->vt != VT_R8 && a->vt != VT_R4)) continue;
            int want = vb6_ComIntParamVt(ti, &fd->lprgelemdescParam[k].tdesc);
            if (!want) continue;
            VARIANT tmp;
            VariantInit(&tmp);
            if (SUCCEEDED(VariantChangeType(&tmp, a, 0, (VARTYPE)want))) {
                VariantClear(a);
                *a = tmp;
            } else {
                VariantClear(&tmp);
            }
        }
        ti->lpVtbl->ReleaseFuncDesc(ti, fd);
        break;
    }
    ti->lpVtbl->ReleaseTypeAttr(ti, ta);
}

// COM方法调用 (返回VARIANT*, 调用方需vb6_ComVarClear释放)
// args_void: VARIANT*[] (指向已分配VARIANT的指针数组), 每个元素由vb6_ComPackXxx分配
static void* vb6_ComCallRcv(void* disp, const wchar_t* methodName,
                            void* args_void, int32_t argc) {
    VARIANT** args = (VARIANT**)args_void;
    vb6_ComTraceRcv("ComCall", disp, methodName);
    /* 账 #305 第二百六十二刀 (丙): 非 IDispatch 接收者 —— 响亮 + 计数 (行为不变) */
    if (vb6_ComLateBindIsMiss(disp)) {
        vb6_ComLateBindMiss(L"ComCall", methodName, disp);
    }
    if (!disp) return NULL;
    /* Fix 112: 宿主对象 (窗体/控件 HWND, Controls 集合, Font 代理) 不是 IDispatch,
     * 直接解引用 lpVtbl 会 AV. 用 Win32 语义的宿主分派应答. */
    if (vb6_Host_IsHostObject(disp)) {
        char hout[64];   /* vb6_VARIANT (vb6com 单元用 Windows VARIANT, 不见 vb6_VARIANT 类型) */
        int32_t hhandled = vb6_Host_Call(disp, methodName, argc, args_void, hout);
        if (args) {
            for (int32_t hi = 0; hi < argc; hi++) { if (args[hi]) free(args[hi]); }
        }
        /* 账 (438 语义刀): -1 = 宿主模型没有这个成员 (uc_hostmodel_call.inc 末尾)。
           `v = Me.rtfText` 这类读在 cgen 里发的是 **Call** 站 (不是 GetProp),
           以前这里无条件答 Empty ⇒ 夹具里 Err.Number 读到 0、反向用例走不到错误分支。
           VB6 在这里抛 438 "Object doesn't support this property or method"。
           与 GetProp 岔口同口径: 抛完仍交一枚空 VARIANT, 让调用方继续 (Resume Next 语义)。 */
        if (hhandled < 0) {
            g_com_438_dangling = 1;
            vb6_RaiseError(438, SysAllocString(L"Object doesn't support this property or method"));
            vb6_Host_ClearVariant(&hout);
            VARIANT* hempty = (VARIANT*)calloc(1, sizeof(VARIANT));
            return (void*)hempty;
        }
        VARIANT* hres = (VARIANT*)calloc(1, sizeof(VARIANT));
        vb6_Host_ToWinVariant(&hout, hres);
        vb6_Host_ClearVariant(&hout);
        return (void*)hres;
    }
    /* 账 #305 第二百九十五刀: Extender 伪对象的**方法形** —— 把「静默丢且护栏数不到」改成
       「响亮地丢」。
       GetProp / SetProp 两站从 Fix 168 起各有 Extender 岔口, 唯独 ComCall 一个都没有 ⇒
       `.Extender.Move l, t, w, h` (UC 自己量完尺寸来请求摆位的那一发) 一路掉到 vb6_getDispid
       那扇 IDispatch 门上, 打一行 method "Move" not found 就丢。而 vb6_ComLateBindIsPseudo
       的豁免句「伪接收者由各自的岔口应答」对 Extender 是**无条件**成立的 ⇒ 这 26 发在护栏的
       计数里一条都不见 (实测靶工程 18 秒: C3_LATEBIND_MISS 只有 5 条, 全是另一族
       OnControlInfoChanged; 295 刀埋点量到 Move 落在 CoolBar1 x10 / ComboBoxW1 x5 /
       ImageCombo1 x4 / StatusBar1 x3 / MonthView1 x3 / Slider1 x1)。
       **写侧这一刀刻意不接** —— 本刀先接过一次 (四档交 vb6_ControlMove), 逐窗口归因后是三条
       不同的塌法: CoolBar1 宿主 900→552 缇 (原生 rebar 60→37, VB6 是 55; 它量的是 band 行高,
       那一路输入仍欠在第二百九十刀); ComboBoxW1 宿主 288↔504 缇来回摆 5 发都不收敛;
       ImageCombo1 交来的四档是**缇**而它容器 (CoolBar, 声明 ScaleMode=3 像素) 要的是像素,
       于是 3672 缇被当 3672 像素落回去, 宿主一度长成 3672x324。
       ⇒ 顺序仍按第二百八十八刀定的那一串: 先定「几何的父窗口该问谁」(今天走 GetParent,
       交不出 UC 自己声明的 ScaleMode), 再谈落地写。这一刀只把看不见改成看得见。
       **账 #305 第三百零一刀: 那两条前置都落地了, `Move` 接回** —— 第二百九十七刀把单位那一问
       收到容器主人 (不再问 GetParent), 第二百九十九刀把 authored 容器戳到窗口上, 第二百九十刀
       补了 ReBar band 行高那一路输入。本条只补通道, 兜底一概不抄: 四档原样交 vb6_ControlMove。 */
    if (vb6_UC_IsExtender(disp)) {
        if (methodName && wcscmp(methodName, L"Move") == 0) {
            extern int32_t vb6_UC_ExtenderMoveApply(double l, double t, double w, double h, int32_t mask);
            double v4[4] = { 0, 0, 0, 0 };
            int32_t mask = 0;
            for (int32_t i = 0; i < argc && i < 4; i++) {
                VARIANT scratch;
                if (!args || !args[i]) continue;
                VariantInit(&scratch);
                if (SUCCEEDED(VariantChangeType(&scratch, args[i], 0, VT_R8))) {
                    v4[i] = scratch.dblVal;
                    mask |= (1 << i);
                }
                VariantClear(&scratch);
            }
            vb6_UC_ExtenderMoveApply(v4[0], v4[1], v4[2], v4[3], mask);
            vb6_ComDropPackedArgs(args, argc);
            return NULL;
        }
        vb6_ComLateBindMiss(L"ExtenderCall", methodName, disp);
        if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
            fwprintf(stderr, L"[C3_EXT] Call name=\"%ls\" argc=%ld dropped=1\n",
                     methodName ? methodName : L"(null)", (long)argc);
            fflush(stderr);
        }
        vb6_ComDropPackedArgs(args, argc);
        return NULL;
    }
    /* Fix 191: IUnknown 三法在**所有** COM 接口的 vtable 里固定在槽 0/1/2,

     * 自定义接口 (IOleInPlaceActiveObject / IOleObject 这类无 IDispatch 的) 也一样。     * 而按名后期绑定要先 GetIDsOfNames —— 那是 IDispatch 的槽 5。对手写伪 vtable
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
        /* 账 #305 第三百三十六刀: VB6 默认成员兜底。接收者没有该方法时, VB6 会沿
         * typelib 的 DISPID 0 默认属性再走一层 —— GUISample frmMain:
         *   With ControlFile.CommandBar
         *       .Add xtpControlButton, ...   ' = CommandBar.Controls.Add(...)
         *   End With
         * ICommandBar 没有 Add, 默认属性是 Controls (DISPID 0); 此前直接报
         * "method not found" 返回 NULL => 菜单/工具栏项全部丢失, 下游对 NULL 的
         * Control.IconId 等写接二连三炸错误 13。名字查找失败后取 DISPID 0 的对象
         * (须 VT_DISPATCH) 原样重调一次; 深度旗标防环。 */
        static int defMemberDepth336 = 0;
        VARIANT defv336;
        VariantInit(&defv336);
        DISPPARAMS dprop336;
        memset(&dprop336, 0, sizeof(dprop336));
        HRESULT hrDef336 = pDisp->lpVtbl->Invoke(pDisp, DISPID_VALUE, &IID_NULL,
            GetUserDefaultLCID(), DISPATCH_METHOD | DISPATCH_PROPERTYGET,
            &dprop336, &defv336, NULL, NULL);
        if (SUCCEEDED(hrDef336) && defv336.vt == VT_DISPATCH && defv336.pdispVal
            && (void*)defv336.pdispVal != (void*)pDisp && defMemberDepth336 == 0) {
            defMemberDepth336 = 1;
            if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
                fprintf(stderr, "[C3_COM] defmember hit: %ls disp=%p defobj=%p\n",
                        methodName, (void*)pDisp, (void*)defv336.pdispVal); fflush(stderr);
            }
            void* r336 = vb6_ComCall((void*)defv336.pdispVal, methodName, args_void, argc);
            defMemberDepth336 = 0;
            VariantClear(&defv336);
            return r336;
        }
        if (GetEnvironmentVariableW(L"C3_COM_TRACE", NULL, 0) > 0) {
            fprintf(stderr, "[C3_COM] defmember miss: %ls disp=%p hr=0x%08lX vt=%d\n",
                    methodName, (void*)pDisp, (unsigned long)hrDef336, defv336.vt); fflush(stderr);
        }
        VariantClear(&defv336);
        vb6_ComDropPackedArgs(args, argc);   /* 账 #305 第二百八十九刀: 早退之前清点来路 */
        fwprintf(stderr, L"vb6_ComCall: method \"%ls\" not found\n", methodName);
        return NULL;
    }

    /* 收窄必须发生在抄进 DISPPARAMS 之前 —— dp.rgvarg 是逐枚 VARIANT 的**值拷贝**,
       改 args[] 影响不到已经抄好的那一份 (第一版就栽在顺序上: 三个实参都识别成该收窄,
       Invoke 依旧 0x80020005)。 */
    vb6_ComNarrowArgs(pDisp, dispid, args, argc);

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

    /* 账 #305 第291刀 (b): 旗标按被调方声明的 INVOKEKIND 走。 */
    WORD callFlags = vb6_ComFlagsFor(pDisp, dispid, argc);
    HRESULT hr = vb6_ComInvoke(pDisp, dispid, callFlags, &dp, result, &excep, &argErr);

    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
        long val = -1;
        if (result) {
            if (result->vt == VT_I4) val = result->lVal;
            else if (result->vt == VT_BOOL) val = result->boolVal;
        }
        fprintf(stderr, "[C3_COM] Call '%ls' hr=0x%08lX vt=%d val=%ld pdisp=%p\n",
                methodName, (unsigned long)hr, result ? result->vt : -1, val,
                (result && result->vt == VT_DISPATCH) ? (void*)result->pdispVal : NULL);
        if (wcscmp(methodName, L"SetScrollBars") == 0) {
            for (int32_t ai = 0; ai < argc; ai++)
                fprintf(stderr, "[C3_COM]   arg[%ld] vt=%d val=%ld\n", (long)ai,
                        args[ai] ? args[ai]->vt : -1,
                        (long)((args[ai] && args[ai]->vt == VT_I4) ? args[ai]->lVal : 0));
        }
        fflush(stderr);
    }

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
// 宿主对象 (窗体/控件 HWND, Controls 集合, Font 代理) 不是 IDispatch, 与
// vb6_ComCall 的 Fix 112 同口径应答 —— 旧路径按名 L"Item" 走 vb6_Host_Call,
// 这里保持完全一致的宿主分派语义 (DISPID_VALUE 对宿主无意义)。
static void* vb6_ComCallByDispidRcv(void* disp, int32_t dispid,
                                    void* args_void, int32_t argc) {
    VARIANT** args = (VARIANT**)args_void;
    /* 账 #305 第二百六十二刀 (丙): 非 IDispatch 接收者 —— 响亮 + 计数 (行为不变) */
    if (vb6_ComLateBindIsMiss(disp)) {
        vb6_ComLateBindMiss(L"ComCallByDispid", L"(byDispid)", disp);
    }
    if (!disp) return NULL;
    /* Fix 同 vb6_ComCall: 宿主对象分派 */
    if (vb6_Host_IsHostObject(disp)) {
        return vb6_ComCall(disp, L"Item", args_void, argc);
    }
    /* 与 vb6_getDispid 的 Fix 191 同口径: 非真 IDispatch 不能解引用 lpVtbl,
     * 否则进程直接重启 (0xC0000005)。返回 NULL, 由调用侧按空 Variant 收场。 */
    if (!vb6_ComIsDispatchable(disp)) return NULL;
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

// ============================================================
// 账 (Outlook2013 frmInbox:331): **默认成员**调用返回对象 —— `coll(idx)`。
// VB6 里 `coll(idx)` 编到默认成员 (DISPID_VALUE = 0), 从不按名字 "Item" 找。
// 生成器此前一律发 `vb6_ComCallObject(recv, L"Item", …)`, 对 VB6 Collection /
// ADODB.Recordset 恰好命中 (它们的默认成员名字就叫 Item), 但对 Codejock 的集合
// 一律落空: ReportColumns 的默认成员叫 Column、Actions 的默认成员不叫 Item ——
// getDispid("Item") = DISPID_UNKNOWN ⇒ Invoke 不发 ⇒ 交回 NULL ⇒
//   wndReportControl.GroupsOrder.Add <NULL>  →  Run-time error 13 (类型不匹配)。
// 口径: 名字 "Item" 能解析就仍按名走 (RTL 伪集合 —— ListImages/Panels/Nodes 一族 ——
// 只登记名字, 没有 0 号成员); 解析不到才按 VB6 语义发 DISPID 0。
// ============================================================
/* 账 (Outlook2013 frmMain:207 运行错误 87): 先按类型信息认**默认成员** (memid 0),
   认得到就直发 DISPID_VALUE, 不再按名 "Item" 找。实测 Codejock ICommandBarActions
   只有 memid {0(Action),1,2,5,6,-4}, 但它的 GetIDsOfNames 在 LOCALE_NEUTRAL(0) 上
   把并不存在的 "Item" 映射成 dispid 3 (USER/SYSTEM 两档才是 DISP_E_UNKNOWNNAME);
   照名走 ⇒ Invoke 3 ⇒ Codejock 抛 helpContext 87「属性名称无效」, 未接住的运行错误
   直接把进程带走。VB6 里 `obj(args)` 本来就编成 DISPID_VALUE, 名字 "Item" 只在
   认不出类型信息 (RTL 伪集合只登记名字) 时兜底。 */
static int vb6_ComHasDefaultMemid(IDispatch* pDisp) {
    ITypeInfo* ti = vb6_ComTypeInfoOf(pDisp);
    if (!ti) return -1;                       /* 无类型信息: 认不出, 交回老口径 */
    TYPEATTR* ta = NULL;
    if (FAILED(ti->lpVtbl->GetTypeAttr(ti, &ta)) || !ta) return -1;
    int has = 0;
    for (UINT f = 0; f < ta->cFuncs; f++) {
        FUNCDESC* fd = NULL;
        if (FAILED(ti->lpVtbl->GetFuncDesc(ti, f, &fd)) || !fd) continue;
        if (fd->memid == DISPID_VALUE) has = 1;
        ti->lpVtbl->ReleaseFuncDesc(ti, fd);
        if (has) break;
    }
    ti->lpVtbl->ReleaseTypeAttr(ti, ta);
    return has;
}

static void* vb6_ComCallObjectDefaultRcv(void* disp, void* args_void, int32_t argc) {
    if (!disp) return NULL;
    DISPID d = DISPID_UNKNOWN;
    /* 宿主对象 (窗体/控件槽/Controls 集合) 不是真 IDispatch: 交给 vb6_ComCallByDispid
     * 的宿主分派 (它内部转 vb6_ComCall(…, L"Item", …)), 这里不碰 vtable。 */
    if (!vb6_Host_IsHostObject(disp) && vb6_ComIsDispatchable(disp)) {
        int hasDefault = vb6_ComHasDefaultMemid((IDispatch*)disp);
        if (hasDefault != 1) d = vb6_getDispid((IDispatch*)disp, L"Item");
    }
    if (d == DISPID_UNKNOWN) d = DISPID_VALUE;

    VARIANT* pv = (VARIANT*)vb6_ComCallByDispid(disp, (int32_t)d, args_void, argc);
    if (!pv) return NULL;
    void* obj = NULL;
    if (pv->vt == VT_DISPATCH) {
        /* 与 vb6_ComCallObject 同口径: 解包 RTL 伪对象, 转移引用所有权 */
        extern void* vb6_UC_UnwrapHost(void* obj);
        obj = vb6_UC_UnwrapHost((void*)pv->pdispVal);
    } else if (pv->vt == VT_UNKNOWN) {
        pv->punkVal->lpVtbl->QueryInterface(pv->punkVal, &IID_IDispatch, &obj);
    } else if (pv->vt != VT_EMPTY && pv->vt != VT_NULL) {
        VARIANT vObj;
        VariantInit(&vObj);
        if (SUCCEEDED(VariantChangeType(&vObj, pv, 0, VT_DISPATCH))) {
            obj = (void*)vObj.pdispVal;
        }
        VariantClear(&vObj);
    }
    vb6_ComVarFree(pv);
    return obj;
}
/* 账 (Outlook2013 左面板 VisualTheme): 宿主模型只认窗口层的 Extender 成员;
 * `For Each Ctrl In Form.Controls: Set CtrlCaption = Ctrl: CtrlCaption.VisualTheme = v`
 * 交出的接收者是控件**窗口** (Codejock 还有内部子窗口), 控件自有属性以前被
 * 宿主模型静默丢弃。这里把窗口换回控件 IDispatch, 但**只在**控件类型库确实认得
 * 这个名字时才换 —— Extender 那一档 (几何/可见性/标题) 仍交宿主模型, 行为不变。 */
static int32_t vb6_ComIsExtenderMember(const wchar_t* name) {
    static const wchar_t* const kExt[] = {
        L"Left", L"Top", L"Width", L"Height", L"Move", L"Visible", L"Enabled",
        L"Caption", L"Text", L"Tag", L"Name", L"Index", L"Parent", L"Container",
        L"hWnd", L"hWndUserControl", L"Object", L"Extender", L"TabIndex", L"TabStop",
        L"DragIcon", L"DragMode", L"OLEDropMode", L"WhatsThisHelpID", L"HelpContextID"
    };
    for (int i = 0; i < (int)(sizeof(kExt) / sizeof(kExt[0])); i++)
        if (_wcsicmp(name, kExt[i]) == 0) return 1;
    return 0;
}

static void* vb6_ComOcxRcvOf(void* disp, const wchar_t* name) {
    if (!disp || !name) return NULL;
    if (vb6_ComIsExtenderMember(name)) return NULL;
    /* Controls 集合交出的项常是宿主包装器: 先透明解包, 解包结果可能是控件对象本身 */
    extern void* vb6_UC_UnwrapHost(void* obj);
    void* raw = vb6_UC_UnwrapHost(disp);
    if (!raw) raw = disp;
    void* ocx = vb6_OcxHost_DispatchOf(raw);
    if (!ocx) return NULL;
    if (vb6_getDispid((IDispatch*)ocx, name) == DISPID_UNKNOWN) return NULL;
    return ocx;
}

/* 账 (Outlook2013 左面板 VisualTheme): VB6 的
 *   `Dim CtrlCaption As ShortcutCaption` … `Set CtrlCaption = Ctrl`
 * 是按**声明接口** QI, 控件类型不符 ⇒ 错误 13; 工程里那句多在
 * `On Error Resume Next` 里, 于是赋值根本没发生 (VB6 实测: 该循环对
 * CommandBars / DockingPaneManager 之类毫无作用)。
 * C3 的 Controls 集合交出的是裸 IDispatch, 赋值是一句裸指针拷贝 ⇒ 循环里
 * 每一次 `CtrlCaption.VisualTheme = v` 都会打到 CommandBars 自己, 把 ribbon
 * 主题改坏 (实测: 整个 ribbon 变黑)。这里补上同一道类型检查:
 * 是本宿主托管的 OCX 控件时按控件类名比对, 不符返回 NULL (调用点据此跳过);
 * 其余对象 (普通 COM 对象 / 工程类 / 宿主包装器) 一律原样通过 —— 不改判。 */
void* vb6_ComObjectTypeCast(void* obj, const wchar_t* typeName) {
    if (!obj || !typeName) return obj;
    wchar_t ctrlType[128];
    if (vb6_OcxHost_TypeNameOf(obj, ctrlType, 128) && _wcsicmp(ctrlType, typeName) != 0)
        return NULL;
    return obj;
}

/* 账 (B3 StatusBar.Visible): VB6 对"带窗口的对象"用的是**宿主 Extender 语义** ——
 * `Visible` 读写的就是窗口可见性本身, 不是 OCX 内部标志。
 * 实测 (vb6_vc6 参照, 1 次启动): `CommandBars.StatusBar` 的 hWnd=3474498,
 * Visible=True/False 与 IsWindowVisible=1/0 逐次同步 (`sb2.Visible=False` 之后窗口
 * 真的隐藏, 回读 False; 再置 True 窗口恢复)。
 * 而同一份调用在 C3 下 (c3_vc6) `put_Visible` 回 S_OK 却不改 OCX 标志, get 恒 False。
 * 该对象在类型库里只有 IStatusBar.hWnd (memid=-515), 没有 Visible 成员 ——
 * 故宿主侧按 hWnd 的真窗口状态应答, 与 VB6 同语义。
 * 只对**非宿主控件**生效: 宿主控件的 Visible 由 vb6_OcxHost_* 先答 (上面的岔口)。 */
static int32_t vb6_ComObjWindow(void* disp, HWND* out) {
    *out = NULL;
    if (!disp || !vb6_ComIsDispatchable(disp)) return 0;
    DISPID did = vb6_getDispid((IDispatch*)disp, L"hWnd");
    if (did == DISPID_UNKNOWN) return 0;
    DISPPARAMS dp;
    memset(&dp, 0, sizeof(dp));
    VARIANT v;
    VariantInit(&v);
    EXCEPINFO ex;
    memset(&ex, 0, sizeof(ex));
    UINT ae = 0;
    HRESULT hr = vb6_ComInvoke((IDispatch*)disp, did,
                               (DISPATCH_METHOD | DISPATCH_PROPERTYGET), &dp, &v, &ex, &ae);
    intptr_t h = 0;
    if (SUCCEEDED(hr)) {
        if (v.vt == VT_I4) h = v.lVal;
        else if (v.vt == VT_I8) h = (intptr_t)v.llVal;
        else if (v.vt == VT_INT) h = v.intVal;
    }
    VariantClear(&v);
    if (!h || !IsWindow((HWND)h)) return 0;
    *out = (HWND)h;
    return 1;
}

// COM属性Get (返回VARIANT*)
static void* vb6_ComGetPropRcv(void* disp, const wchar_t* propName) {
    /* 账 #305 第二百六十二刀 (丙): 非 IDispatch 接收者 —— 响亮 + 计数 (行为不变) */
    vb6_ComTraceRcv("ComGetProp", disp, propName);
    if (vb6_ComLateBindIsMiss(disp)) {
        vb6_ComLateBindMiss(L"ComGetProp", propName, disp);
    }
    if (!disp) return NULL;
    /* 账 (Outlook2013 主区 Report 盖住 Calendar): VB6 的 `控件.Visible` 是**容器
       Extender** 的账, 实际落的是控件自己的窗口。而这里拿到的是控件对象
       (IDispatch*), 硬发下去只是问了控件自身一个 dispid=2 的属性 —— 窗口照样可见。
       先问 OCX 宿主: 命中 (返回 1) 就按容器语义答复。 */
    if (propName && _wcsicmp(propName, L"Visible") == 0) {
        int32_t vis = 0;
        if (vb6_OcxHost_GetVisible(disp, &vis)) {
            VARIANT* vres = (VARIANT*)calloc(1, sizeof(VARIANT));
            if (vres) {
                vres->vt = VT_BOOL;
                vres->boolVal = vis ? VARIANT_TRUE : VARIANT_FALSE;
            }
            return (void*)vres;
        }
        /* 账 (B3): 非宿主控件、但对象自带 hWnd → 读该窗口自己的可见位 (VB6 语义)。
           用 WS_VISIBLE 而不是 IsWindowVisible: 后者连带祖先链, 在 Form_Load 阶段
           父窗还没 ShowWindow 时会读成 False, 与 VB6 的逐控件标志口径不符
           (VB6 参照在 Form_Load 内回读 True)。 */
        HWND owh = NULL;
        if (vb6_ComObjWindow(disp, &owh)) {
            LONG st = GetWindowLongW(owh, GWL_STYLE);
            VARIANT* wres = (VARIANT*)calloc(1, sizeof(VARIANT));
            if (wres) {
                wres->vt = VT_BOOL;
                wres->boolVal = (st & WS_VISIBLE) ? VARIANT_TRUE : VARIANT_FALSE;
            }
            return (void*)wres;
        }
    }
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
                /* 账 #305 第二百七十六刀 kind 5 = 句柄: 按**指针宽度**取, 交 VT_I8。
                   落进 default 的 int32 那档就是 x64 上把 GDI 句柄砍掉高 32 位
                   (与 Fix 296 给 Extender 对象成员立的同一条规矩); 而 VB6 的
                   IFont.hFont 虽在类型库里标 Long, 语料一律把它交给 As LongPtr 的
                   形参 (Common.bas:231 那份 PtrSafe 声明), 所以交 VT_I8 才是应答。 */
                case 5: res->vt = VT_I8;   res->llVal = (LONG_PTR)*(void**)fp; break;
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
                // 对象成员 (kind 4): 传**指针本身**, 不 SysAllocString。此前它们报的是
                // kind 0 = BSTR, 于是这条分支去读对象结构体的头当 BSTR 长度 —— 会炸的读。
                // 语料里 `With UserControl.Extender` 只碰 Height/Left/Top/Width, 这条分支
                // 对对象成员今天没有覆盖, 是按字段类型 (void*) 摆平的 (Fix 296 / 3b 前置)。
                case 4: eres->vt = VT_DISPATCH; eres->punkVal = (IUnknown*)(*(void**)efp); break;
                default: eres->vt = VT_I4;  eres->lVal = *(int32_t*)efp; break;
            }
        }
        return (void*)eres;
    }
    /* 账 (Outlook2013 左面板 VisualTheme): 控件窗口 → 控件 IDispatch (只换控件自有成员) */
    { void* ocxRcv = vb6_ComOcxRcvOf(disp, propName); if (ocxRcv) disp = ocxRcv; }
    /* Fix 112: 宿主对象分派 (见 vb6_ComCall 注释) */
    if (vb6_Host_IsHostObject(disp)) {
        char hout[64];
        int32_t hgot = vb6_Host_GetProp(disp, propName, hout);
        /* 账 (438 语义刀): -1 = 宿主模型没有这个成员 (uc_hostmodel_getprop.inc 末尾)。
           VB6 在这里抛 **438 "Object doesn't support this property or method"**,
           由 On Error Resume Next 吞掉; 以前答 Empty 后链上 `.SelLength` 再记一条
           latebind miss ⇒ GUISample 里 rtfText/SelLength 成对噪声 (274×2)。 */
        if (hgot < 0) {
            g_com_438_dangling = 1;
            vb6_RaiseError(438, SysAllocString(L"Object doesn't support this property or method"));
            VARIANT* hempty = (VARIANT*)calloc(1, sizeof(VARIANT));
            return (void*)hempty;
        }
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

    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
        long val = -1;
        if (result) {
            if (result->vt == VT_I4) val = result->lVal;
            else if (result->vt == VT_BOOL) val = result->boolVal;
        }
        fprintf(stderr, "[C3_COM] GetProp '%ls' rcv=%p hr=0x%08lX vt=%d val=%ld\n",
                propName, disp, (unsigned long)hr, result ? result->vt : -1, val);
        fflush(stderr);
    }

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComGetProp");
        if (result) { VariantClear(result); free(result); result = NULL; }
    }

    return (void*)result;
}


// COM属性Get带参数 (参数化属性读取, 如Dictionary.Item(key))
// 使用DISPATCH_METHOD|DISPATCH_PROPERTYGET组合标志
// 返回VARIANT* (调用方需vb6_ComVarClear释放)
static void* vb6_ComGetPropArgRcv(void* disp, const wchar_t* propName,
                                  void* args_void, int32_t argc) {
    VARIANT** args = (VARIANT**)args_void;
    /* 账 #305 第二百六十二刀 (丙): 非 IDispatch 接收者 —— 响亮 + 计数 (行为不变) */
    if (vb6_ComLateBindIsMiss(disp)) {
        vb6_ComLateBindMiss(L"ComGetPropArg", propName, disp);
    }
    if (!disp) return NULL;
    IDispatch* pDisp = (IDispatch*)disp;

    DISPID dispid = vb6_getDispid(pDisp, propName);
    if (dispid == DISPID_UNKNOWN) {
        vb6_ComDropPackedArgs(args, argc);   /* 账 #305 第二百八十九刀 */
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
static void vb6_ComSetPropRcv(void* disp, const wchar_t* propName, void* value_void) {
    /* 账 #305 第二百六十二刀 (丙): 非 IDispatch 接收者 —— 响亮 + 计数 (行为不变) */
    vb6_ComTraceRcv("ComSetProp", disp, propName);
    if (vb6_ComLateBindIsMiss(disp)) {
        vb6_ComLateBindMiss(L"ComSetProp", propName, disp);
    }
    /* 账 (Outlook2013 主区 Report 盖住 Calendar): `控件.Visible = v` 是**容器
       Extender** 的账。此处 disp 是控件对象 (IDispatch*), 交给控件自身只会改它
       内部的一个属性, 承载它的 AfxOleControl 窗口照旧可见 ⇒ Report 盖住 Calendar。
       先问 OCX 宿主: 命中 (返回 1) 就按容器语义落窗口/站点。 */
    if (disp && propName && _wcsicmp(propName, L"Visible") == 0 && value_void) {
        int32_t vis = (int32_t)vb6_VariantToDouble(*(VARIANT*)value_void);
        if (vb6_OcxHost_SetVisible(disp, vis)) {
            free(value_void);
            return;
        }
        /* 账 (B3): 非宿主控件、自带 hWnd → 落窗口可见性 (VB6 Extender 语义)。
           实测 VB6 参照 `sb2.Visible=False` 之后窗口 IsWindowVisible=0。 */
        HWND owh = NULL;
        if (vb6_ComObjWindow(disp, &owh)) {
            ShowWindow(owh, vis ? SW_SHOW : SW_HIDE);
            free(value_void);
            return;
        }
    }
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
                /* 第二百七十六刀 kind 5 = hFont 那一格: **按只读处理**。
                   stdole2.tlb 里 hFont 只有 propget (IFont 没有 putref_hFont), VB6 里
                   `Font.hFont = x` 是编译期错误 ⇒ 语料 0 处; 落进 default 那档会把
                   一个 int32 写进指针宽度的缓存格、而快照不跟着改 ⇒ 下一发问句
                   交出半新半旧的句柄。不写它、也不报数 (报数是给"认不出"的, 这一档认得出)。 */
                case 5: break;
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
                // 对象成员 (kind 4): 原样搬指针, **不** SysAllocString (不接管所有权),
                // 也不走 int32 截断 (x64 上把 64 位指针砍成 32 位)。
                // 同 ComGetProp 那条: `With UserControl.Extender` 对对象成员今天没有语料覆盖,
                // 这是按字段类型 (void*) 摆平的 (Fix 296 / 3b 前置)。
                case 4: *(void**)efp = (void*)v.punkVal; break;
                default: *(int32_t*)efp = (int32_t)vb6_VariantToDouble(v); break;
            }
            /* 账 #305 第二百八十七刀: `With UserControl.Extender: .Visible = v` 这一路
               与 uc_controls.c 那一条按名的写手同源, 存完字段都交给同一位主人上窗口。 */
            if (ekind == 2 && wcscmp(propName, L"Visible") == 0) vb6_UC_ExtenderVisibleApplied();
            /* 账 #305 第二百九十二刀: 几何那一档的第二位写手 (第一位是 uc_controls.c 按名那条)。 */
            else if (ekind == 3) vb6_UC_ExtenderGeomApplied(propName);

        }
        free(value_void);
        return;
    }
    /* 账 (Outlook2013 左面板 VisualTheme): 控件窗口 → 控件 IDispatch (只换控件自有成员) */
    { void* ocxRcv = vb6_ComOcxRcvOf(disp, propName); if (ocxRcv) disp = ocxRcv; }
    /* Fix 112: 宿主对象分派 (见 vb6_ComCall 注释) */
    if (disp && vb6_Host_IsHostObject(disp)) {
        if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
            fprintf(stderr, "[C3_COM] SetProp '%ls' rcv=%p -> HOST\n", propName, disp);
            fflush(stderr);
        }
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

    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_COM] SetProp '%ls' rcv=%p -> COM dispid=%ld hr=0x%08lX\n",
                propName, disp, (long)dispid, (unsigned long)hr);
        fflush(stderr);
    }

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComSetProp");
    }
    free(value_void);  /* 释放ComPackXxx分配的堆VARIANT结构体 */
}

// P25: COM参数化属性Put (如dic.Item(key) = value)
// propName: 属性名 (如"Item"), args: 索引参数数组, argc: 索引参数个数
// value: 新值(已打包为堆VARIANT*, 同ComSetProp)
static void vb6_ComSetPropArgRcv(void* disp, const wchar_t* propName, void** args, int32_t argc, void* value_void) {
    VARIANT varValue;
    VariantInit(&varValue);
    if (value_void) varValue = *(VARIANT*)value_void;
    /* 账 #305 第二百六十二刀 (丙): 非 IDispatch 接收者 —— 响亮 + 计数 (行为不变) */
    if (vb6_ComLateBindIsMiss(disp)) {
        vb6_ComLateBindMiss(L"ComSetPropArg", propName, disp);
    }
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
static void vb6_ComSetRefRcv(void* disp, const wchar_t* propName, void* objRef) {
    /* 账 #305 第二百六十二刀 (丙): 非 IDispatch 接收者 —— 响亮 + 计数 (行为不变) */
    if (vb6_ComLateBindIsMiss(disp)) {
        vb6_ComLateBindMiss(L"ComSetRef", propName, disp);
    }
    if (!disp) return;
    IDispatch* pDisp = (IDispatch*)disp;

    // Fix <vbeclipse>: 宿主对象 (原生 vb6_cls_X* UC 实例 / HWND) 不是真 IDispatch,
    // 走下面的 getDispid+Invoke 必 "not found"。与 vb6_ComSetProp (line 526) 同口径:
    // 先把 objRef 包成 VT_DISPATCH 的 Win VARIANT, 转成 vb6_VARIANT 交给 Host_SetProp,
    // 后者落到 vb6_UC_OwnPropSet 的名字桥 (uc_hostmodel_setprop.inc)。
    /* 账 (Outlook2013 左面板 VisualTheme): 控件窗口 → 控件 IDispatch (只换控件自有成员) */
    { void* ocxRcv = vb6_ComOcxRcvOf(disp, propName); if (ocxRcv) disp = ocxRcv; }
    if (vb6_Host_IsHostObject(disp)) {
        if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
            fprintf(stderr, "[C3_COM] SetRef '%ls' rcv=%p -> HOST\n", propName, disp);
            fflush(stderr);
        }
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

    if (GetEnvironmentVariableW(L"C3_OCX_TRACE", NULL, 0) > 0) {
        fprintf(stderr, "[C3_COM] SetRef '%ls' rcv=%p -> COM dispid=%ld hr=0x%08lX\n",
                propName, disp, (long)dispid, (unsigned long)hr);
        fflush(stderr);
    }

    if (FAILED(hr)) {
        vb6_ComCheckError(hr, &excep, L"ComSetRef");
    }
}

/* ============================================================
 * 账 #305 第三百十七刀: 八个晚绑定出口的**唯一入口** —— 先换到 IDispatch 视图
 * (vb6_ComDispViewIn 的判据与代价见文件头)。原函数体改名 ...Rcv 后原样保留;
 * 这一层只做「进来换视图 / 出去还引用」, 不改任何出口自己的语义。
 * ============================================================ */
/* 账 (frmDocument 最大化/替代路径): vb6forms_mdi.h 的两枚出口; vb6com 与
 * vb6forms 的依赖方向相反, 这里只前置声明, 不改 include 关系。 */
extern void vb6_MDI_NoteTabWorkspaceMissing(void);

void* vb6_ComCall(void* disp, const wchar_t* methodName,
                  void* args_void, int32_t argc) {
    void* owned = NULL;
    void* r = vb6_ComCallRcv(vb6_ComDispViewIn(disp, &owned), methodName, args_void, argc);
    vb6_ComDispViewOut(owned);
    /* 实测 (c3_vc6 / c3_vs2015 两侧): `CommandBars.ShowTabWorkspace(True)` 由 OCX
     * 回 S_OK + 空指针 ⇒ 脚本的 Workspace 恒 Nothing, 活动 MDI 子窗停在设计尺寸
     * (VB6 同调用拿到真对象 ⇒ 子窗铺满文档区)。这里只**打标记**, 由 vb6_ShowForm
     * 按工作区行为把活动子窗最大化 —— 替代路径, 不是把 OCX 那条链补齐。 */
    /* 账 (B4): 返回的是 VARIANT* (VT_DISPATCH + pdispVal=NULL 也算"没给出对象") ——
       实测 CallObject 走的是这一出口, r 本身非空, 只看 `!r` 永远不成立。 */
    if (methodName && wcscmp(methodName, L"ShowTabWorkspace") == 0) {
        int empty = !r;
        if (r) {
            VARIANT* vr = (VARIANT*)r;
            if (vr->vt == VT_EMPTY || vr->vt == VT_NULL
                || (vr->vt == VT_DISPATCH && vr->pdispVal == NULL)
                || (vr->vt == VT_UNKNOWN && vr->punkVal == NULL))
                empty = 1;
        }
        if (empty) vb6_MDI_NoteTabWorkspaceMissing();
    }
    return r;
}

void* vb6_ComCallByDispid(void* disp, int32_t dispid,
                          void* args_void, int32_t argc) {
    void* owned = NULL;
    void* r = vb6_ComCallByDispidRcv(vb6_ComDispViewIn(disp, &owned), dispid, args_void, argc);
    vb6_ComDispViewOut(owned);
    return r;
}

void* vb6_ComCallObjectDefault(void* disp, void* args_void, int32_t argc) {
    void* owned = NULL;
    void* r = vb6_ComCallObjectDefaultRcv(vb6_ComDispViewIn(disp, &owned), args_void, argc);
    vb6_ComDispViewOut(owned);
    return r;
}

void* vb6_ComGetProp(void* disp, const wchar_t* propName) {
    void* owned = NULL;
    void* r = vb6_ComGetPropRcv(vb6_ComDispViewIn(disp, &owned), propName);
    vb6_ComDispViewOut(owned);
    return r;
}

void* vb6_ComGetPropArg(void* disp, const wchar_t* propName,
                        void* args_void, int32_t argc) {
    void* owned = NULL;
    void* r = vb6_ComGetPropArgRcv(vb6_ComDispViewIn(disp, &owned), propName, args_void, argc);
    vb6_ComDispViewOut(owned);
    return r;
}

void vb6_ComSetProp(void* disp, const wchar_t* propName, void* value_void) {
    void* owned = NULL;
    vb6_ComSetPropRcv(vb6_ComDispViewIn(disp, &owned), propName, value_void);
    vb6_ComDispViewOut(owned);
}

void vb6_ComSetPropArg(void* disp, const wchar_t* propName, void** args,
                       int32_t argc, void* value_void) {
    void* owned = NULL;
    vb6_ComSetPropArgRcv(vb6_ComDispViewIn(disp, &owned), propName, args, argc, value_void);
    vb6_ComDispViewOut(owned);
}

void vb6_ComSetRef(void* disp, const wchar_t* propName, void* objRef) {
    void* owned = NULL;
    vb6_ComSetRefRcv(vb6_ComDispViewIn(disp, &owned), propName, objRef);
    vb6_ComDispViewOut(owned);
}
