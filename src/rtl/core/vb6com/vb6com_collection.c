// vb6com_collection.c - VB6 内建对象 Collection 的 C 实现 (IDispatch + IEnumVARIANT)
//
// 背景 (Fix 103): VB6 的 Collection 是 VBA 运行时内建类型, 实现于 MSVBVM60.DLL 内部,
// 并不是注册表里的 COM 组件 —— 没装 VB6 运行时的机器上 CLSIDFromProgID("Collection")
// 与 ("VBA.Collection") 都返回 REGDB_E_CLASSNOTREG。C3 生成的 New Collection 落到
// vb6_NewObject(L"Collection") → vb6_CreateObject → 运行期错误 429
// (实测: ProgID: Collection, HRESULT: 0x800401F3)。
// C3 是把运行时静态链接进用户程序的, 内建对象必须自带实现, 本文件即 Collection 的实现。
//
// 对外只有一个入口 vb6_NewBuiltinObject() —— 命中内建名返回新实例, 未命中返回 NULL
// 让调用方 (vb6rtl_conv.c 的 vb6_NewObject) 回退到 COM 注册表查找。将来再加内建对象
// (Err/App 等) 只需扩展该函数的名字判定。
//
// 与 vb6rtl 族一致, 本文件的接口只出现 void*, 不暴露 vb6_VARIANT, 避免 VARIANT 定义冲突。

#include "vb6com.h"
#include "vb6com_internal.h"

// ============================================================
// DISPID 约定 (GetIDsOfNames 返回什么, Invoke 就按什么分发)
// ============================================================
#define COLL_DISPID_ITEM     0        // 默认成员 DISPID_VALUE, 支持 c(1) / c!key
#define COLL_DISPID_ADD      1
#define COLL_DISPID_REMOVE   2
#define COLL_DISPID_COUNT    3        // 必须为正数: vb6_ForEach_Init 用 dispid>0 判定 Count 可用
#define COLL_DISPID_NEWENUM  (-4)     // DISPID_NEWENUM, For Each 入口

// VB6 错误号 (描述文本复用 vb6com.c 的 vb6_StdErrorDesc 表)
#define COLL_ERR_INVALID_ARG  5       // 无效的过程调用或参数
#define COLL_ERR_SUBSCRIPT    9       // 下标越界
#define COLL_ERR_NO_MEMBER  438       // 对象不支持此属性或方法
#define COLL_ERR_KEY_EXISTS 457       // 此键已经与该集合的一个元素关联

// ============================================================
// 数据结构
// ============================================================
// vb6_CollEntry / vb6_BuiltinCollection / vb6_CollEnum 与枚举器 vtable 声明
// 见 vb6com_internal.h —— 本文件与 vb6com_collection_enum.c 两个编译单元共用。

static IDispatchVtbl   g_collVtbl;

// ============================================================
// 错误上报
// ============================================================
// 与手写 COM 对象的标准做法一致: 填 EXCEPINFO.scode (VB6 错误号的 OLE 形式
// 0x800A0000|err) 并返回 DISP_E_EXCEPTION, 由调用方 vb6_ComCheckError 统一
// 转成 Err.Raise —— 这里不直接 RaiseError, 否则同一错误会被上报两次。
/* 账 #305 §B185: 形状本身搬进 vb6com.c 的 vb6_ComFailExcep —— 这一枚从此只是转调,
 * 「Collection error」那一句缺省描述留在这里 (它是本族的答案, 不是那道决定的答案)。
 * 描述文本照旧先问 vb6_StdErrorDesc 那张表: 内置 Collection 不经过 vb6_RaiseError,
 * Err 里装的是上一层留下的东西, 不许拿来当本条错误的描述。 */
static HRESULT coll_fail(EXCEPINFO* pexcep, int32_t errNum) {
    return vb6_ComFailExcep(pexcep, errNum, vb6_StdErrorDesc(errNum), L"Collection error");
}

// ============================================================
// 参数取值
// ============================================================
// COM 约定: DISPPARAMS.rgvarg 逆序存放实参, 且实参按源码顺序对应形参列表的最左 n 个
// (vb6_ComCall 内部即按 args[i] → rgvarg[argc-1-i] 填充)。故形参 formalIdx 的实参
// 位于 rgvarg[cArgs-1-formalIdx]; formalIdx >= cArgs 表示调用方未提供该形参。
static VARIANT* coll_argRaw(DISPPARAMS* pdp, int32_t formalIdx) {
    if (!pdp || !pdp->rgvarg) return NULL;
    if ((int32_t)pdp->cArgs <= formalIdx) return NULL;
    return &pdp->rgvarg[pdp->cArgs - 1 - formalIdx];
}

// 可选形参版本: 除边界判断外, 把显式缺省占位 (VT_ERROR/DISP_E_PARAMNOTFOUND, 部分
// codegen 路径会生成) 也视为未提供。必填形参不能用本函数 —— Add 的 Item 允许是
// VT_EMPTY, 会被误判成缺省。
static VARIANT* coll_argOpt(DISPPARAMS* pdp, int32_t formalIdx) {
    VARIANT* pv = coll_argRaw(pdp, formalIdx);
    if (!pv) return NULL;
    if (V_VT(pv) == VT_ERROR || V_VT(pv) == VT_EMPTY) return NULL;
    return pv;
}

// ============================================================
// 容器操作
// ============================================================

// 键比较不区分大小写 (与 VB6 Collection 一致)
static int32_t coll_findKey(vb6_BuiltinCollection* c, const wchar_t* key) {
    if (!c || !key || !key[0]) return -1;
    for (int32_t i = 0; i < c->count; i++) {
        if (c->entries[i].key && _wcsicmp(c->entries[i].key, key) == 0) return i;
    }
    return -1;
}

static int32_t coll_ensureCapacity(vb6_BuiltinCollection* c) {
    if (c->count < c->capacity) return 1;
    int32_t ncap = c->capacity > 0 ? c->capacity * 2 : 8;
    vb6_CollEntry* grown = (vb6_CollEntry*)realloc(c->entries, (size_t)ncap * sizeof(vb6_CollEntry));
    if (!grown) return 0;
    c->entries = grown;
    c->capacity = ncap;
    return 1;
}

// 在 0-based 位置 at 插入 (at 会被夹到 [0, count]); key 所有权移入
static int32_t coll_insertAt(vb6_BuiltinCollection* c, int32_t at, BSTR key, VARIANT* val) {
    if (!coll_ensureCapacity(c)) return 0;
    if (at < 0) at = 0;
    if (at > c->count) at = c->count;
    if (at < c->count) {
        memmove(&c->entries[at + 1], &c->entries[at],
                (size_t)(c->count - at) * sizeof(vb6_CollEntry));
    }
    c->entries[at].key = key;
    VariantInit(&c->entries[at].value);
    if (val) VariantCopy(&c->entries[at].value, val);
    c->count++;
    return 1;
}

static void coll_removeAt(vb6_BuiltinCollection* c, int32_t at) {
    if (at < 0 || at >= c->count) return;
    if (c->entries[at].key) {
        SysFreeString(c->entries[at].key);
        c->entries[at].key = NULL;
    }
    VariantClear(&c->entries[at].value);
    if (at + 1 < c->count) {
        memmove(&c->entries[at], &c->entries[at + 1],
                (size_t)(c->count - at - 1) * sizeof(vb6_CollEntry));
    }
    c->count--;
}

// ============================================================
// Index / Before / After 解析
// ============================================================
// Index 可为 1-based 序号或字符串键。字符串键不存在 → 5; 序号越界 → 9。
// 返回 0-based 内部下标, 失败返回 -1 并置 *perr。
static int32_t coll_resolveIndex(vb6_BuiltinCollection* c, VARIANT* pidx, int32_t* perr) {
    *perr = COLL_ERR_INVALID_ARG;
    if (!pidx) return -1;
    if (V_VT(pidx) == VT_BSTR) {
        int32_t i = coll_findKey(c, V_BSTR(pidx));
        if (i < 0) return -1;                  // 键不存在 → 5
        *perr = 0;
        return i;
    }
    VARIANT tmp;
    VariantInit(&tmp);
    if (FAILED(VariantChangeType(&tmp, pidx, 0, VT_I4))) {
        VariantClear(&tmp);
        return -1;                             // 非数值非字符串 → 5
    }
    int32_t n = (int32_t)V_I4(&tmp);
    VariantClear(&tmp);
    if (n < 1 || n > c->count) {               // 序号越界 → 9
        *perr = COLL_ERR_SUBSCRIPT;
        return -1;
    }
    *perr = 0;
    return n - 1;
}

// Before/After 位置解析: 返回插入位置 (0-based, 可为 count 表示追加)。
// 位置非法时 VB6 抛 5 (不是 9) —— 与 Index 越界的错误号不同。
static int32_t coll_resolvePos(vb6_BuiltinCollection* c, VARIANT* ppos, int32_t after, int32_t* perr) {
    *perr = COLL_ERR_INVALID_ARG;
    if (!ppos) return 0;
    if (V_VT(ppos) == VT_BSTR) {
        int32_t i = coll_findKey(c, V_BSTR(ppos));
        if (i < 0) return 0;
        *perr = 0;
        return after ? i + 1 : i;
    }
    VARIANT tmp;
    VariantInit(&tmp);
    if (FAILED(VariantChangeType(&tmp, ppos, 0, VT_I4))) {
        VariantClear(&tmp);
        return 0;
    }
    int32_t n = (int32_t)V_I4(&tmp);
    VariantClear(&tmp);
    if (n < 1 || n > c->count) return 0;
    *perr = 0;
    return after ? n : n - 1;
}

// ============================================================
// 成员实现
// ============================================================

// Add Item, [Key], [Before], [After]
static HRESULT coll_doAdd(vb6_BuiltinCollection* c, DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    VARIANT* pItem = coll_argRaw(pdp, 0);      // 必填
    VARIANT* pKey  = coll_argOpt(pdp, 1);
    VARIANT* pBef  = coll_argOpt(pdp, 2);
    VARIANT* pAft  = coll_argOpt(pdp, 3);
    if (!pItem) return coll_fail(pexcep, COLL_ERR_INVALID_ARG);

    BSTR key = NULL;
    if (pKey) {
        if (V_VT(pKey) != VT_BSTR) return coll_fail(pexcep, COLL_ERR_INVALID_ARG);
        if (V_BSTR(pKey) && V_BSTR(pKey)[0] != L'\0') {   // 空串键按"无键"处理
            if (coll_findKey(c, V_BSTR(pKey)) >= 0) return coll_fail(pexcep, COLL_ERR_KEY_EXISTS);
            key = SysAllocString(V_BSTR(pKey));
        }
    }

    if (pBef && pAft) {                        // Before 与 After 互斥
        if (key) SysFreeString(key);
        return coll_fail(pexcep, COLL_ERR_INVALID_ARG);
    }

    int32_t at  = c->count;
    int32_t perr = 0;
    if (pBef)      at = coll_resolvePos(c, pBef, 0, &perr);
    else if (pAft) at = coll_resolvePos(c, pAft, 1, &perr);
    if (perr) {
        if (key) SysFreeString(key);
        return coll_fail(pexcep, perr);
    }

    if (!coll_insertAt(c, at, key, pItem)) {
        if (key) SysFreeString(key);
        return E_OUTOFMEMORY;
    }
    return S_OK;
}

// Remove Index
static HRESULT coll_doRemove(vb6_BuiltinCollection* c, DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    int32_t perr = 0;
    int32_t at = coll_resolveIndex(c, coll_argRaw(pdp, 0), &perr);
    if (perr) return coll_fail(pexcep, perr);
    coll_removeAt(c, at);
    return S_OK;
}

// Item / 默认成员 — 取值
static HRESULT coll_doItemGet(vb6_BuiltinCollection* c, DISPPARAMS* pdp,
                              VARIANT* pResult, EXCEPINFO* pexcep) {
    int32_t perr = 0;
    int32_t at = coll_resolveIndex(c, coll_argRaw(pdp, 0), &perr);
    if (perr) return coll_fail(pexcep, perr);
    if (!pResult) return S_OK;
    return VariantCopy(pResult, &c->entries[at].value);
}

// Item / 默认成员 — 赋值。注意 put 的参数布局与 get 不同: vb6_ComSetPropArg 把
// 新值放 rgvarg[0], 索引参数从 rgvarg[1] 起 (值在最前是 DISPATCH 的反序约定)。
static HRESULT coll_doItemPut(vb6_BuiltinCollection* c, DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    if (!pdp || pdp->cArgs < 2 || !pdp->rgvarg) return coll_fail(pexcep, COLL_ERR_INVALID_ARG);
    VARIANT* pVal = &pdp->rgvarg[0];
    VARIANT* pIdx = &pdp->rgvarg[1];
    if (V_VT(pIdx) == VT_ERROR || V_VT(pIdx) == VT_EMPTY) return coll_fail(pexcep, COLL_ERR_INVALID_ARG);

    int32_t perr = 0;
    int32_t at = coll_resolveIndex(c, pIdx, &perr);
    if (perr) return coll_fail(pexcep, perr);

    VariantClear(&c->entries[at].value);
    VariantInit(&c->entries[at].value);
    return VariantCopy(&c->entries[at].value, pVal);
}

// _NewEnum → IEnumVARIANT (以 VT_UNKNOWN 回传, 调用方 QI IID_IEnumVARIANT)
static HRESULT coll_doNewEnum(vb6_BuiltinCollection* c, VARIANT* pResult) {
    if (!pResult) return S_OK;
    vb6_CollEnum* pe = (vb6_CollEnum*)calloc(1, sizeof(vb6_CollEnum));
    if (!pe) return E_OUTOFMEMORY;
    pe->lpVtbl   = &g_collEnumVtbl;
    pe->refCount = 1;
    pe->coll     = c;                          // 枚举器持有集合引用
    pe->index    = 0;
    c->lpVtbl->AddRef((IDispatch*)c);

    V_VT(pResult)    = VT_UNKNOWN;
    V_UNKNOWN(pResult) = (IUnknown*)pe;
    return S_OK;
}

// ============================================================
// IDispatch (Collection)
// ============================================================

static HRESULT STDMETHODCALLTYPE coll_QueryInterface(IDispatch* pDisp, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = (void*)pDisp;
        pDisp->lpVtbl->AddRef(pDisp);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE coll_AddRef(IDispatch* pDisp) {
    return (ULONG)InterlockedIncrement(&((vb6_BuiltinCollection*)pDisp)->refCount);
}

static ULONG STDMETHODCALLTYPE coll_Release(IDispatch* pDisp) {
    vb6_BuiltinCollection* c = (vb6_BuiltinCollection*)pDisp;
    LONG left = InterlockedDecrement(&c->refCount);
    if (left == 0) {
        for (int32_t i = 0; i < c->count; i++) {
            if (c->entries[i].key) SysFreeString(c->entries[i].key);
            VariantClear(&c->entries[i].value);
        }
        free(c->entries);
        free(c);
    }
    return (ULONG)left;
}

static HRESULT STDMETHODCALLTYPE coll_GetTypeInfoCount(IDispatch* pDisp, UINT* pctinfo) {
    (void)pDisp;
    if (pctinfo) *pctinfo = 0;                 // 无类型信息 (纯内建对象)
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE coll_GetTypeInfo(IDispatch* pDisp, UINT iTInfo,
                                                  LCID lcid, ITypeInfo** ppTInfo) {
    (void)pDisp; (void)iTInfo; (void)lcid;
    if (ppTInfo) *ppTInfo = NULL;
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE coll_GetIDsOfNames(IDispatch* pDisp, REFIID riid,
                                                    LPOLESTR* rgszNames, UINT cNames,
                                                    LCID lcid, DISPID* rgDispId) {
    (void)pDisp; (void)riid; (void)lcid;
    if (!rgszNames || !rgDispId) return E_POINTER;
    for (UINT i = 0; i < cNames; i++) {
        const wchar_t* n = rgszNames[i];
        if (!n) { rgDispId[i] = DISPID_UNKNOWN; continue; }
        if      (_wcsicmp(n, L"Add")      == 0) rgDispId[i] = COLL_DISPID_ADD;
        else if (_wcsicmp(n, L"Remove")   == 0) rgDispId[i] = COLL_DISPID_REMOVE;
        else if (_wcsicmp(n, L"Item")     == 0) rgDispId[i] = COLL_DISPID_ITEM;
        else if (_wcsicmp(n, L"Count")    == 0) rgDispId[i] = COLL_DISPID_COUNT;
        else if (_wcsicmp(n, L"_NewEnum") == 0) rgDispId[i] = COLL_DISPID_NEWENUM;
        else {
            rgDispId[i] = DISPID_UNKNOWN;
            return DISP_E_UNKNOWNNAME;
        }
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE coll_Invoke(IDispatch* pDisp, DISPID dispIdMember,
                                             REFIID riid, LCID lcid, WORD wFlags,
                                             DISPPARAMS* pdp, VARIANT* pVarResult,
                                             EXCEPINFO* pexcep, UINT* puArgErr) {
    vb6_BuiltinCollection* c = (vb6_BuiltinCollection*)pDisp;
    (void)riid; (void)lcid; (void)puArgErr;

    switch (dispIdMember) {
        case COLL_DISPID_ADD:
            if (!(wFlags & DISPATCH_METHOD)) break;
            return coll_doAdd(c, pdp, pexcep);
        case COLL_DISPID_REMOVE:
            if (!(wFlags & DISPATCH_METHOD)) break;
            return coll_doRemove(c, pdp, pexcep);
        case COLL_DISPID_COUNT:
            if (pVarResult) { V_VT(pVarResult) = VT_I4; V_I4(pVarResult) = (LONG)c->count; }
            return S_OK;
        case COLL_DISPID_NEWENUM:
            return coll_doNewEnum(c, pVarResult);
        case COLL_DISPID_ITEM:
            if (wFlags & DISPATCH_PROPERTYPUT) return coll_doItemPut(c, pdp, pexcep);
            return coll_doItemGet(c, pdp, pVarResult, pexcep);
        default:
            return coll_fail(pexcep, COLL_ERR_NO_MEMBER);
    }
    return coll_fail(pexcep, COLL_ERR_NO_MEMBER);
}

// ============================================================
// vtable 装配与对外入口
// ============================================================

static IDispatchVtbl g_collVtbl = {
    coll_QueryInterface,
    coll_AddRef,
    coll_Release,
    coll_GetTypeInfoCount,
    coll_GetTypeInfo,
    coll_GetIDsOfNames,
    coll_Invoke,
};

// 创建内建 Collection 实例, 返回 IDispatch* (引用计数 1), 失败返回 NULL
void* vb6_NewBuiltinCollection(void) {
    vb6_BuiltinCollection* c = (vb6_BuiltinCollection*)calloc(1, sizeof(vb6_BuiltinCollection));
    if (!c) return NULL;
    c->lpVtbl   = &g_collVtbl;
    c->refCount = 1;
    return (void*)c;
}

// 名字是否指向 C3 运行时自带实现的内建对象。接受源码里的裸名与 VBA. 限定名。
// 不做前缀裁剪以外的匹配 —— 名字撞车的用户类由 cgen 走类工厂另一条路径, 不会到这。
int32_t vb6_IsBuiltinObjectName(const wchar_t* className) {
    if (!className) return 0;
    if (_wcsicmp(className, L"Collection") == 0)     return 1;
    if (_wcsicmp(className, L"VBA.Collection") == 0) return 1;
    return 0;
}

// 内建对象统一构造入口: 命中返回新实例, 未命中返回 NULL (调用方回退到 COM 注册表)。
void* vb6_NewBuiltinObject(const wchar_t* className) {
    if (!vb6_IsBuiltinObjectName(className)) return NULL;
    return vb6_NewBuiltinCollection();
}
