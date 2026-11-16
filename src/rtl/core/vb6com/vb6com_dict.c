// vb6com_dict.c - 内置 Scripting.Dictionary 的 C 实现 (IDispatch + IEnumVARIANT)
//
// 来历：C30-10 第二十六刀。在此之前 CreateObject("Scripting.Dictionary") 与
// New Scripting.Dictionary 都落到真 scrrun.dll，目标机缺 Scripting Runtime 时 429。
// 与 vb6com_collection.c / vb6com_regexp.c 同一条路线：自带实现 + ProgID 拦截。
//
// 判据不是"看起来对"，而是 tests/dict_native/dict_oracle.txt 那 56 行真 scrrun 读数
// （cscript 跑 DictCases.vbs 生成）。键相等的真值表尤其不许按文档猜：
//   Add Empty  -> 等于 Empty / 0 / ""      （三者互为同一个键）
//   Add Null   -> 只等于 Null
//   Add Nothing-> 只等于 Nothing，不等于 0 / ""
//   Add True   -> 等于 True / -1，不等于 1 / "True"
//   Add CDate  -> 等于同一 Date 与它的 CDbl，不等于它的字符串
//   数值类跨类型互通；字符串只与字符串相等（Text 模式大小写不敏感）
//
// 与真组件的一处口径分歧（ai/039 §4.1）：Item 缺键。cscript 实测不抛、返回 Empty；
// 本项目取 **VB6 语义**，抛 32811 —— C3 编的是 VB6 源码，静默 Empty 属于静默错值。

#include "vb6com.h"
#include "vb6com_internal.h"

// ============================================================
// DISPID 约定（GetIDsOfNames 返回什么，Invoke 就按什么分发）
// ============================================================
#define DICT_DISPID_ITEM        0        // 默认成员 DISPID_VALUE，支持 d("k")
#define DICT_DISPID_ADD         1
#define DICT_DISPID_EXISTS      2
#define DICT_DISPID_ITEMS       3
#define DICT_DISPID_KEYS        4
#define DICT_DISPID_REMOVE      5
#define DICT_DISPID_REMOVEALL   6
#define DICT_DISPID_COMPAREMODE 7
#define DICT_DISPID_COUNT       8        // 必须为正数: vb6_ForEach_Init 用 dispid>0 判定 Count 可用
#define DICT_DISPID_NEWENUM     (-4)     // DISPID_NEWENUM

// VB6 错误号（描述文本复用 vb6com.c 的 vb6_StdErrorDesc 表）
#define DICT_ERR_INVALID_ARG  5         // 无效的过程调用或参数
#define DICT_ERR_NO_MEMBER  438         // 对象不支持此属性或方法
#define DICT_ERR_KEY_EXISTS 457         // 此键已经与该集合的一个元素关联
#define DICT_ERR_NOT_FOUND 32811        // 元素不存在（Dictionary 的口径）

// ============================================================
// 数据结构
// ============================================================
typedef struct vb6_DictEntry {
    VARIANT key;
    VARIANT val;
    int32_t live;      // 0 = 墓碑（Remove 后保留槽位，枚举/输出跳过）
    int32_t next;      // 同桶链下一项下标，-1 = 结束
} vb6_DictEntry;

typedef struct vb6_BuiltinDict {
    IDispatchVtbl*  lpVtbl;    // 必须首成员 (COM ABI)
    LONG            refCount;
    int32_t         count;     // 活元素数
    int32_t         used;      // entries 已用槽数（含墓碑）
    int32_t         capacity;
    int32_t         cmpMode;   // 0=Binary 1=Text 2=Database(按 Text 处理) 3=同 Text
    int32_t         nbuckets;
    int32_t*        buckets;   // 桶头下标，-1 = 空
    vb6_DictEntry*  entries;   // 插入序
} vb6_BuiltinDict;

// _NewEnum 返回的枚举器：按 entries 下标前进，只交出活元素的**键**
typedef struct vb6_DictEnum {
    IEnumVARIANTVtbl*     lpVtbl;   // 必须首成员
    LONG                  refCount;
    vb6_BuiltinDict*      dict;
    int32_t               index;
} vb6_DictEnum;

// 枚举器 vtable（定义在本文件末尾）
static IEnumVARIANTVtbl g_dictEnumVtbl;

static IDispatchVtbl g_dictVtbl;

// ============================================================
// 错误上报（形状与 Collection 同族：填 EXCEPINFO.scode，由 vb6_ComCheckError 转 Err）
// ============================================================
static HRESULT dict_fail(EXCEPINFO* pexcep, int32_t errNum) {
    return vb6_ComFailExcep(pexcep, errNum, vb6_StdErrorDesc(errNum), L"Dictionary error");
}

// ============================================================
// 参数取值（COM 约定：rgvarg 逆序，形参 formalIdx 在 rgvarg[cArgs-1-formalIdx]）
// ============================================================
static VARIANT* dict_argRaw(DISPPARAMS* pdp, int32_t formalIdx) {
    if (!pdp || !pdp->rgvarg) return NULL;
    if ((int32_t)pdp->cArgs <= formalIdx) return NULL;
    return &pdp->rgvarg[pdp->cArgs - 1 - formalIdx];
}

// ============================================================
// 键分类 / 取值 / 相等 / 哈希
// ============================================================
enum {
    DK_EMPTY = 0,   // VT_EMPTY
    DK_NULL,        // VT_NULL
    DK_NOTHING,     // VT_DISPATCH/VT_UNKNOWN 且指针为空
    DK_NUM,         // 数值类：整数/浮点/布尔/日期/货币/DECIMAL
    DK_STR,         // VT_BSTR
    DK_OBJ,         // 非空对象 / 数组（按指针同一性比较）
};

static int dict_class(const VARIANT* v) {
    if (!v) return DK_EMPTY;
    switch (V_VT(v)) {
        case VT_EMPTY:  return DK_EMPTY;
        case VT_NULL:   return DK_NULL;
        case VT_DISPATCH:
            return V_DISPATCH(v) ? DK_OBJ : DK_NOTHING;
        case VT_UNKNOWN:
            return V_UNKNOWN(v) ? DK_OBJ : DK_NOTHING;
        case VT_BSTR:   return DK_STR;
        case VT_I1: case VT_I2: case VT_I4: case VT_I8:
        case VT_UI1: case VT_UI2: case VT_UI4: case VT_UI8:
        case VT_INT: case VT_UINT:
        case VT_R4: case VT_R8: case VT_CY: case VT_DATE:
        case VT_BOOL: case VT_DECIMAL:
            return DK_NUM;
        default:
            return DK_OBJ;   // 数组等：只按指针同一性
    }
}

// 数值类统一取值。CY 缩放 10000；BOOL 按 VB6 真值 -1/0；DECIMAL 先转 R8。
static double dict_numVal(const VARIANT* v) {
    switch (V_VT(v)) {
        case VT_I1:   return (double)V_I1(v);
        case VT_I2:   return (double)V_I2(v);
        case VT_I4:   return (double)V_I4(v);
        case VT_I8:   return (double)V_I8(v);
        case VT_UI1:  return (double)V_UI1(v);
        case VT_UI2:  return (double)V_UI2(v);
        case VT_UI4:  return (double)V_UI4(v);
        case VT_UI8:  return (double)V_UI8(v);
        case VT_INT:  return (double)V_INT(v);
        case VT_UINT: return (double)V_UINT(v);
        case VT_R4:   return (double)V_R4(v);
        case VT_R8:   return V_R8(v);
        case VT_CY:   return (double)V_CY(v).int64 / 10000.0;
        case VT_DATE: return (double)V_DATE(v);
        case VT_BOOL: return (V_BOOL(v) != VARIANT_FALSE) ? -1.0 : 0.0;
        case VT_DECIMAL: {
            VARIANT tmp;
            double d = 0.0;
            VariantInit(&tmp);
            if (SUCCEEDED(VariantChangeType(&tmp, v, 0, VT_R8))) d = V_R8(&tmp);
            VariantClear(&tmp);
            return d;
        }
        default: return 0.0;
    }
}

static int dict_strEq(const vb6_BuiltinDict* d, const VARIANT* a, const VARIANT* b) {
    const wchar_t* sa = V_BSTR(a) ? V_BSTR(a) : L"";
    const wchar_t* sb = V_BSTR(b) ? V_BSTR(b) : L"";
    // 实测：CompareMode 0=Binary；1=Text（大小写不敏感）；2=Database 与 1 同档；
    // 3 被接受且判重等同 Text（见 ai/039 §4.0.1）。>0 一律大小写不敏感。
    if (d->cmpMode != 0) return _wcsicmp(sa, sb) == 0;
    return wcscmp(sa, sb) == 0;
}

// 键相等：唯一入口（Add / Exists / Item / Remove 全走这里，避免双写漂移）
static int dict_keyEq(const vb6_BuiltinDict* d, const VARIANT* a, const VARIANT* b) {
    int ca = dict_class(a), cb = dict_class(b);
    // Empty 是"零值"特例：与 Empty / 数值 0 / 空串三者相等（实测 d105）
    if (ca == DK_EMPTY || cb == DK_EMPTY) {
        const VARIANT* other = (ca == DK_EMPTY) ? b : a;
        int co = dict_class(other);
        if (co == DK_EMPTY) return 1;
        if (co == DK_NUM)   return dict_numVal(other) == 0.0;
        if (co == DK_STR)   return V_BSTR(other) == NULL || V_BSTR(other)[0] == L'\0';
        return 0;
    }
    if (ca != cb) return 0;      // 数值与字符串永不相等（实测 d30/d107）
    switch (ca) {
        case DK_NULL:    return 1;
        case DK_NOTHING: return 1;
        case DK_NUM:     return dict_numVal(a) == dict_numVal(b);
        case DK_STR:     return dict_strEq(d, a, b);
        case DK_OBJ: {
            int va = V_VT(a) & VT_TYPEMASK, vb = V_VT(b) & VT_TYPEMASK;
            if (va != vb) return 0;
            if (V_ISARRAY(a) || V_ISARRAY(b)) return V_ARRAY(a) == V_ARRAY(b);
            if (va == VT_DISPATCH) return V_DISPATCH(a) == V_DISPATCH(b);
            return V_UNKNOWN(a) == V_UNKNOWN(b);
        }
        default: return 0;
    }
}

// FNV-1a 32 位
#define DICT_FNV_INIT 2166136261u
#define DICT_FNV_PRIME 16777619u

static uint32_t dict_fnvBytes(uint32_t h, const void* p, size_t n) {
    const unsigned char* b = (const unsigned char*)p;
    for (size_t i = 0; i < n; i++) { h ^= b[i]; h *= DICT_FNV_PRIME; }
    return h;
}

// 哈希必须与相等一致：Empty / 数值 0 / 空串 落到同一个桶（实测三者互等）。
static uint32_t dict_keyHash(const vb6_BuiltinDict* d, const VARIANT* v) {
    int c = dict_class(v);
    if (c == DK_EMPTY) return 0u;
    if (c == DK_NUM) {
        double x = dict_numVal(v);
        if (x == 0.0) return 0u;                 // +0.0 / -0.0 同桶
        return dict_fnvBytes(DICT_FNV_INIT, &x, sizeof(x));
    }
    if (c == DK_STR) {
        const wchar_t* s = V_BSTR(v) ? V_BSTR(v) : L"";
        if (!s[0]) return 0u;
        uint32_t h = DICT_FNV_INIT;
        for (const wchar_t* p = s; *p; p++) {
            wchar_t ch = (d->cmpMode != 0) ? (wchar_t)towlower((wint_t)*p) : *p;
            h = dict_fnvBytes(h, &ch, sizeof(ch));
        }
        return h;
    }
    if (c == DK_NULL)    return 0x4E554C4Cu;     // 'NULL'
    if (c == DK_NOTHING) return 0x4E4F5448u;     // 'NOTH'
    if (V_ISARRAY(v))    return (uint32_t)((uintptr_t)V_ARRAY(v) >> 4);
    if ((V_VT(v) & VT_TYPEMASK) == VT_DISPATCH)
        return (uint32_t)((uintptr_t)V_DISPATCH(v) >> 4);
    return (uint32_t)((uintptr_t)V_UNKNOWN(v) >> 4);
}

// ============================================================
// 桶 / 槽管理
// ============================================================
static int dict_growEntries(vb6_BuiltinDict* d, int32_t need) {
    if (need <= d->capacity) return 1;
    int32_t ncap = d->capacity > 0 ? d->capacity : 8;
    while (ncap < need) ncap *= 2;
    vb6_DictEntry* grown = (vb6_DictEntry*)realloc(d->entries, (size_t)ncap * sizeof(vb6_DictEntry));
    if (!grown) return 0;
    d->entries = grown;
    d->capacity = ncap;
    return 1;
}

static int dict_rehash(vb6_BuiltinDict* d, int32_t nbuckets) {
    if (nbuckets < 16) nbuckets = 16;
    int32_t* nb = (int32_t*)malloc((size_t)nbuckets * sizeof(int32_t));
    if (!nb) return 0;
    for (int32_t i = 0; i < nbuckets; i++) nb[i] = -1;
    for (int32_t i = 0; i < d->used; i++) {
        if (!d->entries[i].live) { d->entries[i].next = -1; continue; }
        uint32_t h = dict_keyHash(d, &d->entries[i].key) & (uint32_t)(nbuckets - 1);
        d->entries[i].next = nb[h];
        nb[h] = i;
    }
    free(d->buckets);
    d->buckets = nb;
    d->nbuckets = nbuckets;
    return 1;
}

static int dict_ensureBuckets(vb6_BuiltinDict* d) {
    if (d->buckets && d->nbuckets > 0) {
        if (d->count * 2 <= d->nbuckets) return 1;
        return dict_rehash(d, d->nbuckets * 2);
    }
    return dict_rehash(d, 16);
}

// 查键：返回 entries 下标（仅活元素），未命中 -1
static int32_t dict_find(vb6_BuiltinDict* d, const VARIANT* key) {
    if (!d->buckets || d->nbuckets <= 0) return -1;
    uint32_t h = dict_keyHash(d, key) & (uint32_t)(d->nbuckets - 1);
    for (int32_t i = d->buckets[h]; i >= 0; i = d->entries[i].next) {
        if (!d->entries[i].live) continue;
        if (dict_keyEq(d, &d->entries[i].key, key)) return i;
    }
    return -1;
}

// 追加一个键（调用方已判重），保持插入序
static int32_t dict_append(vb6_BuiltinDict* d, const VARIANT* key, const VARIANT* val) {
    if (!dict_ensureBuckets(d)) return -1;
    if (!dict_growEntries(d, d->used + 1)) return -1;
    int32_t at = d->used++;
    VariantInit(&d->entries[at].key);
    VariantInit(&d->entries[at].val);
    if (FAILED(VariantCopy(&d->entries[at].key, key))) { d->used--; return -1; }
    if (val && FAILED(VariantCopy(&d->entries[at].val, val))) {
        VariantClear(&d->entries[at].key);
        d->used--;
        return -1;
    }
    d->entries[at].live = 1;
    uint32_t h = dict_keyHash(d, &d->entries[at].key) & (uint32_t)(d->nbuckets - 1);
    d->entries[at].next = d->buckets[h];
    d->buckets[h] = at;
    d->count++;
    return at;
}

// 删除：摘链 + 标墓碑（保留槽位，枚举与 Keys/Items 跳过）
static void dict_kill(vb6_BuiltinDict* d, int32_t at) {
    if (at < 0 || at >= d->used || !d->entries[at].live) return;
    uint32_t h = dict_keyHash(d, &d->entries[at].key) & (uint32_t)(d->nbuckets - 1);
    int32_t prev = -1, cur = d->buckets[h];
    while (cur >= 0 && cur != at) { prev = cur; cur = d->entries[cur].next; }
    if (cur == at) {
        if (prev < 0) d->buckets[h] = d->entries[at].next;
        else          d->entries[prev].next = d->entries[at].next;
    }
    d->entries[at].next = -1;
    d->entries[at].live = 0;
    VariantClear(&d->entries[at].key);
    VariantClear(&d->entries[at].val);
    d->count--;

    // 墓碑过多时压实（保持相对顺序）
    if (d->used > 64 && d->used > d->count * 2) {
        int32_t w = 0;
        for (int32_t i = 0; i < d->used; i++) {
            if (!d->entries[i].live) continue;
            if (w != i) d->entries[w] = d->entries[i];
            w++;
        }
        d->used = w;
        dict_rehash(d, d->nbuckets);
    }
}

// ============================================================
// 数组构造（Keys / Items 的返回值）
// ============================================================
static HRESULT dict_buildArray(vb6_BuiltinDict* d, int wantKeys, VARIANT* pResult) {
    if (!pResult) return S_OK;
    SAFEARRAYBOUND bound;
    bound.lLbound = 0;
    bound.cElements = (ULONG)(d->count > 0 ? d->count : 0);
    SAFEARRAY* sa = SafeArrayCreate(VT_VARIANT, 1, &bound);
    if (!sa) return E_OUTOFMEMORY;
    LONG idx = 0;
    for (int32_t i = 0; i < d->used; i++) {
        if (!d->entries[i].live) continue;
        VARIANT* slot = NULL;
        if (FAILED(SafeArrayPtrOfIndex(sa, &idx, (void**)&slot))) { idx++; continue; }
        VariantInit(slot);
        VariantCopy(slot, wantKeys ? &d->entries[i].key : &d->entries[i].val);
        idx++;
    }
    V_VT(pResult) = VT_ARRAY | VT_VARIANT;
    V_ARRAY(pResult) = sa;
    return S_OK;
}

// ============================================================
// 成员实现
// ============================================================
static HRESULT dict_doAdd(vb6_BuiltinDict* d, DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    VARIANT* pKey = dict_argRaw(pdp, 0);
    VARIANT* pVal = dict_argRaw(pdp, 1);
    if (!pKey) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
    if (dict_find(d, pKey) >= 0) return dict_fail(pexcep, DICT_ERR_KEY_EXISTS);
    if (dict_append(d, pKey, pVal) < 0) return E_OUTOFMEMORY;
    return S_OK;
}

static HRESULT dict_doExists(vb6_BuiltinDict* d, DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    VARIANT* pKey = dict_argRaw(pdp, 0);
    if (!pKey) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
    if (pResult) {
        V_VT(pResult) = VT_BOOL;
        V_BOOL(pResult) = (dict_find(d, pKey) >= 0) ? VARIANT_TRUE : VARIANT_FALSE;
    }
    return S_OK;
}

// Item / 默认成员 —— 取值。缺键按 VB6 语义抛 32811（ai/039 §4.1）
static HRESULT dict_doItemGet(vb6_BuiltinDict* d, DISPPARAMS* pdp, VARIANT* pResult, EXCEPINFO* pexcep) {
    VARIANT* pKey = dict_argRaw(pdp, 0);
    if (!pKey) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
    int32_t at = dict_find(d, pKey);
    if (at < 0) return dict_fail(pexcep, DICT_ERR_NOT_FOUND);
    if (!pResult) return S_OK;
    return VariantCopy(pResult, &d->entries[at].val);
}

// Item / 默认成员 —— 赋值：已有键只改值（位置不变），缺键则追加（实测 d70/d109）
static HRESULT dict_doItemPut(vb6_BuiltinDict* d, DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    if (!pdp || pdp->cArgs < 2 || !pdp->rgvarg) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
    VARIANT* pVal = &pdp->rgvarg[0];              // 值在最前（DISPATCH 反序约定）
    VARIANT* pKey = &pdp->rgvarg[1];
    int32_t at = dict_find(d, pKey);
    if (at < 0) {
        if (dict_append(d, pKey, pVal) < 0) return E_OUTOFMEMORY;
        return S_OK;
    }
    VariantClear(&d->entries[at].val);
    VariantInit(&d->entries[at].val);
    return VariantCopy(&d->entries[at].val, pVal);
}

static HRESULT dict_doRemove(vb6_BuiltinDict* d, DISPPARAMS* pdp, EXCEPINFO* pexcep) {
    VARIANT* pKey = dict_argRaw(pdp, 0);
    if (!pKey) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
    int32_t at = dict_find(d, pKey);
    if (at < 0) return dict_fail(pexcep, DICT_ERR_NOT_FOUND);
    dict_kill(d, at);
    return S_OK;
}

static HRESULT dict_doRemoveAll(vb6_BuiltinDict* d) {
    for (int32_t i = 0; i < d->used; i++) {
        VariantClear(&d->entries[i].key);
        VariantClear(&d->entries[i].val);
        d->entries[i].live = 0;
        d->entries[i].next = -1;
    }
    d->used = 0;
    d->count = 0;
    if (d->buckets) {
        for (int32_t i = 0; i < d->nbuckets; i++) d->buckets[i] = -1;
    }
    return S_OK;
}

static HRESULT dict_doNewEnum(vb6_BuiltinDict* d, VARIANT* pResult) {
    if (!pResult) return S_OK;
    vb6_DictEnum* pe = (vb6_DictEnum*)calloc(1, sizeof(vb6_DictEnum));
    if (!pe) return E_OUTOFMEMORY;
    pe->lpVtbl = &g_dictEnumVtbl;
    pe->refCount = 1;
    pe->dict = d;
    pe->index = 0;
    d->lpVtbl->AddRef((IDispatch*)d);
    V_VT(pResult) = VT_UNKNOWN;
    V_UNKNOWN(pResult) = (IUnknown*)pe;
    return S_OK;
}

// ============================================================
// IDispatch
// ============================================================
static HRESULT STDMETHODCALLTYPE dict_QueryInterface(IDispatch* pDisp, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = (void*)pDisp;
        pDisp->lpVtbl->AddRef(pDisp);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE dict_AddRef(IDispatch* pDisp) {
    return (ULONG)InterlockedIncrement(&((vb6_BuiltinDict*)pDisp)->refCount);
}

static ULONG STDMETHODCALLTYPE dict_Release(IDispatch* pDisp) {
    vb6_BuiltinDict* d = (vb6_BuiltinDict*)pDisp;
    LONG left = InterlockedDecrement(&d->refCount);
    if (left == 0) {
        for (int32_t i = 0; i < d->used; i++) {
            VariantClear(&d->entries[i].key);
            VariantClear(&d->entries[i].val);
        }
        free(d->entries);
        free(d->buckets);
        free(d);
    }
    return (ULONG)left;
}

static HRESULT STDMETHODCALLTYPE dict_GetTypeInfoCount(IDispatch* pDisp, UINT* pctinfo) {
    (void)pDisp;
    if (pctinfo) *pctinfo = 0;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE dict_GetTypeInfo(IDispatch* pDisp, UINT iTInfo,
                                                  LCID lcid, ITypeInfo** ppTInfo) {
    (void)pDisp; (void)iTInfo; (void)lcid;
    if (ppTInfo) *ppTInfo = NULL;
    return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE dict_GetIDsOfNames(IDispatch* pDisp, REFIID riid,
                                                    LPOLESTR* rgszNames, UINT cNames,
                                                    LCID lcid, DISPID* rgDispId) {
    (void)pDisp; (void)riid; (void)lcid;
    if (!rgszNames || !rgDispId) return E_POINTER;
    for (UINT i = 0; i < cNames; i++) {
        const wchar_t* n = rgszNames[i];
        if (!n) { rgDispId[i] = DISPID_UNKNOWN; continue; }
        if      (_wcsicmp(n, L"Item")        == 0) rgDispId[i] = DICT_DISPID_ITEM;
        else if (_wcsicmp(n, L"Add")         == 0) rgDispId[i] = DICT_DISPID_ADD;
        else if (_wcsicmp(n, L"Exists")      == 0) rgDispId[i] = DICT_DISPID_EXISTS;
        else if (_wcsicmp(n, L"Items")       == 0) rgDispId[i] = DICT_DISPID_ITEMS;
        else if (_wcsicmp(n, L"Keys")        == 0) rgDispId[i] = DICT_DISPID_KEYS;
        else if (_wcsicmp(n, L"Remove")      == 0) rgDispId[i] = DICT_DISPID_REMOVE;
        else if (_wcsicmp(n, L"RemoveAll")   == 0) rgDispId[i] = DICT_DISPID_REMOVEALL;
        else if (_wcsicmp(n, L"CompareMode") == 0) rgDispId[i] = DICT_DISPID_COMPAREMODE;
        else if (_wcsicmp(n, L"Count")       == 0) rgDispId[i] = DICT_DISPID_COUNT;
        else if (_wcsicmp(n, L"_NewEnum")    == 0) rgDispId[i] = DICT_DISPID_NEWENUM;
        else {
            rgDispId[i] = DISPID_UNKNOWN;
            return DISP_E_UNKNOWNNAME;
        }
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE dict_Invoke(IDispatch* pDisp, DISPID dispIdMember,
                                             REFIID riid, LCID lcid, WORD wFlags,
                                             DISPPARAMS* pdp, VARIANT* pVarResult,
                                             EXCEPINFO* pexcep, UINT* puArgErr) {
    vb6_BuiltinDict* d = (vb6_BuiltinDict*)pDisp;
    (void)riid; (void)lcid; (void)puArgErr;

    switch (dispIdMember) {
        case DICT_DISPID_ADD:
            if (!(wFlags & DISPATCH_METHOD)) break;
            return dict_doAdd(d, pdp, pexcep);
        case DICT_DISPID_EXISTS:
            if (!(wFlags & DISPATCH_METHOD)) break;
            return dict_doExists(d, pdp, pVarResult, pexcep);
        case DICT_DISPID_ITEMS:
            if (wFlags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET)) {
                return dict_buildArray(d, 0, pVarResult);
            }
            break;
        case DICT_DISPID_KEYS:
            if (wFlags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET)) {
                return dict_buildArray(d, 1, pVarResult);
            }
            break;
        case DICT_DISPID_REMOVE:
            if (!(wFlags & DISPATCH_METHOD)) break;
            return dict_doRemove(d, pdp, pexcep);
        case DICT_DISPID_REMOVEALL:
            if (wFlags & (DISPATCH_METHOD | DISPATCH_PROPERTYGET)) {
                return dict_doRemoveAll(d);
            }
            break;
        case DICT_DISPID_COMPAREMODE:
            if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
                VARIANT* pv = (pdp && pdp->rgvarg) ? &pdp->rgvarg[0] : NULL;
                if (!pv) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
                long want = 0;
                if (VT_I2 == V_VT(pv))      want = V_I2(pv);
                else if (VT_I4 == V_VT(pv)) want = V_I4(pv);
                else if (VT_UI1 == V_VT(pv)) want = V_UI1(pv);
                else if (VT_R8 == V_VT(pv)) want = (long)V_R8(pv);
                else return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
                // 实测：负数报 5；0..3 接受（3 判重等同 Text）；有元素时改报 5
                if (want < 0) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
                if (d->count > 0) return dict_fail(pexcep, DICT_ERR_INVALID_ARG);
                if (want > 3) want = 3;
                d->cmpMode = (int32_t)want;
                return S_OK;
            }
            if (pVarResult) {
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = (LONG)d->cmpMode;
            }
            return S_OK;
        case DICT_DISPID_COUNT:
            if (pVarResult) {
                V_VT(pVarResult) = VT_I4;
                V_I4(pVarResult) = (LONG)d->count;
            }
            return S_OK;
        case DICT_DISPID_NEWENUM:
            return dict_doNewEnum(d, pVarResult);
        case DICT_DISPID_ITEM:
            if (wFlags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF))
                return dict_doItemPut(d, pdp, pexcep);
            return dict_doItemGet(d, pdp, pVarResult, pexcep);
        default:
            return dict_fail(pexcep, DICT_ERR_NO_MEMBER);
    }
    return dict_fail(pexcep, DICT_ERR_NO_MEMBER);
}

static IDispatchVtbl g_dictVtbl = {
    dict_QueryInterface,
    dict_AddRef,
    dict_Release,
    dict_GetTypeInfoCount,
    dict_GetTypeInfo,
    dict_GetIDsOfNames,
    dict_Invoke,
};

// ============================================================
// IEnumVARIANT（For Each 的入口；枚举的是**键**）
// ============================================================
static HRESULT STDMETHODCALLTYPE de_QueryInterface(IEnumVARIANT* pEnum, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IEnumVARIANT)) {
        *ppv = (void*)pEnum;
        pEnum->lpVtbl->AddRef(pEnum);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE de_AddRef(IEnumVARIANT* pEnum) {
    return (ULONG)InterlockedIncrement(&((vb6_DictEnum*)pEnum)->refCount);
}

static ULONG STDMETHODCALLTYPE de_Release(IEnumVARIANT* pEnum) {
    vb6_DictEnum* pe = (vb6_DictEnum*)pEnum;
    LONG left = InterlockedDecrement(&pe->refCount);
    if (left == 0) {
        if (pe->dict) pe->dict->lpVtbl->Release((IDispatch*)pe->dict);
        free(pe);
    }
    return (ULONG)left;
}

// 下一个活元素的键；到底返回 S_FALSE（VB6 的 For Each 据此收尾）
static HRESULT STDMETHODCALLTYPE de_Next(IEnumVARIANT* pEnum, ULONG celt,
                                         VARIANT* rgVar, ULONG* pCeltFetched) {
    vb6_DictEnum* pe = (vb6_DictEnum*)pEnum;
    ULONG got = 0;
    if (!pe || !pe->dict || !rgVar) return E_POINTER;
    while (got < celt && pe->index < pe->dict->used) {
        int32_t i = pe->index++;
        if (!pe->dict->entries[i].live) continue;
        VariantInit(&rgVar[got]);
        VariantCopy(&rgVar[got], &pe->dict->entries[i].key);
        got++;
    }
    if (pCeltFetched) *pCeltFetched = got;
    return (got == celt) ? S_OK : S_FALSE;
}

static HRESULT STDMETHODCALLTYPE de_Skip(IEnumVARIANT* pEnum, ULONG celt) {
    vb6_DictEnum* pe = (vb6_DictEnum*)pEnum;
    ULONG skipped = 0;
    if (!pe || !pe->dict) return E_POINTER;
    while (skipped < celt && pe->index < pe->dict->used) {
        if (pe->dict->entries[pe->index].live) skipped++;
        pe->index++;
    }
    return (skipped == celt) ? S_OK : S_FALSE;
}

static HRESULT STDMETHODCALLTYPE de_Reset(IEnumVARIANT* pEnum) {
    vb6_DictEnum* pe = (vb6_DictEnum*)pEnum;
    if (!pe) return E_POINTER;
    pe->index = 0;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE de_Clone(IEnumVARIANT* pEnum, IEnumVARIANT** ppEnum) {
    vb6_DictEnum* pe = (vb6_DictEnum*)pEnum;
    vb6_DictEnum* pn = NULL;
    if (!ppEnum) return E_POINTER;
    *ppEnum = NULL;
    if (!pe || !pe->dict) return E_POINTER;
    pn = (vb6_DictEnum*)calloc(1, sizeof(vb6_DictEnum));
    if (!pn) return E_OUTOFMEMORY;
    pn->lpVtbl = &g_dictEnumVtbl;
    pn->refCount = 1;
    pn->dict = pe->dict;
    pn->index = pe->index;
    pe->dict->lpVtbl->AddRef((IDispatch*)pe->dict);
    *ppEnum = (IEnumVARIANT*)pn;
    return S_OK;
}

static IEnumVARIANTVtbl g_dictEnumVtbl = {
    de_QueryInterface,
    de_AddRef,
    de_Release,
    de_Next,
    de_Skip,
    de_Reset,
    de_Clone,
};

// ============================================================
// 对外入口
// ============================================================
// 是否 Scripting.Dictionary（按 ProgID 认；带版本后缀 .1 也认）
int32_t vb6_Dict_IsProgramId(const wchar_t* progId) {
    if (!progId || !*progId) return 0;
    if (_wcsicmp(progId, L"Scripting.Dictionary") == 0) return 1;
    if (_wcsicmp(progId, L"Scripting.Dictionary.1") == 0) return 1;
    return 0;
}

// 新建实例（引用计数 1）。失败返回 NULL 让调用方回退到注册表。
void* vb6_Dict_New(void) {
    vb6_BuiltinDict* d = (vb6_BuiltinDict*)calloc(1, sizeof(vb6_BuiltinDict));
    if (!d) return NULL;
    d->lpVtbl = &g_dictVtbl;
    d->refCount = 1;
    d->cmpMode = 0;
    if (!dict_rehash(d, 16)) { free(d); return NULL; }
    // TypeName 的回答走账 #611 那张 ProgID 表：登记成 "Scripting.Dictionary"，
    // 于是 TypeName 取最后一段答 "Dictionary"（与真组件同源）。
    vb6_ComRememberProgID(d, L"Scripting.Dictionary");
    return (void*)d;
}

// 名字是否指向内置 Dictionary（New 那两条出口经 vb6_NewBuiltinObject 到这里）
int32_t vb6_Dict_IsClassName(const wchar_t* className) {
    if (!className || !*className) return 0;
    if (_wcsicmp(className, L"Dictionary") == 0) return 1;
    if (_wcsicmp(className, L"Scripting.Dictionary") == 0) return 1;
    return 0;
}

// C30-10 第二十八刀: 供 vb6_ComCreateTyped 判定「这枚指针是不是内置 Dictionary 实例」。
// 声明成类名的槽位 (Dim d As Dictionary) 生成的是 vb6_ComCreateTyped(progId, iidStr),
// 它会拿 scrrun 的 typelib IID 去 QI —— 内置实现没有那层接口, QI 必失败, 槽位就变 NULL。
// 认亲只认首字段 vtable 指针: 不额外加字段、不改对象布局, 也不会误认别人的对象。
int vb6_Dict_IsInstance(void* obj) {
    if (!obj) return 0;
    return *(void***)obj == (void**)&g_dictVtbl;
}