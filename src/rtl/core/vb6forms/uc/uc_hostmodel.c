// uc_hostmodel.c - vb6forms_uc 拆分片：窗体/控件宿主对象模型（注册/判定/分派）
//
// 内容 = 拆分前 vb6forms_uc.c 第 755~819 / 1187~1300 / 1303~1303 / 1391~1391 / 1423~1423 / 1497~1542 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"
#include "vb6forms_prop_ctrl.h"   // Fix 143: vb6_AddItem/RemoveItem/ClearList (原生列表 COM 晚绑定)

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 窗体/控件登记 (供宿主对象分派)
// ============================================================

void vb6_HostObj_Register(void* hwnd, const char* name, const char* vbTypeName,
                          int32_t isForm, int32_t index) {
    if (!hwnd) return;
    if (vb6_ho_find(hwnd)) return;
    if (g_hoCount >= VB6_UC_MAX_OBJ) return;
    vb6_HostObjRec* h = &g_ho[g_hoCount++];
    memset(h, 0, sizeof(*h));
    h->hwnd = hwnd;
    h->isForm = isForm;
    h->index = index;
    if (name) {
        size_t n = strlen(name);
        if (n >= VB6_UC_NAME_LEN) n = VB6_UC_NAME_LEN - 1;
        for (size_t i = 0; i < n; i++) h->name[i] = (wchar_t)name[i];
    }
    if (vbTypeName) {
        const char* dot = strrchr(vbTypeName, '.');
        if (dot) vbTypeName = dot + 1;
        size_t n = strlen(vbTypeName);
        if (n >= VB6_UC_NAME_LEN) n = VB6_UC_NAME_LEN - 1;
        for (size_t i = 0; i < n; i++) h->typeName[i] = (wchar_t)vbTypeName[i];
    }
}

// Fix 148: 宿主对象登记表只读访问器 —— 供容器/控件对象模型使用 (axcontainer.c 调用)
int32_t vb6_HostObj_GetIndex(void* hwnd) {
    vb6_HostObjRec* r = vb6_ho_find(hwnd);
    return r ? r->index : -1;
}
const wchar_t* vb6_HostObj_GetTypeName(void* hwnd) {
    vb6_HostObjRec* r = vb6_ho_find(hwnd);
    return r ? r->typeName : L"";
}
const wchar_t* vb6_HostObj_GetName(void* hwnd) {
    vb6_HostObjRec* r = vb6_ho_find(hwnd);
    return r ? r->name : L"";
}
int32_t vb6_HostObj_Count(void) { return g_hoCount; }
void* vb6_HostObj_At(int32_t i) {
    if (i < 0 || i >= g_hoCount) return NULL;
    return g_ho[i].hwnd;
}

vb6_HostObjRec* vb6_ho_findWindow(const void* hwnd) {
    return vb6_ho_find(hwnd);
}

typedef struct Vb6WrapPair { void* wrap; void* target; } Vb6WrapPair;
static Vb6WrapPair* g_uc_wraps = NULL;
static int g_uc_wrapCount = 0, g_uc_wrapCap = 0;

int32_t vb6_Host_IsHostObject(void* obj) {
    if (!obj) return 0;
    {   // czUI fix: 包装器视为宿主对象 (含内部透明解包)
        for (int i = 0; i < g_uc_wrapCount; i++)
            if (g_uc_wraps[i].wrap == obj) return 1;
    }
    // 真实窗口 (窗体/标准控件) HWND: IsWindow 仅查句柄表, 不解引用, 对任意指针安全
    if (IsWindow((HWND)obj)) return 1;
    if (vb6_uc_isControls(obj)) return 1;
    if (vb6_uc_isFont(obj)) return 1;
    if (vb6_uc_isColl(obj)) return 1;   // Fix 112c: RTL 内建 Collection
    if (vb6_uc_findByHwnd(obj)) return 1;
    // Fix <vbeclipse> rev18: UserControl 的**实例指针**也算宿主对象。
    // 晚绑定调用点 (`Dim x As Variant: Set x = Controls.Item(id)` 之后 `x.Move` /
    // `x.LeftPos`) 传进来的是原生 `vb6_cls_X*`, 不是 HWND 也不是包装器; 少了这一条
    // vb6_ComCall/ComGetProp 会落进 IDispatch 路径 (把实例首字段 `__comObj` 当 lpVtbl,
    // 经包装器再问一次名字) ⇒ 宿主模型永远接不到, Move / LeftPos 全部落空 ——
    // 正是 7 个 ucFolder 宿主停在设计期 569x441 重叠的原因。
    // 判据用 vb6_uc_findByInstance: 只比指针, 对任意指针安全。
    if (vb6_uc_findByInstance(obj)) return 1;
    return vb6_ho_find(obj) != NULL;
}

const wchar_t* vb6_Host_TypeNameOf(void* obj) {
    if (vb6_uc_isControls(obj)) return L"Collection";
    if (vb6_uc_isFont(obj)) return L"Font";
    if (vb6_uc_isColl(obj)) return L"Collection";   // Fix 112c
    vb6_HostObjRec* h = vb6_ho_find(obj);
    if (h) {
        if (h->typeName[0]) return h->typeName;
        return h->isForm ? L"Form" : L"Control";
    }
    vb6_UCRec* r = vb6_uc_findByHwnd(obj);
    if (r && r->desc && r->desc->typeName) {
        const char* dot = strrchr(r->desc->typeName, '.');
        static wchar_t buf[VB6_UC_NAME_LEN];
        const char* p = dot ? dot + 1 : r->desc->typeName;
        size_t i = 0;
        for (; p[i] && i < VB6_UC_NAME_LEN - 1; i++) buf[i] = (wchar_t)(unsigned char)p[i];
        buf[i] = 0;
        return buf;
    }
    if (IsWindow((HWND)obj)) return L"Control";
    return NULL;
}

// ============================================================
// 宿主对象属性/方法分派
// ============================================================

static int32_t vb6_ho_isForm(const void* hwnd) {
    vb6_HostObjRec* h = vb6_ho_find(hwnd);
    return (h && h->isForm) ? 1 : 0;
}

static int32_t vb6_ho_isControl(const void* hwnd) {
    vb6_HostObjRec* h = vb6_ho_find(hwnd);
    return (h && !h->isForm) ? 1 : 0;
}

void vb6_ho_setVariantLong(vb6_VARIANT* out, int32_t v) {
    memset(out, 0, sizeof(*out));
    out->vt = vb6_vtLong;   // 由调用方按 VT_I4 解释
    out->lVal = v;
}

void vb6_ho_setVariantBstr(vb6_VARIANT* out, BSTR s) {
    memset(out, 0, sizeof(*out));
    out->vt = vb6_vtBSTR;
    out->bstrVal = s;
}

void vb6_ho_setVariantDispatch(vb6_VARIANT* out, void* p) {
    // 注意: 宿主合成的"对象"(Controls 集合 / Font 代理 / 控件 HWND)不是真 COM 对象,
    // 调用方 vb6_ComVarClear → VariantClear 会对 VT_DISPATCH 调 Release → 崩溃.
    // 因此这里一律返回 Empty, 让上层走 Nothing 分支 (等价 VB6 On Error Resume Next).
    (void)p;
    memset(out, 0, sizeof(*out));
    out->vt = vb6_vtEmpty;
}

// Fix <vbeclipse> rev14: 与上面 setVariantDispatch 相反 —— **真的**产出
// VT_DISPATCH, 供 Controls.Item / Controls.Add 返回工程内 UserControl 对象。
// 为什么可以: vb6_ReleaseObject / vb6_ComAddRefDispatch 都用 vb6_ComIsDispatchable
// 判据 (首字段是否是可执行的 vtable), 裸实例/HWND 一律跳过, 不会越权调用 Release;
// 而 vb6_ComCallObject 走 vb6_ComVarFree (只 free 结构体, 不 Release)。
// 单列一个函数是为了不动 setVariantDispatch 的既有 Empty 语义 (Font 代理等
// 依赖它继续走 Nothing 分支, 避免回归)。
void vb6_ho_setVariantObject(vb6_VARIANT* out, void* p) {
    memset(out, 0, sizeof(*out));
    if (!p) { out->vt = vb6_vtEmpty; return; }
    out->vt = vb6_vtDispatch;
    out->pdispVal = p;
}

void vb6_ho_setVariantEmpty(vb6_VARIANT* out) {
    memset(out, 0, sizeof(*out));
    out->vt = vb6_vtEmpty;
}

void vb6_ho_setVariantDouble(vb6_VARIANT* out, double v) {
    memset(out, 0, sizeof(*out));
    out->vt = vb6_vtDouble;
    out->dblVal = v;
}

int32_t vb6_ho_variantToLong(const vb6_VARIANT* v) {
    if (!v) return 0;
    switch (v->vt) {
        case vb6_vtInteger: case vb6_vtLong: case vb6_vtBoolean: case vb6_vtByte:
            return v->lVal;
        case vb6_vtSingle: case vb6_vtDouble: return (int32_t)v->dblVal;
        default: return 0;
    }
}

double vb6_ho_variantToDouble(const vb6_VARIANT* v) {
    if (!v) return 0.0;
    switch (v->vt) {
        case vb6_vtSingle: case vb6_vtDouble: return v->dblVal;
        case vb6_vtInteger: case vb6_vtLong: case vb6_vtBoolean: return (double)v->lVal;
        default: return 0.0;
    }
}

BSTR vb6_ho_variantToBstr(const vb6_VARIANT* v) {
    return (v && v->vt == vb6_vtBSTR) ? v->bstrVal : NULL;
}

// 控件的几何: 相对父窗口客户区, 缇
static void vb6_ho_ctrlRect(void* hwnd, int32_t* l, int32_t* t, int32_t* w, int32_t* h) {
    RECT rc; GetWindowRect((HWND)hwnd, &rc);
    POINT pt = { rc.left, rc.top };
    HWND p = GetParent((HWND)hwnd);
    ScreenToClient(p ? p : hwnd, &pt);
    if (l) *l = vb6_XToTwipX(pt.x);
    if (t) *t = vb6_YToTwipY(pt.y);
    if (w) *w = vb6_XToTwipX(rc.right - rc.left);
    if (h) *h = vb6_YToTwipY(rc.bottom - rc.top);
}

static void vb6_ho_clientTwips(void* hwnd, int32_t* w, int32_t* h) {
    RECT rc; GetClientRect((HWND)hwnd, &rc);
    if (w) *w = vb6_XToTwipX(rc.right - rc.left);
    if (h) *h = vb6_YToTwipY(rc.bottom - rc.top);
}

static int32_t vb6_ho_getFontMember(void* fontProxy, const wchar_t* name, vb6_VARIANT* out) {
    vb6_ComIface_Font* f = vb6_uc_fontOf(fontProxy);
    if (!f) return 0;
    if (_wcsicmp(name, L"Size") == 0)      { vb6_ho_setVariantDouble(out, f->Size); return 1; }
    if (_wcsicmp(name, L"Bold") == 0)      { vb6_ho_setVariantLong(out, f->Bold ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Italic") == 0)    { vb6_ho_setVariantLong(out, f->Italic ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Underline") == 0) { vb6_ho_setVariantLong(out, f->Underline ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Strikethrough") == 0) { vb6_ho_setVariantLong(out, f->Strikethrough ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Weight") == 0)    { vb6_ho_setVariantLong(out, f->Weight); return 1; }
    if (_wcsicmp(name, L"Charset") == 0)   { vb6_ho_setVariantLong(out, f->Charset); return 1; }
    if (_wcsicmp(name, L"Name") == 0)      { vb6_ho_setVariantBstr(out, SysAllocString(f->Name)); return 1; }
    return 0;
}

static int32_t vb6_ho_putFontMember(void* fontProxy, const wchar_t* name, const vb6_VARIANT* v) {
    vb6_ComIface_Font* f = vb6_uc_fontOf(fontProxy);
    if (!f) return 0;
    if (_wcsicmp(name, L"Size") == 0)      { f->Size = (float)vb6_ho_variantToDouble(v); return 1; }
    if (_wcsicmp(name, L"Bold") == 0)      { f->Bold = (int16_t)(vb6_ho_variantToLong(v) ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Italic") == 0)    { f->Italic = (int16_t)(vb6_ho_variantToLong(v) ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Underline") == 0) { f->Underline = (int16_t)(vb6_ho_variantToLong(v) ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Strikethrough") == 0) { f->Strikethrough = (int16_t)(vb6_ho_variantToLong(v) ? -1 : 0); return 1; }
    if (_wcsicmp(name, L"Weight") == 0)    { f->Weight = vb6_ho_variantToLong(v); return 1; }
    if (_wcsicmp(name, L"Charset") == 0)   { f->Charset = vb6_ho_variantToLong(v); return 1; }
    if (_wcsicmp(name, L"Name") == 0)      { f->Name = vb6_ho_variantToBstr(v); return 1; }
    return 0;
}

int32_t vb6_Host_GetProp(void* obj, const wchar_t* name, void* outV) {
#include "uc_hostmodel_getprop.inc"

}

int32_t vb6_Host_SetProp(void* obj, const wchar_t* name, const void* inV) {
#include "uc_hostmodel_setprop.inc"

}

int32_t vb6_Host_Call(void* obj, const wchar_t* name, int32_t argc, void** argv, void* outV) {
#include "uc_hostmodel_call.inc"

}

// ============================================================
// Windows VARIANT ↔ vb6_VARIANT (供 vb6_com_invoke/wrap 挂接点使用)
// ============================================================

void vb6_Host_ToWinVariant(const void* inV, void* outV) {
    const vb6_VARIANT* in = (const vb6_VARIANT*)inV;
    VARIANT* out = (VARIANT*)outV;
    VariantInit(out);
    if (!in) return;
    switch (in->vt) {
        case vb6_vtInteger:  V_VT(out) = VT_I2; V_I2(out) = (short)in->iVal; break;
        case vb6_vtLong:     V_VT(out) = VT_I4; V_I4(out) = in->lVal; break;
        case vb6_vtBoolean:  V_VT(out) = VT_BOOL; V_BOOL(out) = in->boolVal; break;
        case vb6_vtByte:     V_VT(out) = VT_UI1; V_UI1(out) = in->bVal; break;
        case vb6_vtSingle:   V_VT(out) = VT_R4; V_R4(out) = in->fltVal; break;
        case vb6_vtDouble:   V_VT(out) = VT_R8; V_R8(out) = in->dblVal; break;
        case vb6_vtBSTR:     V_VT(out) = VT_BSTR; V_BSTR(out) = SysAllocString(in->bstrVal); break;
        // Fix <vbeclipse> rev14: 缺这一条时 `.Item(...)` 返回的对象变体落 default →
        // VT_EMPTY, vb6_ComCallObject 取不到 pdispVal → `Set x = Controls.Item(...)`
        // 恒为 NULL。宿主对象 (UC 实例/HWND/集合) 不是真 IDispatch, 这里只透传指针;
        // 释放侧由 vb6_ComIsDispatchable 守卫, 不会对裸指针调 Release。
        case vb6_vtDispatch: V_VT(out) = VT_DISPATCH; V_DISPATCH(out) = (IDispatch*)in->pdispVal; break;
        case vb6_vtEmpty:    V_VT(out) = VT_EMPTY; break;
        case vb6_vtNull:     V_VT(out) = VT_NULL; break;
        default:             V_VT(out) = VT_EMPTY; break;
    }
}

void vb6_Host_FromWinVariant(const void* inV, void* outV) {
    const VARIANT* in = (const VARIANT*)inV;
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    memset(out, 0, sizeof(*out));
    if (!in) { out->vt = vb6_vtEmpty; return; }
    switch (in->vt) {
        case VT_I2:    out->vt = vb6_vtInteger; out->iVal = in->iVal; break;
        case VT_I4:    out->vt = vb6_vtLong; out->lVal = in->lVal; break;
        case VT_BOOL:  out->vt = vb6_vtBoolean; out->boolVal = in->boolVal; break;
        case VT_UI1:   out->vt = vb6_vtByte; out->bVal = in->bVal; break;
        case VT_R4:    out->vt = vb6_vtSingle; out->fltVal = in->fltVal; break;
        case VT_R8:    out->vt = vb6_vtDouble; out->dblVal = in->dblVal; break;
        case VT_BSTR:  out->vt = vb6_vtBSTR; out->bstrVal = SysAllocString(in->bstrVal); break;
        default:       out->vt = vb6_vtEmpty; break;
    }
}

// 释放 vb6_Host_Call/GetProp 填出的 vb6_VARIANT (BSTR 需 SysFreeString)
void vb6_Host_ClearVariant(void* v) {
    vb6_VARIANT* p = (vb6_VARIANT*)v;
    if (!p) return;
    if (p->vt == vb6_vtBSTR && p->bstrVal) { SysFreeString(p->bstrVal); }
    memset(p, 0, sizeof(*p));
    p->vt = vb6_vtEmpty;
}

// ============================================================
// czUI fix: 宿主对象的 IDispatch 包装 — vb6_ComPackObject 对宿主对象
// (HWND/UC 实例/集合/字体) 直接走 ((IDispatch*)obj)->lpVtbl->AddRef,
// 会把结构体首字段当 vtable → AV (Charts2020 ClsResizer 实测)。
// 包装成真实 COM 对象: AddRef/Release 引用计数安全, 晚绑定转发到
// vb6_Host_Call/GetProp/SetProp。真实 COM/ActiveX 对象不走包装
// (vb6_Host_IsHostObject 排除), 仍按原 AddRef/Release。
// ============================================================
typedef struct Vb6HostWrap {
    IDispatch disp;
    ULONG refs;
    void* target;
    wchar_t names[24][64];
    int nameCount;
} Vb6HostWrap;

static HRESULT STDMETHODCALLTYPE HW_QI(IDispatch* self, REFIID riid, void** out) {
    if (!out) return E_POINTER;
    *out = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *out = self; self->lpVtbl->AddRef(self); return S_OK;
    }
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE HW_AddRef(IDispatch* self) {
    Vb6HostWrap* w = (Vb6HostWrap*)self; return ++w->refs;
}
static ULONG STDMETHODCALLTYPE HW_Release(IDispatch* self) {
    Vb6HostWrap* w = (Vb6HostWrap*)self;
    ULONG r = --w->refs;
    if (r == 0) free(w);
    return r;
}
static HRESULT STDMETHODCALLTYPE HW_GetTypeInfoCount(IDispatch* self, UINT* n) {
    if (n) *n = 0; return S_OK;
}
static HRESULT STDMETHODCALLTYPE HW_GetTypeInfo(IDispatch* self, UINT i, LCID l, ITypeInfo** t) {
    (void)t; return E_NOTIMPL;
}
static HRESULT STDMETHODCALLTYPE HW_GetIDsOfNames(IDispatch* self, REFIID riid,
        LPOLESTR* names, UINT cNames, LCID lcid, DISPID* out) {
    (void)riid; (void)lcid;
    Vb6HostWrap* w = (Vb6HostWrap*)self;
    if (!out || !cNames || !names || !names[0]) return E_POINTER;
    for (UINT k = 0; k < cNames; k++) {
        int found = -1;
        for (int i = 0; i < w->nameCount; i++)
            if (_wcsicmp(w->names[i], names[k]) == 0) { found = i; break; }
        if (found < 0 && w->nameCount < 24) {
            _snwprintf(w->names[w->nameCount], 63, L"%s", names[k]);
            found = w->nameCount++;
        }
        if (found < 0) return DISP_E_UNKNOWNNAME;
        out[k] = found + 1;
    }
    return S_OK;
}
static HRESULT STDMETHODCALLTYPE HW_Invoke(IDispatch* self, DISPID id, REFIID riid, LCID lcid,
        WORD flags, DISPPARAMS* pd, VARIANT* result, EXCEPINFO* ei, UINT* argErr) {
    (void)riid; (void)lcid; (void)ei; (void)argErr;
    Vb6HostWrap* w = (Vb6HostWrap*)self;
    int idx = (int)id - 1;
    if (idx < 0 || idx >= w->nameCount || !pd) return DISP_E_MEMBERNOTFOUND;
    wchar_t* name = w->names[idx];
    if (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
        if (pd->cArgs < 1) return DISP_E_BADPARAMCOUNT;
        vb6_VARIANT hv; vb6_ho_setVariantEmpty(&hv);
        vb6_Host_FromWinVariant(&pd->rgvarg[0], &hv);
        vb6_Host_SetProp(w->target, name, &hv);
        return S_OK;
    }
    // 方法/属性读取: DISPPARAMS 逆序 → RTL 正序变体数组
    vb6_VARIANT* hargs = NULL;
    if (pd->cArgs > 0) {
        hargs = (vb6_VARIANT*)calloc((size_t)pd->cArgs, sizeof(vb6_VARIANT));
        for (int i = 0; i < pd->cArgs; i++)
            vb6_Host_FromWinVariant(&pd->rgvarg[pd->cArgs - 1 - i], &hargs[i]);
    }
    vb6_VARIANT out; vb6_ho_setVariantEmpty(&out);
    int handled = vb6_Host_Call(w->target, name, pd->cArgs, (void**)hargs, &out);
    if (!handled && (flags & DISPATCH_PROPERTYGET))
        handled = vb6_Host_GetProp(w->target, name, &out);
    if (hargs) {
        for (int i = 0; i < pd->cArgs; i++) vb6_Host_ClearVariant(&hargs[i]);
        free(hargs);
    }
    if (handled && result) {
        VariantInit(result);
        vb6_Host_ToWinVariant(&out, result);
    }
    vb6_Host_ClearVariant(&out);
    return S_OK;
}

static IDispatchVtbl g_vb6HostWrapVtbl = {
    HW_QI, HW_AddRef, HW_Release,
    HW_GetTypeInfoCount, HW_GetTypeInfo, HW_GetIDsOfNames, HW_Invoke
};

void* vb6_UC_WrapHostObject(void* obj) {
    Vb6HostWrap* w = (Vb6HostWrap*)calloc(1, sizeof(Vb6HostWrap));
    if (!w) return NULL;
    w->disp.lpVtbl = &g_vb6HostWrapVtbl;
    w->refs = 1;
    w->target = obj;
    // 注册包装器 → 原对象 映射, 供 IsHostObject/Host_* 解包 (透明性)
    // Fix 126: 容量不足(或尚未分配)时扩容 —— 旧逻辑 `g_uc_wrapCount < g_uc_wrapCap`
    // 在 g_uc_wrapCap==0 时恒为 false, 导致 g_uc_wraps 永不分配、任何宿主对象都
    // 不被登记, vb6_UC_UnwrapHost 永远返回包装器本身而非原对象 → 字体/集合经
    // ReadProperty 默认路径取回后变成 Vb6HostWrap 而非真实结构体, 图表标题/百分比
    // 文字全部缺失 (Charts 2020 实测)。改为"满则扩容"。
    if (g_uc_wrapCount >= g_uc_wrapCap) {
        int32_t ncap = g_uc_wrapCap ? g_uc_wrapCap * 2 : 32;
        Vb6WrapPair* ng = (Vb6WrapPair*)realloc(g_uc_wraps, sizeof(Vb6WrapPair) * ncap);
        if (ng) { g_uc_wraps = ng; g_uc_wrapCap = ncap; }
    }
    if (g_uc_wraps) {
        g_uc_wraps[g_uc_wrapCount].wrap = w;
        g_uc_wraps[g_uc_wrapCount].target = obj;
        g_uc_wrapCount++;
    }
    return w;
}

void* vb6_UC_UnwrapHost(void* obj) {
    for (int i = 0; i < g_uc_wrapCount; i++)
        if (g_uc_wraps[i].wrap == obj) return g_uc_wraps[i].target;
    return obj;
}

#ifdef __cplusplus
} // extern "C"
#endif
