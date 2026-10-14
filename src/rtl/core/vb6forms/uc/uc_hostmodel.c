// uc_hostmodel.c - vb6forms_uc 拆分片：窗体/控件宿主对象模型（注册/判定/分派）
//
// 内容 = 拆分前 vb6forms_uc.c 第 755~819 / 1187~1300 / 1303~1303 / 1391~1391 / 1423~1423 / 1497~1542 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"
#include "vb6forms_prop.h"      // 账 #247: 控件几何的唯一出口 (vb6_Get/SetControlLeft|Top|Width|Height)
#include "vb6forms_prop_ctrl.h"   // Fix 143: vb6_AddItem/RemoveItem/ClearList (原生列表 COM 晚绑定)

#ifdef __cplusplus
extern "C" {
#endif

// ============================================================
// 窗体/控件登记 (供宿主对象分派)
// ============================================================

static void ho_fillField(wchar_t* dst, const char* src) {
    if (!src) return;
    size_t n = strlen(src);
    if (n >= VB6_UC_NAME_LEN) n = VB6_UC_NAME_LEN - 1;
    for (size_t i = 0; i < n; i++) dst[i] = (wchar_t)src[i];
    dst[n] = 0;
}

void vb6_HostObj_Register(void* hwnd, const char* name, const char* vbTypeName,
                          int32_t isForm, int32_t index) {
    if (!hwnd) return;
    /* 账 #257: 「同 hwnd 已登记」以前整条 return —— 后到的**带名**登记被静默丢掉，
       于是同一枚窗口第二次说话永远说不出口（窗体先由 vb6_Forms_Register 登记，
       名字晚一步到就再也进不去）。现在只补还空着的 name / typeName 两格。
       isForm 与 index 刻意不动：那是存量答案的一部分（非数组控件的 .Index 恒 -1，
       夹具 GC22 就钉着这个数），不能让后到的登记顺手改。 */
    vb6_HostObjRec* found = vb6_ho_find(hwnd);
    if (found) {
        if (!found->name[0] && name) ho_fillField(found->name, name);
        if (!found->typeName[0] && vbTypeName) {
            const char* dot = strrchr(vbTypeName, '.');
            ho_fillField(found->typeName, dot ? dot + 1 : vbTypeName);
        }
        return;
    }
    if (vb6_ucHoCount >= VB6_UC_MAX_OBJ) return;
    vb6_HostObjRec* h = &vb6_ucHo[vb6_ucHoCount++];
    memset(h, 0, sizeof(*h));
    h->hwnd = hwnd;
    h->isForm = isForm;
    h->index = index;
    ho_fillField(h->name, name);
    if (vbTypeName) {
        const char* dot = strrchr(vbTypeName, '.');
        ho_fillField(h->typeName, dot ? dot + 1 : vbTypeName);
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
int32_t vb6_HostObj_Count(void) { return vb6_ucHoCount; }
void* vb6_HostObj_At(int32_t i) {
    if (i < 0 || i >= vb6_ucHoCount) return NULL;
    return vb6_ucHo[i].hwnd;
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
        case VT_I8: return (int32_t)v->llVal;   /* 账 #305: 这一档以前落 default ⇒ 读回 0 */
        case VT_UI4: return (int32_t)v->ulVal;  /* ai/032 §6.6: 无符号两档同格 */
        case VT_UI8: return (int32_t)v->ullVal;
        default: return 0;
    }
}

double vb6_ho_variantToDouble(const vb6_VARIANT* v) {
    if (!v) return 0.0;
    switch (v->vt) {
        case vb6_vtSingle: case vb6_vtDouble: return v->dblVal;
        case vb6_vtInteger: case vb6_vtLong: case vb6_vtBoolean: return (double)v->lVal;
        case VT_I8: return (double)v->llVal;   /* 账 #305: 同上 */
        case VT_UI4: return (double)v->ulVal;  /* ai/032 §6.6: 同上 */
        case VT_UI8: return (double)v->ullVal;
        default: return 0.0;
    }
}

BSTR vb6_ho_variantToBstr(const vb6_VARIANT* v) {
    return (v && v->vt == vb6_vtBSTR) ? v->bstrVal : NULL;
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
    /* 账 #305 §B142 (第五十三刀): 这一格原来是一张十档的手抄表，Decimal / LongLong(VT_I8) /
       Currency / Date 一律落到 default ⇒ VT_EMPTY。实测 `c.Add CDec(5)` 之后 `VarType(c(1))` 与
       `For Each` 读回来都是 0/Empty，x64 与 x86 两台逐行相同(`.build/k53_p/x64b.out` / `x86b.out`)，
       而同一枚值直接装箱读是 14/Decimal ⇒ 罪就在这一格，不在写侧。
       「这枚槽的数值住在真 VARIANT 的哪一格」§B145 已经收成一枚权威(vb6_VariantPackBorrow，
       十四档 + 数组那一问都答)，这里改问它，只补它刻意不做的那一步: 宿主管道交出去的串是
       **带权副本**(调用方随后 Host_ClearVariant 会放掉它)。**对象那一档照旧不 AddRef** ——
       与改前一字不差：反向的 FromWinVariant 借来的指针不计数，Clear 也只放串，
       这里给一次 AddRef 就成双份(§B144 记的那条收支)。 */
    vb6_VariantPackBorrow(*in, out);
    if (out->vt == VT_BSTR && out->bstrVal) out->bstrVal = SysAllocString(out->bstrVal);
}

void vb6_Host_FromWinVariant(const void* inV, void* outV) {
    const VARIANT* in = (const VARIANT*)inV;
    vb6_VARIANT* out = (vb6_VARIANT*)outV;
    memset(out, 0, sizeof(*out));
    if (!in) { out->vt = vb6_vtEmpty; return; }
    /* 账 #305 §B165 (第五十四刀): 上面那张九档手抄表撤了 —— Date/Currency/I8/Decimal 一律
       落 default ⇒ VT_EMPTY。实测 `c.Add CDec(5)` 之后 `VarType(c("d"))` 与 For Each 读回
       都是 0/Empty，两台逐行相同(.build/k54_p/x64b.out / x86b.out)，而同一枚值直接装箱读是
       14/Decimal/5 ⇒ 罪就在这一格，不在写侧，也不在第五十三刀刚并好的那枚核里。
       「对端那 16 字节的数值住在哪一格」§B144/§B162 收成 vb6_VariantBorrowIntoSlot;
       「这一格该自己持有什么」§B165 收成 vb6_VariantValueCopy。两问各只一枚答案，
       这里正是它们的交点，所以改问它们，不抄第三遍。
       ⚠ 对象那一档**照旧借用不计数**，与改前一字不差: ValueCopy 不 AddRef，而本环的
       vb6_Host_ClearVariant 也刻意不放对象 —— 环上递手的可能是 Vb6HostWrap
       (vb6com_foreach.c:160 那一趟的引用数是调用栈上那次 VariantClear 放的)，再取一份就是双份。
       ⚠ 数组 / ByRef / IUnknown 三形**刻意不收**，读数与改前同为 Empty: 数组归 §B159/§B161
       (BorrowIntoSlot 造的 owned 描述符在这条环上没有释放点，收了就漏)，ByRef 归 §B144 的
       另一问，IUnknown 折成对象档那一步住在发码侧而不是这里。 */
    if ((in->vt & (VT_ARRAY | VT_BYREF)) || in->vt == VT_UNKNOWN) {
        out->vt = vb6_vtEmpty;
        return;
    }
    vb6_VARIANT borrowed;
    vb6_VariantBorrowIntoSlot(in, &borrowed);
    vb6_VariantValueCopy(out, &borrowed);
}

// 释放 vb6_Host_Call/GetProp 填出的 vb6_VARIANT: 值档问 §B165 那唯一一枚出口
// (改前只放串 ⇒ 第五十四刀新存下来的 x86 DECIMAL 会漏)。对象那一档照旧不动，与
// vb6_Host_FromWinVariant 的「借用不计数」配平。
void vb6_Host_ClearVariant(void* v) {
    vb6_VARIANT* p = (vb6_VARIANT*)v;
    if (!p) return;
    vb6_VariantValueFree(p);
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
/* 账 #255: 包装器和它在 (wrap->target) 表里的登记必须同生同死。
   表只在 vb6_UC_WrapHostObject 里追加、从来没人撤，而 free 掉的块地址会被下一枚
   包装器原样复用 —— vb6_UC_UnwrapHost 从 i=0 扫表、先撞上那条早已释放的旧登记，
   于是「每一枚成员」都解回「第一枚被包过的对象」。
   实测 (probe255 模式 M)：For Each o In Me.Controls 的三枚成员 hWnd 同为 4326406、
   Left 同为 120（那是 Timer 的设计期值），对它写 .Width 只落在这枚 Timer 的存属性上，
   txtA/picP 一动不动；Charts 的 ClsResizer 因此把 N 格存成同一枚控件。 */
static void vb6_UC_DropWrapPair(void* wrap) {
    for (int i = 0; i < g_uc_wrapCount; i++) {
        if (g_uc_wraps[i].wrap == wrap) {
            g_uc_wraps[i] = g_uc_wraps[--g_uc_wrapCount];  /* 末位填补：查表按指针，次序无意义 */
            return;
        }
    }
}
static ULONG STDMETHODCALLTYPE HW_Release(IDispatch* self) {
    Vb6HostWrap* w = (Vb6HostWrap*)self;
    ULONG r = --w->refs;
    if (r == 0) { vb6_UC_DropWrapPair(w); free(w); }
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
