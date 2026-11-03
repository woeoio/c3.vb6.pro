// uc_hostmodel.c - vb6forms_uc 拆分片：窗体/控件宿主对象模型（注册/判定/分派）
//
// 内容 = 拆分前 vb6forms_uc.c 第 755~819 / 1187~1300 / 1303~1303 / 1391~1391 / 1423~1423 / 1497~1542 行，纯搬移零重排无行为改动
// 跨族共享符号见 vb6forms_uc_internal.h

#include "vb6forms_uc_internal.h"
#include "vb6forms_prop.h"      // 账 #247: 控件几何的唯一出口 (vb6_Get/SetControlLeft|Top|Width|Height)
#include "vb6forms_prop_ctrl.h"   // Fix 143: vb6_AddItem/RemoveItem/ClearList (原生列表 COM 晚绑定)
extern void vb6_ShowForm(void* hwnd, int modal);  /* 窗体 Show 的唯一出口 (vb6forms.c) */

#ifdef __cplusplus
extern "C" {
#endif

// 账 #305 第二百六十二刀 (丙): 宿主模型「不认识的成员」计数。
// 宿主通道对**任何**成员名都返回「已处理」 (三条 .inc 的末尾 return 1),
// 于是 ScaleX / ScaleY / Size 这类没人认领的成员被静默答成 Empty ——
// 调用方拿到空值也看不出漏了谁。这里把「不认识的名字」数出来并报一行,
// 行为不变 (仍 return 1), 只是把静默换成响亮。
static int32_t g_host_unknown_member = 0;

int32_t vb6_Host_UnknownMemberCount(void) { return g_host_unknown_member; }

/* 账 (438 语义刀): 读侧 (GetProp / Call) 不认识的名字不再算「缺口」——
 * VB6 在这里抛 438, 由 On Error Resume Next 吞掉, 引擎侧的行为是**对的**。
 * 它们单独记一枚计数器与一行 [C3_HOST_438], 与真缺口 (SetProp 等仍未实现、
 * 静默返回 1 的名字) 分开计数: GUISample 实测 = 274 条 GetProp "rtfText"
 * (frmDocument 没有该成员) + 3 条 SetProp "Icons" (真缺口)。 */
static int32_t g_host_438_member = 0;

int32_t vb6_Host_Error438Count(void) { return g_host_438_member; }

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
    /* 账 #305 第二百八十二刀: 无窗宿主 (InvisibleAtRuntime) 递进来的是**实例指针**,
       上面两问都落空 ⇒ 以前这里答 NULL，晚绑定那头就报 "Object"，而 Krool 那一族全靠
       `If TypeName(ControlEnum) = "ImageList" Then` 认控件（靶工程 ToolBar/ListView/
       TreeView/OptionButtonW 四处都这么写）—— 答错一个名，整条 ImageList 接线全断。 */
    h = vb6_ho_findInstance(obj);
    if (h && h->typeName[0]) return h->typeName;
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

/* 账 #305 第二百七十六刀: 句柄那一档 —— 与上面那枚差在**宽度**。
 * VB6 的 IFont.hFont 在类型库里标的是 Long, 而语料一律把它交给 As LongPtr 的形参
 * (靶工程 Common.bas:231 那份 PtrSafe 声明), x64 上按 int32 交就是把 GDI 句柄砍掉
 * 高 32 位。标签交 VT_I8: vb6_vartype 那张表里没有这一档, 按 §B142 那格的口径
 * 由 vb6_VariantPackBorrow 原样递进真 VARIANT 的 llVal (第五十三刀刚把它并表)。 */
void vb6_ho_setVariantLongPtr(vb6_VARIANT* out, void* p) {
    memset(out, 0, sizeof(*out));
    out->vt = (vb6_vartype)VT_I8;
    out->llVal = (LONG_PTR)p;
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

// 账 #305 §B224: 晚绑定对**带型变量**递 VT_BYREF|<type> (见 vb6_ComPackIntRef 那一族),
// union 里存的是调用方变量的地址, 不是值。宿主对象模型此前直接按值读 union ⇒
// `c.Remove j` 读成指针低位/0, Builtin Collection 的 Remove 静默无效 ——
// Charts 2020 ucTreeMaps.AddLineSeries 的降序排序 (`Remove j : Add v, , i`)
// 因此只增不删, 标签重复、Count 从 7 涨到 13。这里按 §B224 的标签逐档解引用。
// 注意: 这些指针**指向标量本身**, 不是指向另一枚 VARIANT (VT_VARIANT 那档除外)。
static const void* vb6_ho_byrefAddr(const vb6_VARIANT* v) {
    if (!v || !(v->vt & VT_BYREF)) return NULL;
    return (const void*)v->pdispVal;
}

int32_t vb6_ho_variantToLong(const vb6_VARIANT* v) {
    if (!v) return 0;
    if (v->vt & VT_BYREF) {
        const void* p = vb6_ho_byrefAddr(v);
        if (!p) return 0;
        switch (v->vt & ~VT_BYREF) {
            case vb6_vtInteger: case vb6_vtBoolean: return (int32_t)*(const int16_t*)p;
            case vb6_vtLong: return *(const int32_t*)p;
            case vb6_vtByte: return (int32_t)*(const uint8_t*)p;
            case vb6_vtSingle: return (int32_t)*(const float*)p;
            case vb6_vtDouble: return (int32_t)*(const double*)p;
            case VT_I8: return (int32_t)*(const int64_t*)p;
            case (vb6_vartype)VT_I1:  return (int32_t)*(const int8_t*)p;
            case (vb6_vartype)VT_UI2: return (int32_t)*(const uint16_t*)p;
            case VT_UI4: return (int32_t)*(const uint32_t*)p;
            case VT_UI8: return (int32_t)*(const uint64_t*)p;
            case VT_VARIANT: return vb6_ho_variantToLong((const vb6_VARIANT*)p);
            default: return 0;
        }
    }
    switch (v->vt) {
        case vb6_vtInteger: case vb6_vtLong: case vb6_vtBoolean: case vb6_vtByte:
            return v->lVal;
        case vb6_vtSingle: case vb6_vtDouble: return (int32_t)v->dblVal;
        case VT_I8: return (int32_t)v->llVal;   /* 账 #305: 这一档以前落 default ⇒ 读回 0 */
        case (vb6_vartype)VT_I1:  return (int32_t)v->sbVal;  /* ai/032 §6.9 */
        case (vb6_vartype)VT_UI2: return (int32_t)v->usVal;  /* ai/032 §6.9 */
        case VT_UI4: return (int32_t)v->ulVal;  /* ai/032 §6.6: 无符号两档同格 */
        case VT_UI8: return (int32_t)v->ullVal;
        default: return 0;
    }
}

double vb6_ho_variantToDouble(const vb6_VARIANT* v) {
    if (!v) return 0.0;
    if (v->vt & VT_BYREF) {   /* 账 #305 §B224: 与 ToLong 同一口径 */
        const void* p = vb6_ho_byrefAddr(v);
        if (!p) return 0.0;
        switch (v->vt & ~VT_BYREF) {
            case vb6_vtSingle: return (double)*(const float*)p;
            case vb6_vtDouble: return *(const double*)p;
            case vb6_vtInteger: case vb6_vtBoolean: return (double)*(const int16_t*)p;
            case vb6_vtLong: return (double)*(const int32_t*)p;
            case vb6_vtByte: return (double)*(const uint8_t*)p;
            case VT_I8: return (double)*(const int64_t*)p;
            case (vb6_vartype)VT_I1:  return (double)*(const int8_t*)p;
            case (vb6_vartype)VT_UI2: return (double)*(const uint16_t*)p;
            case VT_UI4: return (double)*(const uint32_t*)p;
            case VT_UI8: return (double)*(const uint64_t*)p;
            case VT_VARIANT: return vb6_ho_variantToDouble((const vb6_VARIANT*)p);
            default: return 0.0;
        }
    }
    switch (v->vt) {
        case vb6_vtSingle: case vb6_vtDouble: return v->dblVal;
        case vb6_vtInteger: case vb6_vtLong: case vb6_vtBoolean: return (double)v->lVal;
        case VT_I8: return (double)v->llVal;   /* 账 #305: 同上 */
        case (vb6_vartype)VT_I1:  return (double)v->sbVal;  /* ai/032 §6.9 */
        case (vb6_vartype)VT_UI2: return (double)v->usVal;  /* ai/032 §6.9 */
        case VT_UI4: return (double)v->ulVal;  /* ai/032 §6.6: 同上 */
        case VT_UI8: return (double)v->ullVal;
        default: return 0.0;
    }
}

BSTR vb6_ho_variantToBstr(const vb6_VARIANT* v) {
    if (!v) return NULL;
    if (v->vt & VT_BYREF) {   /* 账 #305 §B224: ByRef String 递的是 BSTR* */
        const void* p = vb6_ho_byrefAddr(v);
        if (!p) return NULL;
        switch (v->vt & ~VT_BYREF) {
            case vb6_vtBSTR: return *(BSTR*)p;
            case VT_VARIANT: return vb6_ho_variantToBstr((const vb6_VARIANT*)p);
            default: return NULL;
        }
    }
    return (v->vt == vb6_vtBSTR) ? v->bstrVal : NULL;
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
    /* 账 #305 第二百七十六刀: `hFont` —— 与 vb6com_invoke.c 那条 kind 5 臂**同一个答案、
       同一位主人** (vb6_UC_FontHFont → 内部问 vb6_UC_FontLogFont 那唯一一份换算)。
       造不出句柄就答"认不出"(return 0), 由调用方那一条报数出口说话 —— 与第二百七十三刀
       给方法臂立的口径一致: 这一臂以前对认不出的名字静默答 Empty, 所以靶工程那 34 发
       `name=hFont` 在 C3_HOST_MEMBER_MISS 名册里一条也没有 (第二百七十五刀记的那格洞)。 */
    if (_wcsicmp(name, L"hFont") == 0) {
        void* h = vb6_UC_FontHFont(f);
        if (!h) return 0;
        vb6_ho_setVariantLongPtr(out, h); return 1;
    }
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

/* 账 #305 第二百七十三刀: 「报一行宿主通道认不出的名字」从三条岔口的手抄收成一位主人。
 * 第二百六十二刀装护栏时, 字体那一臂在 uc_hostmodel_call.inc 的第一行就 return 1 ——
 * 位置在护栏之前, 于是靶工程那 13 发 Font.Clone 既空转又一声不响 (实测 stderr 里
 * 既无 C3_HOST_MEMBER_MISS、也无 getDispid("Clone")、也无 method "Clone" not found)。
 * 行为面一字不改: 认不出仍是「答了但没答」(调用方拿空值), 只是从此响亮。
 * 站点名由调用方交进来, 打出的字节与改前那三条各自的手抄完全同形。 */
static void vb6_ho_reportUnknown(const wchar_t* site, const wchar_t* name, void* obj) {
    g_host_unknown_member++;
    {   /* 诊断: 认不出时把接收者身份一起打出来 (窗口? 类名? 父类名?) */
        wchar_t cls[128] = L"-", pcls[128] = L"-";
        HWND h = (HWND)obj;
        if (h && IsWindow(h)) {
            GetClassNameW(h, cls, 128);
            HWND p = GetParent(h);
            if (p) GetClassNameW(p, pcls, 128);
        }
        fwprintf(stderr, L"[C3_HOST_MEMBER_MISS #%ld] site=%ls name=\"%ls\" obj=%p isWin=%d cls=%ls parent=%ls\n",
                 (long)g_host_unknown_member, site, name ? name : L"(null)", obj,
                 (h && IsWindow(h)) ? 1 : 0, cls, pcls);
    }
    fflush(stderr);
}

/* 账 (438 语义刀): 读侧成员不存在 = VB6 的 error 438, 不是缺口。单独一行标记,
 * 便于把「被 On Error Resume Next 吞掉的 438」与真缺口分开数。 */
static void vb6_ho_report438(const wchar_t* site, const wchar_t* name, void* obj) {
    g_host_438_member++;
    wchar_t cls[128] = L"-", pcls[128] = L"-";
    HWND h = (HWND)obj;
    if (h && IsWindow(h)) {
        GetClassNameW(h, cls, 128);
        HWND p = GetParent(h);
        if (p) GetClassNameW(p, pcls, 128);
    }
    fwprintf(stderr, L"[C3_HOST_438 #%ld] site=%ls name=\"%ls\" obj=%p isWin=%d cls=%ls parent=%ls\n",
             (long)g_host_438_member, site, name ? name : L"(null)", obj,
             (h && IsWindow(h)) ? 1 : 0, cls, pcls);
    fflush(stderr);
}

/* 宿主字体伪对象的**方法**面 (账 #305 第二百七十三刀)。
 * stdole2.tlb 里 IFont.Clone 的档位 = invkind=1 (method) + 一枚 [out] 参
 * (FUNC 17 memid=1610678289 nparams=1) ⇒ 结果**只**走实参, VARIANT 返回值天生是空档
 * (实测 `CallObject: disp=… method=Clone vt=0`), 而发码把这一发当**语句**发、
 * 返回值就地丢弃 ⇒ 光答 out 拿不到东西, 必须把新对象写进 argv[0] 那枚
 * VT_DISPATCH|VT_BYREF 指向的存储 (= vb6_ComPackObjectRef 交出的调用方槽)。
 * 其余三枚方法 IsEqual / SetRatio / AddRefHfont / ReleaseHfont 在靶工程发码里 0 处
 * ⇒ 不发明没人走的答案, 认不出就走上面那枚报告。
 * 返回值那一腿**刻意不碰** out: 岔口进来时已被 `vb6_ho_setVariantEmpty` 清零, 而伪字体
 * 不许包成 VT_DISPATCH (= 上面 159 行那两条注释说的同一桩事故: VariantClear 会对它调
 * Release)。写 argv[0] 是唯一出路, 也是一条不发明新档的路。
 * 所有权: 新建那枚伪字体进 vb6_ucFonts 那张表 (= 与 getprop 的 Font 那一臂同一个工厂),
 * 回收仍欠 §B164 那一族; 这里刻意不发明第二种生命周期。 */
static int32_t vb6_ho_fontMethod(void* fontProxy, const wchar_t* name, int32_t argc,
                                 void** argv) {
    vb6_ComIface_Font* src = vb6_uc_fontOf(fontProxy);
    if (!src || !name) return 0;
    if (_wcsicmp(name, L"Clone") != 0 || argc < 1 || !argv || !argv[0]) return 0;
    vb6_ComIface_Font* dst = (vb6_ComIface_Font*)vb6_UC_NewFont();
    if (!dst) return 0;
    if (dst->Name) SysFreeString(dst->Name);
    dst->Name = src->Name ? vb6_BSTR_FromStr(src->Name) : NULL;
    dst->Size = src->Size;
    dst->Bold = src->Bold;
    dst->Italic = src->Italic;
    dst->Underline = src->Underline;
    dst->Strikethrough = src->Strikethrough;
    dst->Weight = src->Weight;
    dst->Charset = src->Charset;
    /* vb6_VARIANT 与 Windows VARIANT 同布局 (rev11) ⇒ 这一读不是猜型, 是那一刀的成品。 */
    VARIANT* a = (VARIANT*)argv[0];
    if ((a->vt & VT_BYREF) && a->ppdispVal) *(void**)(void*)a->ppdispVal = (void*)dst;
    return 1;
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
