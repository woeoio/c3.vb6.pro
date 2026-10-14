// vb6forms_memberobj.c — 控件"成员对象 / 成员集合"的原生 IDispatch 实现
// (ai/029 C29-3 立样, C29-7 ListView 复用)
//
// 背景: VB6 里 ListView1.ListItems.Add(..) / ListView1.ColumnHeaders.Add(..) /
// TreeView1.Nodes.Add(..) / StatusBar1.Panels(i) 返回的都是**对象**: 可以
// `Set itm = ...` 接住再 `itm.Text` / `itm.SubItems(1)` 访问成员; `For Each n In Nodes`
// 还要能枚举 —— 集合必须可枚举, 元素必须是对象。
//
// 计划书 D1 的口径 = **走真 IDispatch**:
//   ① `As Object` 的晚绑定 (vb6_ComGetProp / vb6_ComSetProp / vb6_ComCallObject 那一套)
//      直接生效, cgen 侧几乎零改动;
//   ② 不用为每个成员属性在 cgen 里再长一条字符串级特例 —— 历史上 vb6_CStr 那一族
//      长出一百多个 Fix, 正是"特例互相打脸"的形状;
//   ③ 顺带把 022/D71 "工程内类交给 As Object 点不到成员" 那个缺口需要的同一套建起来。
//
// 现有族 (按 kind 分派; 新增控件只需加一张名字表 + 一段 get/set):
//   VB6_MEMK_LISTIMAGE    / VB6_MEMCK_LISTIMAGES     ImageList  owner = vb6_com_X(复刻实例指针)
//   VB6_MEMK_LISTITEM     / VB6_MEMCK_LISTITEMS      ListView   owner = vb6_hwnd_X(HWND)
//   VB6_MEMK_COLUMNHEADER / VB6_MEMCK_COLUMNHEADERS  ListView   owner = vb6_hwnd_X(HWND)
//   VB6_MEMK_PANEL        / VB6_MEMCK_PANELS         StatusBar  owner = vb6_hwnd_X(HWND)
//     (C29-4: PanelClick/PanelDblClick 的事件参数是 Panel 对象 —— 标量面/集合面
//      P20-40 已用特例直译通了, 这里只为**事件参数**立对象, 不抢集合那条路。)
//
// ⚠ **owner 的语义按族不同**: ImageList 无窗口, 槽里放的是复刻实例指针; ListView 是真窗口,
// 槽里就是 HWND。各族自己解释那个 void* (名字表/分派里一起做)。
//
// ⚠ **参数化属性** (ListItem.SubItems(i)) 的 PUT 走 vb6_ComSetPropArg, 它的 result 传 **NULL**,
// 所以 PUT 分支必须排在 `if (!result) return S_OK;` **之前**, 否则整条写入被静默吞掉。
// 另外 PUT 的参数布局是 `rgvarg[0]`=新值、`rgvarg[1..]`=索引(**正序**), 与 CALL/GET 的
// **逆序**不同 —— 抄错了症状是"写进去了但写的是别的下标"。
//
// 骨架照 vb6forms_axcontainer.c (那份也把 name→dispid 表变成了真 IDispatch)。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6forms_prop_ctrl.h"   /* vb6_ImageList_* / vb6_ListView_* 原型 */
#include "vb6rtl_bstr.h"
#include "vb6rtl_variant.h"
#include "vb6rtl_class_com.h"
#include <oleauto.h>
#include <olectl.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>

#ifdef _WIN32

// ===================== 族号 =====================
// 成员对象族 (Vb6MemObj)
#define VB6_MEMK_LISTIMAGE      1
#define VB6_MEMK_LISTITEM       2
#define VB6_MEMK_COLUMNHEADER   3
#define VB6_MEMK_PANEL          4
#define VB6_MEMK_RS             5   // Data.Recordset (C29-Data)
#define VB6_MEMK_FIELD          6   // Recordset.Fields("x") 的 Field
#define VB6_MEMK_NODE           7   // TreeView 的 Node (C29-8b)
#define VB6_MEMK_BUTTON         8   // Toolbar 的 Button (C29-5b)
// 成员集合族 (Vb6MemColl)
#define VB6_MEMCK_LISTIMAGES    1
#define VB6_MEMCK_LISTITEMS     2
#define VB6_MEMCK_COLUMNHEADERS 3
#define VB6_MEMCK_PANELS        4
#define VB6_MEMCK_NODES         5   // TreeView 的 Nodes (C29-8b)
#define VB6_MEMCK_BUTTONS       6   // Toolbar 的 Buttons (C29-5b)
// (Recordset/Field 不是集合, RS.Fields 是带参 GET 直接产 Field 对象)

// ===================== 实例 =====================
typedef struct Vb6MemObj {
    IDispatchVtbl* lpVtbl;
    LONG      ref;
    int32_t   kind;
    void*     owner;    // 宿主控件的 RTL 槽 (ImageList = 实例指针; ListView = HWND)
    int32_t   index;    // 1 基
} Vb6MemObj;

typedef struct Vb6MemColl {
    IDispatchVtbl* lpVtbl;
    LONG      ref;
    int32_t   kind;
    void*     owner;
} Vb6MemColl;

// 集合的可枚举器 (IEnumVARIANT)。For Each 走的就是它。
typedef struct Vb6MemEnum {
    IEnumVARIANTVtbl* lpVtbl;
    LONG      ref;
    int32_t   kind;     // 集合族号 (决定产出什么成员对象)
    void*     owner;
    int32_t   total;    // 枚举开始时快照的 count
    int32_t   pos;      // 下一个要产出的 1 基下标
} Vb6MemEnum;

static const IDispatchVtbl g_memObjVtbl;
static const IDispatchVtbl g_memCollVtbl;
static const IEnumVARIANTVtbl g_memEnumVtbl;

// ===================== 名字表 (DISPID = 下标 + 1) =====================
static const wchar_t* kListImageNames[] = { L"Key", L"Index", L"Picture", NULL };
// ListView 的 ListItem。**没有 Icon/SmallIcon 的读**: 扁平 RTL 那层没提供取图标的入口,
// 名字表里放进去只会"看着支持、实则答错", 所以干脆不登记 (边界写进手册)。
static const wchar_t* kListItemNames[] = {
    L"Key", L"Index", L"Text", L"SubItems", L"Selected", L"Checked", NULL };
// ListView 的 ColumnHeader。同理不收 SubItemIndex。
static const wchar_t* kColumnHeaderNames[] = {
    L"Key", L"Index", L"Text", L"Width", L"Alignment", NULL };
// C29-4: StatusBar 的 Panel (事件参数)。数据面照 vb6forms_statusbar.c 的 getter/setter。
static const wchar_t* kPanelNames[] = {
    L"Key", L"Index", L"Text", L"Width", L"MinWidth", L"AutoSize", L"Style",
    L"ToolTipText", L"Left", NULL };   // Left = 账 #206 第二格补的那一枚 (DISPID 9)
// C29-8b: TreeView 的 Node。名字表顺序即 DISPID (下标 + 1)。
// 只登记扁平层真做得动的成员: Bold / Sorted / RelativeX / Node.Style 一律不放,
// 放进去就是「看着支持、实则答错」 (与 ListView 那族同一取舍)。
static const wchar_t* kNodeNames[] = {
    L"Key", L"Index", L"Text", L"Tag", L"Checked", L"Expanded",
    L"Parent", L"Child", L"Children", L"Next", L"Previous", L"Root",
    L"EnsureVisible", NULL };

// C29-5b: Toolbar 的 Button。文本那条在 VB6 里叫 Caption (不是 Text)。
// Width 登记了但**宽度真变不了** —— 原生 cx 只对分隔符起作用，普通按钮的宽度由
// 图标与文字量出来；这条读数是"VB 侧请求值"，不是控件量出来的那个，注释里写明。
static const wchar_t* kButtonNames[] = {
    L"Key", L"Index", L"Caption", L"Tag", L"ToolTipText",
    L"Style", L"Image", L"Enabled", L"Visible", L"Width", L"Value", NULL };

// 三个集合同构: Count / Item / Add / Remove / Clear / _NewEnum
static const wchar_t* kCollNames[] = {
    L"Count", L"Item", L"Add", L"Remove", L"Clear", L"_NewEnum", NULL };
// C29-Data: Data.Recordset (方法与标量混排; DISPID = 下标+1)
static const wchar_t* kRsNames[] = {
    L"BOF", L"EOF", L"RecordCount", L"FieldCount", L"Fields",
    L"MoveFirst", L"MoveLast", L"MoveNext", L"MovePrevious", L"Refresh",
    L"CurrentRow", NULL };
// Recordset.Fields("x") 的 Field
static const wchar_t* kRsFieldNames[] = {
    L"Name", L"Value", L"Index", NULL };

// DISPID 明细 (与上面各表的顺序严格一致)
#define VB6_MEMD_KEY        1
#define VB6_MEMD_INDEX      2
#define VB6_MEMD_TEXT       3
#define VB6_MEMD_PICTURE    3   // ListImage.Picture —— 与 ListItem/ColumnHeader.Text 同号, 按 kind 分派
#define VB6_MEMD_SUBITEMS   4   // ListItem.SubItems(i) —— 参数化属性
#define VB6_MEMD_WIDTH      4   // ColumnHeader.Width (同一 dispid, 按 kind 分派)
#define VB6_MEMD_SELECTED   5
#define VB6_MEMD_ALIGNMENT  5   // ColumnHeader.Alignment
#define VB6_MEMD_CHECKED    6
// Panel (kPanelNames 顺序: Key,Index,Text,Width,MinWidth,AutoSize,Style,ToolTipText)
#define VB6_MEMD_MINWIDTH   5
#define VB6_MEMD_AUTOSIZE   6
#define VB6_MEMD_STYLE      7
#define VB6_MEMD_TOOLTIP    8
#define VB6_MEMD_PANEL_LEFT 9   // Panel.Left —— 排版后的左界 (缇), 与 Width 同源
// Node (kNodeNames 顺序; Key/Index/Text 沿用上面同号的定义)
#define VB6_MEMD_NODE_TAG       4
#define VB6_MEMD_NODE_CHECKED   5
#define VB6_MEMD_NODE_EXPANDED  6
#define VB6_MEMD_NODE_PARENT    7
#define VB6_MEMD_NODE_CHILD     8
#define VB6_MEMD_NODE_CHILDREN  9
#define VB6_MEMD_NODE_NEXT      10
#define VB6_MEMD_NODE_PREV      11
#define VB6_MEMD_NODE_ROOT      12
#define VB6_MEMD_NODE_ENSVIS    13
// Button (kButtonNames 顺序; Key/Index 沿用上面同号的定义)
#define VB6_MEMD_BTN_CAPTION   3
#define VB6_MEMD_BTN_TAG       4
#define VB6_MEMD_BTN_TIP       5
#define VB6_MEMD_BTN_STYLE     6
#define VB6_MEMD_BTN_IMAGE     7
#define VB6_MEMD_BTN_ENABLED   8
#define VB6_MEMD_BTN_VISIBLE   9
#define VB6_MEMD_BTN_WIDTH    10
#define VB6_MEMD_BTN_VALUE    11
// RS (kRsNames 顺序)
#define VB6_MEMRS_BOF        1
#define VB6_MEMRS_EOF        2
#define VB6_MEMRS_RECORDCNT  3
#define VB6_MEMRS_FIELDCNT   4
#define VB6_MEMRS_FIELDS     5
#define VB6_MEMRS_MOVEFIRST  6
#define VB6_MEMRS_MOVELAST   7
#define VB6_MEMRS_MOVENEXT   8
#define VB6_MEMRS_MOVEPREV   9
#define VB6_MEMRS_REFRESH    10
#define VB6_MEMRS_CURROW     11
// Field (kRsFieldNames 顺序)
#define VB6_MEMF_NAME        1
#define VB6_MEMF_VALUE       2
#define VB6_MEMF_INDEX       3

#define VB6_MEMCD_COUNT     1
#define VB6_MEMCD_ITEM      2
#define VB6_MEMCD_ADD       3
#define VB6_MEMCD_REMOVE    4
#define VB6_MEMCD_CLEAR     5
#define VB6_MEMCD_NEWENUM   6

#define VB6_DISPID_NEWENUM  (-4)   // OLE 约定: _NewEnum

static const wchar_t* const* memObjNamesOf(int32_t kind) {
    switch (kind) {
    case VB6_MEMK_LISTITEM:     return kListItemNames;
    case VB6_MEMK_COLUMNHEADER: return kColumnHeaderNames;
    case VB6_MEMK_PANEL:        return kPanelNames;
    case VB6_MEMK_NODE:        return kNodeNames;
    case VB6_MEMK_BUTTON:      return kButtonNames;
    case VB6_MEMK_RS:           return kRsNames;
    case VB6_MEMK_FIELD:        return kRsFieldNames;
    default:                    return kListImageNames;
    }
}
static const wchar_t* const* memCollNamesOf(int32_t kind) {
    (void)kind;   // 三个集合同构
    return kCollNames;
}

static int memLookup(const wchar_t* const* tbl, const wchar_t* name) {
    if (!name) return 0;
    for (int i = 0; tbl[i]; i++)
        if (_wcsicmp(tbl[i], name) == 0) return i + 1;
    return 0;
}

// C29-3 临时排障探针 (C3_MEMOBJ_TRACE=1 打开)。x86 上"对象被过早释放"会立刻踩到,
// x64 却常常侥幸不崩 —— 定位 use-after-free 只有把 new/Release 打出来才看得见。
static int memTrace(void) {
    static int c = -1;
    if (c < 0) c = (GetEnvironmentVariableW(L"C3_MEMOBJ_TRACE", NULL, 0) > 0) ? 1 : 0;
    return c;
}

// ===================== VARIANT 写出助手 =====================
// ⚠ 这里必须用 **Windows 的 VARIANT / VT_\***, 不能按 vb6_VARIANT 写:
// IDispatch::Invoke 的 result 是调用方按 Windows VARIANT 分配的 (vb6_ComGetProp 就是
// `calloc(1, sizeof(VARIANT))`), 而 `sizeof(vb6_VARIANT)` 在 x86 下是 **24**、
// Windows VARIANT 只有 **16** ⇒ 按 vb6_VARIANT 去 memset 会**越界 8 字节**, 堆当场被
// 写坏。症状是"崩溃点到处漂移"(同一份用例这次崩在 MO2、下次崩在 MO7), 极难归因。
// x64 下两者恰好都是 24 ⇒ 只在 32 位暴露 —— 这类"只在 x86 炸"的根因基本都在宽度假设上。
static void memSetEmpty(VARIANT* o) { VariantInit(o); }
static void memSetI4(VARIANT* o, int32_t v) {
    VariantInit(o); o->vt = VT_I4; o->lVal = v;
}
static void memSetStr(VARIANT* o, const wchar_t* s) {
    VariantInit(o); o->vt = VT_BSTR;
    o->bstrVal = SysAllocString(s ? s : L"");
}
static void memSetObj(VARIANT* o, void* iface) {
    VariantInit(o); o->vt = VT_DISPATCH; o->pdispVal = (void*)iface;
}

// 参数里取整数 (下标)。自己写而不用外部 helper: 这里只吃"整数族 + 数字串"。
static int32_t memVariantToI4(VARIANT* v) {
    if (!v) return 0;
    switch (v->vt) {
    case VT_I4: case VT_INT: case VT_UI4: case VT_UINT: return (int32_t)v->lVal;
    case VT_I2: return (int32_t)v->iVal;
    case VT_UI1: return (int32_t)v->bVal;
    case VT_I8: return (int32_t)v->llVal;
    case VT_UI8: return (int32_t)v->ullVal;   /* ai/032 §6.6: 无符号 64 位档同格 */
    case VT_BOOL: return (v->boolVal == VARIANT_TRUE) ? -1 : 0;
    case VT_R8: return (int32_t)v->dblVal;
    case VT_R4: return (int32_t)v->fltVal;
    case VT_BSTR: return v->bstrVal ? (int32_t)wcstol(v->bstrVal, NULL, 10) : 0;
    default: return 0;
    }
}

// ===================== 构造 =====================
static Vb6MemObj* memObjNew(int32_t kind, void* owner, int32_t index) {
    Vb6MemObj* p = (Vb6MemObj*)calloc(1, sizeof(Vb6MemObj));
    if (!p) return NULL;
    p->lpVtbl = &g_memObjVtbl;
    p->ref = 1;
    p->kind = kind;
    p->owner = owner;
    p->index = index;
    if (memTrace()) fprintf(stderr, "[MOBJ] new obj   %p kind=%d idx=%d\n", (void*)p, (int)kind, (int)index);
    return p;
}

static Vb6MemColl* memCollNew(int32_t kind, void* owner) {
    Vb6MemColl* p = (Vb6MemColl*)calloc(1, sizeof(Vb6MemColl));
    if (!p) return NULL;
    p->lpVtbl = &g_memCollVtbl;
    p->ref = 1;
    p->kind = kind;
    p->owner = owner;
    if (memTrace()) fprintf(stderr, "[MOBJ] new coll  %p kind=%d\n", (void*)p, (int)kind);
    return p;
}

// 集合当前的元素个数 (枚举快照 + Count 共用)
static int32_t memCollCount(int32_t kind, void* owner) {
    switch (kind) {
    case VB6_MEMCK_LISTITEMS:     return vb6_ListView_GetItemCount(owner);
    case VB6_MEMCK_COLUMNHEADERS: return vb6_ListView_GetColumnCount(owner);
    case VB6_MEMCK_PANELS:        return vb6_StatusBar_GetPanelsCount(owner);
    case VB6_MEMCK_NODES:       return vb6_TreeView_NodeCount(owner);
    case VB6_MEMCK_BUTTONS:    return vb6_Toolbar_ButtonCount(owner);
    default:                      return vb6_GetImageListCount(owner);
    }
}

static Vb6MemEnum* memEnumNew(int32_t kind, void* owner) {
    Vb6MemEnum* p = (Vb6MemEnum*)calloc(1, sizeof(Vb6MemEnum));
    if (!p) return NULL;
    p->lpVtbl = &g_memEnumVtbl;
    p->ref = 1;
    p->kind = kind;
    p->owner = owner;
    p->pos = 1;
    // count 在**开始枚举那一刻**快照: VB6 的 For Each 也是在进入循环时定下集合的
    // 成员 (循环体里 Remove 不会让枚举器跳过/越界)。
    p->total = memCollCount(kind, owner);
    if (memTrace()) fprintf(stderr, "[MOBJ] new enum  %p kind=%d total=%d\n", (void*)p, (int)kind, (int)p->total);
    return p;
}

// ===================== IUnknown (三个 vtable 共用一套实现) =====================
static HRESULT STDMETHODCALLTYPE mem_QI(IUnknown* This, REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IDispatch)) {
        *ppv = This; This->lpVtbl->AddRef(This); return S_OK;
    }
    // 枚举器另认 IEnumVARIANT —— For Each 那条路就是这么拿到的。
    if (IsEqualIID(riid, &IID_IEnumVARIANT)) {
        // 只有枚举器实例能给出 IEnumVARIANT; 其余返回 E_NOINTERFACE。
        if (This->lpVtbl == (IUnknownVtbl*)&g_memEnumVtbl) {
            *ppv = This; This->lpVtbl->AddRef(This); return S_OK;
        }
    }
    return E_NOINTERFACE;
}
static ULONG STDMETHODCALLTYPE mem_AddRef(IUnknown* This) {
    return (ULONG)InterlockedIncrement(&((Vb6MemObj*)This)->ref);
}
static ULONG STDMETHODCALLTYPE mem_Release(IUnknown* This) {
    LONG r = InterlockedDecrement(&((Vb6MemObj*)This)->ref);
    if (memTrace()) fprintf(stderr, "[MOBJ] release %p -> ref=%ld\n", (void*)This, (long)r);
    if (r <= 0) free(This);
    return (ULONG)r;
}

// ===================== GetTypeInfo 三件套 =====================
static HRESULT STDMETHODCALLTYPE mem_GetTypeInfoCount(IDispatch* This, UINT* p) {
    (void)This; if (p) *p = 0; return S_OK;
}
static HRESULT STDMETHODCALLTYPE mem_GetTypeInfo(IDispatch* This, UINT i, LCID l, ITypeInfo** p) {
    (void)This; (void)i; (void)l; if (p) *p = NULL; return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE memObj_GetIDsOfNames(IDispatch* This, REFIID riid, LPOLESTR* names,
                                                     UINT cNames, LCID lcid, DISPID* dispids) {
    (void)riid; (void)lcid;
    if (!names || !dispids) return E_POINTER;
    const wchar_t* const* tbl = memObjNamesOf(((Vb6MemObj*)This)->kind);
    for (UINT i = 0; i < cNames; i++) {
        int d = memLookup(tbl, names[i]);
        if (memTrace()) fprintf(stderr, "[MOBJ] obj.gidn this=%p '%ls' -> %d\n",
                                (void*)This, names[i] ? names[i] : L"?", d);
        dispids[i] = (DISPID)(d ? d : DISPID_UNKNOWN);
    }
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE memColl_GetIDsOfNames(IDispatch* This, REFIID riid, LPOLESTR* names,
                                                      UINT cNames, LCID lcid, DISPID* dispids) {
    (void)riid; (void)lcid;
    if (!names || !dispids) return E_POINTER;
    const wchar_t* const* tbl = memCollNamesOf(((Vb6MemColl*)This)->kind);
    for (UINT i = 0; i < cNames; i++) {
        // `_NewEnum` 要走 OLE 的固定 DISPID -4 (vb6_ForEach_Init 是硬发 -4 的),
        // 不能给名字表里的序号 —— 给了它 Invoke 也认 -4, 但别的调用方 (编号路径)
        // 拿到 6 会当成第 6 个成员, 两边对不上。
        if (names[i] && _wcsicmp(names[i], L"_NewEnum") == 0) {
            dispids[i] = (DISPID)VB6_DISPID_NEWENUM;
            continue;
        }
        int d = memLookup(tbl, names[i]);
        dispids[i] = (DISPID)(d ? d : DISPID_UNKNOWN);
    }
    return S_OK;
}

// ===================== Invoke 参数工具 =====================
// vb6_ComCall / vb6_ComGetPropArg 把**位置参数逆序**放: rgvarg[cArgs-1-i] 才是第 i 个实参。
static VARIANT* memArg(DISPPARAMS* dp, int pos) {
    if (!dp || !dp->rgvarg || pos < 0 || pos >= (int)dp->cArgs) return NULL;
    return &dp->rgvarg[dp->cArgs - 1 - pos];
}

static int32_t memArgIsMissing(VARIANT* v) {
    return (!v || v->vt == VT_ERROR || v->vt == VT_EMPTY);
}

static const wchar_t* memArgStr(VARIANT* v) {
    if (!v) return NULL;
    if (v->vt == VT_BSTR) return v->bstrVal;
    return NULL;
}

// 实参当成"下标或 Key"解析。
// 返回: >0 = 1 基下标; -1 = 给的是字符串 (key 走 outKey); 0 = 没给/不认识。
static int32_t memParseIndexOrKey(VARIANT* v, const wchar_t** outKey) {
    if (outKey) *outKey = NULL;
    if (memArgIsMissing(v)) return 0;
    if (v->vt == VT_BSTR) {
        if (outKey) *outKey = v->bstrVal;
        return -1;
    }
    return memVariantToI4(v);
}

// ===================== 成员对象: 分 kind 的取值 =====================
static HRESULT memInvokeListImage(Vb6MemObj* p, int dispid, VARIANT* out) {
    switch (dispid) {
    case VB6_MEMD_KEY:
        memSetStr(out, (const wchar_t*)vb6_GetImageListKeyAt(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_INDEX:
        memSetI4(out, p->index);
        return S_OK;
    case VB6_MEMD_PICTURE: {
        void* pic = vb6_ImageList_PictureAt(p->owner, p->index);
        if (pic) memSetObj(out, pic);
        return S_OK;
    }
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

static HRESULT memInvokeListItem(Vb6MemObj* p, int dispid, VARIANT* out, DISPPARAMS* dp) {
    switch (dispid) {
    case VB6_MEMD_KEY:
        memSetStr(out, (const wchar_t*)vb6_ListView_GetItemKey(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_INDEX:
        memSetI4(out, p->index);
        return S_OK;
    case VB6_MEMD_TEXT:
        memSetStr(out, (const wchar_t*)vb6_ListView_GetItemText(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_SUBITEMS: {
        // 参数化属性读: `itm.SubItems(1)` → args=[1] (逆序, 单参就是 rgvarg[0])
        int32_t sub = memParseIndexOrKey(memArg(dp, 0), NULL);
        if (sub < 1) { memSetEmpty(out); return S_OK; }
        memSetStr(out, (const wchar_t*)vb6_ListView_GetItemSub(p->owner, p->index, sub));
        return S_OK;
    }
    case VB6_MEMD_SELECTED:
        memSetI4(out, vb6_ListView_GetItemSelected(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_CHECKED:
        memSetI4(out, vb6_ListView_GetItemChecked(p->owner, p->index));
        return S_OK;
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

static HRESULT memInvokeColumnHeader(Vb6MemObj* p, int dispid, VARIANT* out) {
    switch (dispid) {
    case VB6_MEMD_KEY:
        memSetStr(out, (const wchar_t*)vb6_ListView_GetColumnKey(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_INDEX:
        memSetI4(out, p->index);
        return S_OK;
    case VB6_MEMD_TEXT:
        memSetStr(out, (const wchar_t*)vb6_ListView_GetColumnText(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_WIDTH:
        memSetI4(out, vb6_ListView_GetColumnWidth(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_ALIGNMENT:
        memSetI4(out, vb6_ListView_GetColumnAlign(p->owner, p->index));
        return S_OK;
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// C29-4: StatusBar 的 Panel (事件参数)。数据面照 vb6forms_statusbar.c 的 getter/setter。
static HRESULT memInvokePanel(Vb6MemObj* p, int dispid, VARIANT* out) {
    switch (dispid) {
    case VB6_MEMD_KEY:
        memSetStr(out, (const wchar_t*)vb6_StatusBar_GetPanelKey(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_INDEX:
        // 账 #299: 以前这里是 inline 交 `p->index` —— 同一个事实的**第三份答案**
        // (另两份: 发码侧 recognizer 的 `vb6_StatusBar_GetPanelIndex`, 与存储层的
        // `SbAt`)。改成问那唯一一枚 getter: 面板已被 Remove 时它答 0, 而 inline 那
        // 一份会照样交出旧下标。
        memSetI4(out, vb6_StatusBar_GetPanelIndex(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_TEXT:
        memSetStr(out, (const wchar_t*)vb6_StatusBar_GetPanelText(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_WIDTH:
        memSetI4(out, vb6_StatusBar_GetPanelWidth(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_MINWIDTH:
        memSetI4(out, vb6_StatusBar_GetPanelMinWidth(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_AUTOSIZE:
        memSetI4(out, vb6_StatusBar_GetPanelAutoSize(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_STYLE:
        memSetI4(out, vb6_StatusBar_GetPanelStyle(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_TOOLTIP:
        memSetStr(out, (const wchar_t*)vb6_StatusBar_GetPanelToolTip(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_PANEL_LEFT:
        memSetI4(out, vb6_StatusBar_GetPanelLeft(p->owner, p->index));
        return S_OK;
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// C29-8b: TreeView 的 Node。导航读数 (Parent/Child/Children/Next/Previous/Root) 问的是
// **原生树**, 字符串/勾选/展开问的是 treeview.c 那张表与原生 state 位 —— 没有第二份真相。
static HRESULT memInvokeNode(Vb6MemObj* p, int dispid, VARIANT* out) {
    switch (dispid) {
    case VB6_MEMD_KEY:
        memSetStr(out, vb6_TreeView_GetNodeKey(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_INDEX:
        memSetI4(out, p->index);
        return S_OK;
    case VB6_MEMD_TEXT:
        memSetStr(out, vb6_TreeView_GetNodeText(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_TAG:
        memSetStr(out, vb6_TreeView_GetNodeTag(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_CHECKED:
        memSetI4(out, vb6_TreeView_GetNodeChecked(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_EXPANDED:
        memSetI4(out, vb6_TreeView_GetNodeExpanded(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_PARENT:
        memSetI4(out, vb6_TreeView_GetNodeParent(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_CHILD:
        memSetI4(out, vb6_TreeView_GetNodeChild(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_CHILDREN:
        memSetI4(out, vb6_TreeView_GetNodeChildren(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_NEXT:
        memSetI4(out, vb6_TreeView_GetNodeNext(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_PREV:
        memSetI4(out, vb6_TreeView_GetNodePrev(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_ROOT:
        memSetI4(out, vb6_TreeView_GetNodeRoot(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_NODE_ENSVIS:
        vb6_TreeView_NodeEnsureVisible(p->owner, p->index);
        memSetEmpty(out);
        return S_OK;
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// C29-5b: Toolbar 的 Button。Caption / Image / Enabled / Visible / Value 现问控件
// (TB_GETBUTTONINFOW)，Key / Tag / ToolTipText / Style / Width 问 5a 那张表 ——
// 分界线是"原生答不答得了"，两边都不是第二份真相的复制品。
static HRESULT memInvokeButton(Vb6MemObj* p, int dispid, VARIANT* out) {
    switch (dispid) {
    case VB6_MEMD_KEY:
        memSetStr(out, vb6_Toolbar_GetButtonKey(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_INDEX:
        memSetI4(out, p->index);
        return S_OK;
    case VB6_MEMD_BTN_CAPTION:
        memSetStr(out, vb6_Toolbar_GetButtonCaption(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_TAG:
        memSetStr(out, vb6_Toolbar_GetButtonTag(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_TIP:
        memSetStr(out, vb6_Toolbar_GetButtonToolTip(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_STYLE:
        memSetI4(out, vb6_Toolbar_GetButtonStyle(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_IMAGE:
        memSetI4(out, vb6_Toolbar_GetButtonImage(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_ENABLED:
        memSetI4(out, vb6_Toolbar_GetButtonEnabled(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_VISIBLE:
        memSetI4(out, vb6_Toolbar_GetButtonVisible(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_WIDTH:
        memSetI4(out, vb6_Toolbar_GetButtonWidth(p->owner, p->index));
        return S_OK;
    case VB6_MEMD_BTN_VALUE:
        memSetI4(out, vb6_Toolbar_GetButtonValue(p->owner, p->index));
        return S_OK;
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// C29-Data: Data.Recordset 的 Invoke。带参的 Fields("x") 在 dp 里 (PROPERTYGET 逆序)。
static HRESULT memInvokeRs(Vb6MemObj* p, int dispid, VARIANT* out, DISPPARAMS* dp, WORD flags) {
    if (memTrace()) fprintf(stderr, "[MOBJ-RS] dispid=%d flags=%04X owner=%p\n", dispid, (unsigned)flags, p->owner);
    switch (dispid) {
    case VB6_MEMRS_BOF:
        memSetI4(out, vb6_Data_BOF(p->owner));
        return S_OK;
    case VB6_MEMRS_EOF:
        memSetI4(out, vb6_Data_EOF(p->owner));
        return S_OK;
    case VB6_MEMRS_RECORDCNT:
        memSetI4(out, vb6_Data_RecordCount(p->owner));
        return S_OK;
    case VB6_MEMRS_FIELDCNT:
        memSetI4(out, vb6_Data_FieldCount(p->owner));
        return S_OK;
    case VB6_MEMRS_CURROW:
        memSetI4(out, vb6_Data_CurrentRow(p->owner));
        return S_OK;
    case VB6_MEMRS_FIELDS: {
        // PROPERTYGET 带参: 参数在 dp->rgvarg 逆序 (rgvarg[cArgs-1] 是第一个实参)
        int32_t argc = dp ? (int32_t)dp->cArgs : 0;
        if (argc < 1) { memSetEmpty(out); return DISP_E_PARAMNOTFOUND; }
        VARIANT* a0 = &dp->rgvarg[argc - 1];
        // 字段名/下标 —— 找到后产 **Field 对象** (VB6: Set f = rs.Fields("id"); f.Value)
        const wchar_t* fname = memArgStr(a0);
        int32_t fidx = 0;
        int32_t allDigit = fname[0] != 0;
        for (const wchar_t* q = fname; *q; q++)
            if (!(*q >= L'0' && *q <= L'9')) { allDigit = 0; break; }
        if (allDigit) fidx = (int32_t)wcstol(fname, NULL, 10);
        else {
            for (int c = 1; c <= vb6_Data_FieldCount(p->owner); c++) {
                wchar_t tmp[64];
                vb6_Data_FieldValue(p->owner, L"", tmp, 0);  /* noop, 占位 */
                (void)tmp;
                if (_wcsicmp(vb6_Data_FieldName(p->owner, c), fname) == 0) { fidx = c; break; }
            }
        }
        if (fidx < 1) { memSetEmpty(out); return DISP_E_MEMBERNOTFOUND; }
        Vb6MemObj* f = memObjNew(VB6_MEMK_FIELD, p->owner, fidx);
        if (!f) { memSetEmpty(out); return E_OUTOFMEMORY; }
        memSetObj(out, f);
        return S_OK;
    }
    case VB6_MEMRS_MOVEFIRST: vb6_Data_MoveFirst(p->owner);    memSetEmpty(out); return S_OK;
    case VB6_MEMRS_MOVELAST:  vb6_Data_MoveLast(p->owner);     memSetEmpty(out); return S_OK;
    case VB6_MEMRS_MOVENEXT:  vb6_Data_MoveNext(p->owner);     memSetEmpty(out); return S_OK;
    case VB6_MEMRS_MOVEPREV:  vb6_Data_MovePrevious(p->owner); memSetEmpty(out); return S_OK;
    case VB6_MEMRS_REFRESH:   if (memTrace()) fprintf(stderr, "[MOBJ-RS] calling vb6_Data_Refresh owner=%p\n", p->owner);
                              vb6_Data_Refresh(p->owner);      memSetEmpty(out); return S_OK;
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// C29-Data: Field (Name/Value/Index)
static HRESULT memInvokeRsField(Vb6MemObj* p, int dispid, VARIANT* out) {
    switch (dispid) {
    case VB6_MEMF_NAME:
        memSetStr(out, vb6_Data_FieldName(p->owner, p->index));
        return S_OK;
    case VB6_MEMF_VALUE: {
        wchar_t vbuf[1024];
        vb6_Data_FieldValueByIdx(p->owner, p->index, vbuf, 1024);
        if (memTrace()) fprintf(stderr, "[MOBJ-F] idx=%d val=[%ls]\n", p->index, vbuf);
        memSetStr(out, vbuf);
        return S_OK;
    }
    case VB6_MEMF_INDEX:
        memSetI4(out, p->index);
        return S_OK;
    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// 成员对象的**属性写** (只给可写的成员: ListItem.Text/Selected/Checked、ColumnHeader.Text/Width/Alignment)
static void memObjPutProp(Vb6MemObj* p, int dispid, DISPPARAMS* dp) {
    if (!dp || dp->cArgs < 1) return;
    VARIANT* v = &dp->rgvarg[0];   // PUT 时新值在 rgvarg[0] (与 CALL 的逆序不同)
    if (p->kind == VB6_MEMK_LISTITEM) {
        switch (dispid) {
        case VB6_MEMD_TEXT:
            vb6_ListView_SetItemText(p->owner, p->index, (void*)memArgStr(v));
            break;
        case VB6_MEMD_SELECTED:
            vb6_ListView_SetItemSelected(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_CHECKED:
            vb6_ListView_SetItemChecked(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_KEY:
            vb6_ListView_SetItemKey(p->owner, p->index, (void*)memArgStr(v));
            break;
        case VB6_MEMD_SUBITEMS: {
            // **参数化属性的写**: `itm.SubItems(1) = "销售部"`。
            // vb6_ComSetPropArg 的布局是 rgvarg[0]=新值、rgvarg[1..]=索引(**正序**),
            // 与 CALL/GET 的逆序不同 —— 抄错了会"写进了别的下标"而不是报错。
            // (读那条在 memInvokeListItem 里, 用的是 memArg() 的逆序取值。)
            VARIANT* vIdx = (dp && dp->cArgs >= 2) ? &dp->rgvarg[1] : NULL;
            int32_t sub = vIdx ? memVariantToI4(vIdx) : 0;
            if (sub >= 1)
                vb6_ListView_SetItemSub(p->owner, p->index, sub, (void*)memArgStr(v));
            break;
        }
        default: break;
        }
        return;
    }
    if (p->kind == VB6_MEMK_NODE) {
        switch (dispid) {
        case VB6_MEMD_TEXT:
            vb6_TreeView_SetNodeText(p->owner, p->index, memArgStr(v));
            break;
        case VB6_MEMD_KEY:
            vb6_TreeView_SetNodeKey(p->owner, p->index, memArgStr(v));
            break;
        case VB6_MEMD_NODE_TAG:
            vb6_TreeView_SetNodeTag(p->owner, p->index, memArgStr(v));
            break;
        case VB6_MEMD_NODE_CHECKED:
            vb6_TreeView_SetNodeChecked(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_NODE_EXPANDED:
            vb6_TreeView_SetNodeExpanded(p->owner, p->index, memVariantToI4(v));
            break;
        default: break;
        }
        return;
    }
    if (p->kind == VB6_MEMK_BUTTON) {
        switch (dispid) {
        case VB6_MEMD_BTN_CAPTION:
            vb6_Toolbar_SetButtonCaption(p->owner, p->index, memArgStr(v)); break;
        case VB6_MEMD_BTN_TAG:
            vb6_Toolbar_SetButtonTag(p->owner, p->index, memArgStr(v)); break;
        case VB6_MEMD_BTN_TIP:
            vb6_Toolbar_SetButtonToolTip(p->owner, p->index, memArgStr(v)); break;
        case VB6_MEMD_KEY:
            vb6_Toolbar_SetButtonKey(p->owner, p->index, memArgStr(v)); break;
        case VB6_MEMD_BTN_STYLE:
        case VB6_MEMD_BTN_IMAGE:
        case VB6_MEMD_BTN_ENABLED:
        case VB6_MEMD_BTN_VISIBLE:
        case VB6_MEMD_BTN_WIDTH:
        case VB6_MEMD_BTN_VALUE: {
            int32_t iv = memVariantToI4(v);
            if (dispid == VB6_MEMD_BTN_STYLE)  vb6_Toolbar_SetButtonStyle(p->owner, p->index, iv);
            else if (dispid == VB6_MEMD_BTN_IMAGE)   vb6_Toolbar_SetButtonImage(p->owner, p->index, iv);
            else if (dispid == VB6_MEMD_BTN_ENABLED) vb6_Toolbar_SetButtonEnabled(p->owner, p->index, iv);
            else if (dispid == VB6_MEMD_BTN_VISIBLE) vb6_Toolbar_SetButtonVisible(p->owner, p->index, iv);
            else if (dispid == VB6_MEMD_BTN_WIDTH)   vb6_Toolbar_SetButtonWidth(p->owner, p->index, iv);
            else                                     vb6_Toolbar_SetButtonValue(p->owner, p->index, iv);
            break;
        }
        default: break;
        }
        return;
    }
    if (p->kind == VB6_MEMK_PANEL) {
        switch (dispid) {
        case VB6_MEMD_TEXT:
            vb6_StatusBar_SetPanelText(p->owner, p->index, (void*)memArgStr(v));
            break;
        case VB6_MEMD_WIDTH:
            vb6_StatusBar_SetPanelWidth(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_MINWIDTH:
            vb6_StatusBar_SetPanelMinWidth(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_AUTOSIZE:
            vb6_StatusBar_SetPanelAutoSize(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_STYLE:
            vb6_StatusBar_SetPanelStyle(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_KEY:
            vb6_StatusBar_SetPanelKey(p->owner, p->index, (void*)memArgStr(v));
            break;
        case VB6_MEMD_TOOLTIP:
            vb6_StatusBar_SetPanelToolTip(p->owner, p->index, (void*)memArgStr(v));
            break;
        default: break;
        }
        return;
    }
    if (p->kind == VB6_MEMK_COLUMNHEADER) {
        switch (dispid) {
        case VB6_MEMD_TEXT:
            vb6_ListView_SetColumnText(p->owner, p->index, (void*)memArgStr(v));
            break;
        case VB6_MEMD_WIDTH:
            vb6_ListView_SetColumnWidth(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_ALIGNMENT:
            vb6_ListView_SetColumnAlign(p->owner, p->index, memVariantToI4(v));
            break;
        case VB6_MEMD_KEY: {
            // ColumnHeader 没有 SetColumnKey (键是插入时定的)。**不静默吞掉**:
            // 留一条 stderr 说明, 免得"赋值了却没生效"被当成玄学。
            fprintf(stderr, "[C3] ColumnHeader.Key is read-only in this build\n");
            break;
        }
        default: break;
        }
        return;
    }
}

static HRESULT STDMETHODCALLTYPE memObj_Invoke(IDispatch* This, DISPID dispid, REFIID riid, LCID lcid,
                                               WORD flags, DISPPARAMS* dp, VARIANT* result,
                                               EXCEPINFO* ei, UINT* ae) {
    (void)riid; (void)lcid; (void)ei; (void)ae;
    Vb6MemObj* p = (Vb6MemObj*)This;
    if (memTrace()) fprintf(stderr, "[MOBJ] obj.invoke this=%p kind=%d dispid=%d idx=%d flags=0x%04X\n",
                            (void*)This, (int)p->kind, (int)dispid, (int)p->index, (unsigned)flags);

    // 属性写 (PROPERTYPUT / PROPERTYPUTREF) —— vb6_ComSetPropArg 的 result 传 NULL,
    // 所以这一支必须排在 `!result` 检查**之前**。
    if (flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) {
        memObjPutProp(p, (int)dispid, dp);
        return S_OK;
    }

    if (!result) return S_OK;
    VARIANT* out = result;
    memSetEmpty(out);
    switch (p->kind) {
    case VB6_MEMK_LISTIMAGE:    return memInvokeListImage(p, (int)dispid, out);
    case VB6_MEMK_LISTITEM:     return memInvokeListItem(p, (int)dispid, out, dp);
    case VB6_MEMK_COLUMNHEADER: return memInvokeColumnHeader(p, (int)dispid, out);
    case VB6_MEMK_PANEL:        return memInvokePanel(p, (int)dispid, out);
    case VB6_MEMK_NODE:        return memInvokeNode(p, (int)dispid, out);
    case VB6_MEMK_BUTTON:      return memInvokeButton(p, (int)dispid, out);
    case VB6_MEMK_RS:           return memInvokeRs(p, (int)dispid, out, dp, flags);
    case VB6_MEMK_FIELD:        return memInvokeRsField(p, (int)dispid, out);
    default: break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// ===================== 成员集合: 分 kind 的 Add =====================
// 各族 Add 的实参 (VB6 原文):
//   ListImages.Add   (index, key, picture)
//   ListItems.Add    (index, key, text, iconIdx, smalliconIdx)
//   ColumnHeaders.Add(index, key, text, width, alignment, iconIdx)
static int32_t memCollAdd(int32_t kind, void* owner, DISPPARAMS* dp, void** outObj) {
    *outObj = NULL;
    VARIANT* a0 = memArg(dp, 0);   // index
    VARIANT* a1 = memArg(dp, 1);   // key
    VARIANT* a2 = memArg(dp, 2);   // text / picture
    VARIANT* a3 = memArg(dp, 3);
    VARIANT* a4 = memArg(dp, 4);

    int32_t idx = 0;
    if (!memArgIsMissing(a0)) idx = memVariantToI4(a0);
    const wchar_t* key = memArgStr(a1);

    switch (kind) {
    case VB6_MEMCK_LISTIMAGES: {
        void* pic = NULL;
        if (!memArgIsMissing(a2)) {
            if (a2->vt == VT_DISPATCH || a2->vt == VT_UNKNOWN || a2->vt == VT_BYREF)
                pic = a2->pdispVal;
        }
        int32_t newIdx = vb6_ImageList_AddPicture(owner, idx, key, pic);
        if (newIdx <= 0) return 0;
        Vb6MemObj* o = memObjNew(VB6_MEMK_LISTIMAGE, owner, newIdx);
        if (!o) return 0;
        *outObj = o;
        return newIdx;
    }
    case VB6_MEMCK_LISTITEMS: {
        const wchar_t* text = memArgStr(a2);
        int32_t iconIdx = memArgIsMissing(a3) ? 0 : memVariantToI4(a3);
        // ⚠ 别把局部变量叫 `small`: Windows SDK 的 rpcndr.h 里有 `#define small char`
        // (老 MIDL 的遗留), 于是 `int32_t small = …` 会被展开成 `int32_t char = …` →
        // C2628 "int32_t 后面接 char 是非法的" —— 报错行还会指向上/下一行, 极难认。
        // 同族的宏还有 `hyper`(→__int64)。`smallIcon` 这类**完整标识符**不受影响。
        int32_t smallIdx = memArgIsMissing(a4) ? 0 : memVariantToI4(a4);
        int32_t newIdx = vb6_ListView_AddItem(owner, idx, (void*)key, (void*)text, iconIdx, smallIdx);
        if (newIdx <= 0) return 0;
        Vb6MemObj* o = memObjNew(VB6_MEMK_LISTITEM, owner, newIdx);
        if (!o) return 0;
        *outObj = o;
        return newIdx;
    }
    case VB6_MEMCK_COLUMNHEADERS: {
        const wchar_t* text = memArgStr(a2);
        int32_t width = memArgIsMissing(a3) ? 0 : memVariantToI4(a3);
        int32_t align = memArgIsMissing(a4) ? 0 : memVariantToI4(a4);
        int32_t newIdx = vb6_ListView_AddColumn(owner, idx, (void*)key, (void*)text, width, align);
        if (newIdx <= 0) return 0;
        Vb6MemObj* o = memObjNew(VB6_MEMK_COLUMNHEADER, owner, newIdx);
        if (!o) return 0;
        *outObj = o;
        return newIdx;
    }
    case VB6_MEMCK_NODES: {
        // Nodes.Add([relative], [relationship], [key], [text], [image], [selectedimage])
        // relative 可以是下标也可以是 Key (VB6 两者都收)；Key 先换成下标再往下传。
        const wchar_t* rKey = NULL;
        int32_t rel = memParseIndexOrKey(memArg(dp, 0), &rKey);
        if (rel < 0) rel = (rKey && *rKey) ? vb6_TreeView_NodeIndexByKey(owner, rKey) : 0;
        int32_t relship = memArgIsMissing(a1) ? 4 /* tvwChild */ : memVariantToI4(a1);
        int32_t newIdx = vb6_TreeView_AddNode(owner, rel, relship, memArgStr(a2), memArgStr(a3),
                                            memArgIsMissing(a4) ? 0 : memVariantToI4(a4),
                                            memArgIsMissing(memArg(dp, 5)) ? 0 : memVariantToI4(memArg(dp, 5)));
        if (newIdx <= 0) return 0;
        Vb6MemObj* o = memObjNew(VB6_MEMK_NODE, owner, newIdx);
        if (!o) return 0;
        *outObj = o;
        return newIdx;
    }
    case VB6_MEMCK_BUTTONS: {
        int32_t newIdx = vb6_Toolbar_AddButton(owner, idx, key, memArgStr(a2),
                                            memArgIsMissing(a3) ? 0 : memVariantToI4(a3),
                                            memArgIsMissing(a4) ? -1 : memVariantToI4(a4),
                                            NULL, 0);
        if (newIdx <= 0) return 0;
        Vb6MemObj* o = memObjNew(VB6_MEMK_BUTTON, owner, newIdx);
        if (!o) return 0;
        *outObj = o;
        return newIdx;
    }
    default:
        break;
    }
    return 0;
}

static void memCollRemove(int32_t kind, void* owner, DISPPARAMS* dp) {
    const wchar_t* key = NULL;
    int32_t r = memParseIndexOrKey(memArg(dp, 0), &key);
    int32_t idx = r;
    if (r < 0 && key && *key) {
        // 按 Key: 各族查表换算成 1 基下标
        if (kind == VB6_MEMCK_LISTITEMS)
            idx = vb6_ListView_GetItemIndexByKey(owner, (void*)key);
        else if (kind == VB6_MEMCK_LISTIMAGES)
            idx = vb6_ImageListIndexByKey(owner, key);
        else if (kind == VB6_MEMCK_NODES)
            idx = vb6_TreeView_NodeIndexByKey(owner, key);
        else if (kind == VB6_MEMCK_BUTTONS)
            idx = vb6_Toolbar_ButtonIndexByKey(owner, key);
        else
            idx = 0;   // ColumnHeader 没有按 Key 取 (扁平层没这条)
    }
    if (idx <= 0) return;
    switch (kind) {
    case VB6_MEMCK_LISTITEMS:     vb6_ListView_RemoveItem(owner, idx); break;
    case VB6_MEMCK_COLUMNHEADERS: vb6_ListView_RemoveColumn(owner, idx); break;
    case VB6_MEMCK_NODES:         vb6_TreeView_RemoveNode(owner, idx); break;
    case VB6_MEMCK_BUTTONS:     vb6_Toolbar_RemoveButton(owner, idx); break;
    case VB6_MEMCK_PANELS:        break;  /* P20-40 已有特例直译; 这里不重复 */
    default:                      vb6_ImageList_RemoveAtIndex(owner, idx); break;
    }
}

static void memCollClear(int32_t kind, void* owner) {
    switch (kind) {
    case VB6_MEMCK_LISTITEMS:     vb6_ListView_ClearItems(owner); break;
    case VB6_MEMCK_COLUMNHEADERS: vb6_ListView_ClearColumns(owner); break;
    case VB6_MEMCK_PANELS:        vb6_StatusBar_ClearPanels(owner); break;
    case VB6_MEMCK_NODES:       vb6_TreeView_ClearNodes(owner); break;
    case VB6_MEMCK_BUTTONS:   vb6_Toolbar_ClearButtons(owner); break;
    default:                      vb6_ImageList_ClearImages(owner); break;
    }
}

static int32_t memObjKindOfColl(int32_t collKind) {
    switch (collKind) {
    case VB6_MEMCK_LISTITEMS:     return VB6_MEMK_LISTITEM;
    case VB6_MEMCK_COLUMNHEADERS: return VB6_MEMK_COLUMNHEADER;
    case VB6_MEMCK_PANELS:        return VB6_MEMK_PANEL;
    case VB6_MEMCK_NODES:       return VB6_MEMK_NODE;
    case VB6_MEMCK_BUTTONS:   return VB6_MEMK_BUTTON;
    default:                      return VB6_MEMK_LISTIMAGE;
    }
}

static HRESULT STDMETHODCALLTYPE memColl_Invoke(IDispatch* This, DISPID dispid, REFIID riid, LCID lcid,
                                                WORD flags, DISPPARAMS* dp, VARIANT* result,
                                                EXCEPINFO* ei, UINT* ae) {
    (void)riid; (void)lcid; (void)flags; (void)ei; (void)ae;
    Vb6MemColl* p = (Vb6MemColl*)This;
    if (!result) return S_OK;
    VARIANT* out = result;
    memSetEmpty(out);

    void* owner = p->owner;

    // _NewEnum: 交给 For Each。硬发 -4。
    if (dispid == (DISPID)VB6_DISPID_NEWENUM || (int)dispid == VB6_MEMCD_NEWENUM) {
        Vb6MemEnum* e = memEnumNew(p->kind, owner);
        if (!e) return E_OUTOFMEMORY;
        memSetObj(out, e);
        return S_OK;
    }

    switch ((int)dispid) {
    case VB6_MEMCD_COUNT:
        memSetI4(out, memCollCount(p->kind, owner));
        return S_OK;

    case VB6_MEMCD_ITEM: {
        // Item(i) / Item("key") —— 越界/找不到返回 Nothing (VB6 也是 Nothing)。
        const wchar_t* key = NULL;
        int32_t r = memParseIndexOrKey(memArg(dp, 0), &key);
        int32_t idx = r;
        if (r < 0 && key && *key) {
            if (p->kind == VB6_MEMCK_LISTITEMS)      idx = vb6_ListView_GetItemIndexByKey(owner, (void*)key);
            else if (p->kind == VB6_MEMCK_LISTIMAGES) idx = vb6_ImageListIndexByKey(owner, key);
            else if (p->kind == VB6_MEMCK_NODES)      idx = vb6_TreeView_NodeIndexByKey(owner, key);
            else if (p->kind == VB6_MEMCK_BUTTONS)  idx = vb6_Toolbar_ButtonIndexByKey(owner, key);
            // 账 #299: 集合对象立起来之后, `Panels("k")` 走的就是这一支。以前名单里没有
            // PANELS ⇒ 掉进 `else idx = 0` ⇒ 交回 Nothing, 而 VB6 交的是那一格面板。
            else if (p->kind == VB6_MEMCK_PANELS)   idx = vb6_StatusBar_GetPanelIndexByKey(owner, key);
            else                                     idx = 0;
        }
        if (idx <= 0) { memSetEmpty(out); return S_OK; }
        Vb6MemObj* o = memObjNew(memObjKindOfColl(p->kind), owner, idx);
        if (!o) return E_OUTOFMEMORY;
        memSetObj(out, o);
        return S_OK;
    }

    case VB6_MEMCD_ADD: {
        void* newObj = NULL;
        memCollAdd(p->kind, owner, dp, &newObj);
        if (newObj) memSetObj(out, newObj);
        return S_OK;
    }

    case VB6_MEMCD_REMOVE:
        memCollRemove(p->kind, owner, dp);
        return S_OK;

    case VB6_MEMCD_CLEAR:
        memCollClear(p->kind, owner);
        return S_OK;

    default:
        break;
    }
    return DISP_E_MEMBERNOTFOUND;
}

// ===================== IEnumVARIANT (集合枚举器) =====================
static HRESULT STDMETHODCALLTYPE memEnum_Next(IEnumVARIANT* This, ULONG celt, VARIANT* rgVar,
                                              ULONG* pCeltFetched) {
    Vb6MemEnum* e = (Vb6MemEnum*)This;
    ULONG got = 0;
    while (got < celt && e->pos <= e->total) {
        Vb6MemObj* o = memObjNew(memObjKindOfColl(e->kind), e->owner, e->pos);
        if (!o) break;
        memSetObj(&rgVar[got], o);
        e->pos++;
        got++;
    }
    if (pCeltFetched) *pCeltFetched = got;
    return (got == celt) ? S_OK : S_FALSE;
}

static HRESULT STDMETHODCALLTYPE memEnum_Skip(IEnumVARIANT* This, ULONG celt) {
    Vb6MemEnum* e = (Vb6MemEnum*)This;
    e->pos += (int32_t)celt;
    if (e->pos > e->total + 1) e->pos = e->total + 1;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE memEnum_Reset(IEnumVARIANT* This) {
    ((Vb6MemEnum*)This)->pos = 1;
    return S_OK;
}

static HRESULT STDMETHODCALLTYPE memEnum_Clone(IEnumVARIANT* This, IEnumVARIANT** ppEnum) {
    Vb6MemEnum* e = (Vb6MemEnum*)This;
    if (!ppEnum) return E_POINTER;
    Vb6MemEnum* c = memEnumNew(e->kind, e->owner);
    if (!c) return E_OUTOFMEMORY;
    c->pos = e->pos;
    c->total = e->total;
    *ppEnum = (IEnumVARIANT*)c;
    return S_OK;
}

// ===================== vtable 装配 =====================
static const IDispatchVtbl g_memObjVtbl = {
    (HRESULT (STDMETHODCALLTYPE*)(IUnknown*, REFIID, void**))mem_QI,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))mem_AddRef,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))mem_Release,
    mem_GetTypeInfoCount,
    mem_GetTypeInfo,
    memObj_GetIDsOfNames,
    memObj_Invoke
};

static const IDispatchVtbl g_memCollVtbl = {
    (HRESULT (STDMETHODCALLTYPE*)(IUnknown*, REFIID, void**))mem_QI,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))mem_AddRef,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))mem_Release,
    mem_GetTypeInfoCount,
    mem_GetTypeInfo,
    memColl_GetIDsOfNames,
    memColl_Invoke
};

static const IEnumVARIANTVtbl g_memEnumVtbl = {
    (HRESULT (STDMETHODCALLTYPE*)(IUnknown*, REFIID, void**))mem_QI,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))mem_AddRef,
    (ULONG (STDMETHODCALLTYPE*)(IUnknown*))mem_Release,
    memEnum_Next,
    memEnum_Skip,
    memEnum_Reset,
    memEnum_Clone
};

// ===================== 对外入口 =====================
// ImageList1.ListImages —— 返回集合对象 (真 IDispatch)。宿主槽是 vb6_com_<Name> 里
// 那个 ImageList 实例指针 (见 vb6forms_imagelist.c 的注释: 无窗口控件, 槽里不是 HWND)。
void* vb6_ImageList_ListImages(void* slot) {
    if (!slot) return NULL;
    return (void*)memCollNew(VB6_MEMCK_LISTIMAGES, slot);
}

// ListImages.Add([index], [key], [picture]) —— **返回 ListImage 对象**。
// comctl32 句柄与平行 key 表那一套仍在 imagelist.c 的扁平实现里, 这里只把返回的
// 1 基下标包成成员对象: VB6 的 Add 返回的就是对象, 不是下标。
void* vb6_ImageList_ListImages_Add(void* slot, int32_t index, const wchar_t* key, void* pic) {
    if (!slot) return NULL;
    int32_t idx = vb6_ImageList_AddPicture(slot, index, key, pic);
    if (idx <= 0) return NULL;
    return (void*)memObjNew(VB6_MEMK_LISTIMAGE, slot, idx);
}

// ListView1.ListItems / .ColumnHeaders —— owner 是 **HWND** (真窗口, 与 ImageList 不同)。
void* vb6_ListView_ListItems(void* hwnd) {
    if (!hwnd) return NULL;
    return (void*)memCollNew(VB6_MEMCK_LISTITEMS, hwnd);
}
void* vb6_ListView_ColumnHeaders(void* hwnd) {
    if (!hwnd) return NULL;
    return (void*)memCollNew(VB6_MEMCK_COLUMNHEADERS, hwnd);
}

// ListItems.Add([index], [key], [text], [iconIdx], [smalliconIdx]) → ListItem 对象
void* vb6_ListView_ListItems_Add(void* hwnd, int32_t index, const wchar_t* key,
                                 const wchar_t* text, int32_t iconIdx, int32_t smallIcon) {
    if (!hwnd) return NULL;
    int32_t idx = vb6_ListView_AddItem(hwnd, index, (void*)key, (void*)text, iconIdx, smallIcon);
    if (idx <= 0) return NULL;
    return (void*)memObjNew(VB6_MEMK_LISTITEM, hwnd, idx);
}

// ColumnHeaders.Add([index], [key], [text], [width], [alignment]) → ColumnHeader 对象
void* vb6_ListView_ColumnHeaders_Add(void* hwnd, int32_t index, const wchar_t* key,
                                     const wchar_t* text, int32_t width, int32_t align) {
    if (!hwnd) return NULL;
    int32_t idx = vb6_ListView_AddColumn(hwnd, index, (void*)key, (void*)text, width, align);
    if (idx <= 0) return NULL;
    return (void*)memObjNew(VB6_MEMK_COLUMNHEADER, hwnd, idx);
}

// 事件派发用: 按 1 基下标取成员对象 —— VB6 的
// `Sub ListView1_ItemClick(ByVal Item As ListItem)` 参数是**对象**,
// 所以 WM_NOTIFY 里换算出下标之后要先把对象造出来再回调。
// 越界返回 NULL (那种情况下事件本就不该发, cgen 侧用 `> 0` 兜住)。
void* vb6_ListView_ListItemAt(void* hwnd, int32_t index) {
    if (!hwnd || index < 1) return NULL;
    if (index > vb6_ListView_GetItemCount(hwnd)) return NULL;
    return (void*)memObjNew(VB6_MEMK_LISTITEM, hwnd, index);
}
void* vb6_ListView_ColumnHeaderAt(void* hwnd, int32_t index) {
    if (!hwnd || index < 1) return NULL;
    if (index > vb6_ListView_GetColumnCount(hwnd)) return NULL;
    return (void*)memObjNew(VB6_MEMK_COLUMNHEADER, hwnd, index);
}

// C29-Data: Data1.Recordset 的**真 IDispatch 对象** —— readFn 直接返回它,
// 之后整条链 (BOF/EOF/RecordCount/Fields("x")/Move*) 走 C29-3 的晚绑定通道, cgen 零特例。
void* vb6_MemberObj_NewRs(void* hwnd) {
    return (void*)memObjNew(VB6_MEMK_RS, hwnd, 0);
}

// C29-4: StatusBar1_PanelClick(ByVal Panel As Panel) 的事件参数对象。
// 1 基下标来自 NMHDR.idFrom (原生 status bar 约定)。
void* vb6_StatusBar_PanelAt(void* hwnd, int32_t index) {
    if (!hwnd || index < 1) return NULL;
    if (index > vb6_StatusBar_GetPanelsCount(hwnd)) return NULL;
    return (void*)memObjNew(VB6_MEMK_PANEL, hwnd, index);
}

// 账 #299: `StatusBar1.Panels` 的**集合对象**。姊妹控件这一枚早就有
// (vb6_TreeView_Nodes / vb6_Toolbar_Buttons / vb6_ListView_ListItems /
// vb6_ImageList_ListImages, 同一处机制), Panels 一族只补了事件参数用的
// PanelAt, 集合本体一直没人造 ⇒ `Set po = SB1.Panels(1)` / 模块级 `As Object`
// / `For Each pv In SB1.Panels` / `With SB1.Panels(1)` 四形的头都掉回
// vb6_ComGetObjectProp(HWND, L"Panels") = 拿 HWND 当 IDispatch 用, 一条读数都
// 交不出(实测 `.build/b351_probe/P300b.frm`)。立起来之后整条链
// (Count/Item(i)/Item("key")/For Each/Panel.八成员) 由这里答, cgen 零特例。
void* vb6_StatusBar_Panels(void* hwnd) {
    if (!hwnd) return NULL;
    return (void*)memCollNew(VB6_MEMCK_PANELS, hwnd);
}

// C29-8b: TreeView1.Nodes —— owner 是 HWND (与 ListView / StatusBar 同)。
// readFn 直接返回集合对象，之后整条链 (Count/Item/Add/Remove/For Each/Node.xxx)
// 走 C29-3 的晚绑定通道，cgen 零特例 —— 与 C29-Data 同一 cheapest route。
void* vb6_TreeView_Nodes(void* hwnd) {
    if (!hwnd) return NULL;
    return (void*)memCollNew(VB6_MEMCK_NODES, hwnd);
}

// 事件参数用: NodeClick(ByVal Node As Node) 要的是对象，不是下标。
void* vb6_TreeView_NodeAt(void* hwnd, int32_t index) {
    if (!hwnd || index < 1) return NULL;
    if (index > vb6_TreeView_NodeCount(hwnd)) return NULL;
    return (void*)memObjNew(VB6_MEMK_NODE, hwnd, index);
}

// C29-5b: Toolbar1.Buttons —— owner 是 HWND，之后整条链走 C29-3 的晚绑定通道
// (与 C29-8b 的 Nodes 同一 cheapest route，cgen 零特例)。
void* vb6_Toolbar_Buttons(void* hwnd) {
    if (!hwnd) return NULL;
    return (void*)memCollNew(VB6_MEMCK_BUTTONS, hwnd);
}

// 事件参数用: ButtonClick(ByVal Button As Button) 要的是对象，不是下标。
void* vb6_Toolbar_ButtonAt(void* hwnd, int32_t index) {
    if (!hwnd || index < 1) return NULL;
    if (index > vb6_Toolbar_ButtonCount(hwnd)) return NULL;
    return (void*)memObjNew(VB6_MEMK_BUTTON, hwnd, index);
}

#endif // _WIN32
