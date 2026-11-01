// vb6forms_treeview.c — ai/029 C29-8a: VB6 TreeView 控件的标量属性面 (原生 SysTreeView32)
//
// 复刻口径 (ai/内置控件/TreeView 控件.md + D6): **不加载 MSCOMCTL.OCX** (它是 32 位 inproc,
//   x64 进程里 CoCreateInstance 直接失败)。SysTreeView32 由 comctl32 注册, 创建那一刀早就在
//   (controlTypeToWin32Class 有格, 实测 vb6_CreateControl 拿到过真句柄), 本文件补的是**属性面**:
//   此前 cgen 的读写表里 TreeView 零格 ⇒ 属性一律落到通用兜底 vb6_ComGetObjectProp(裸 HWND, L"…"),
//   于是 `tv1.CheckBoxes = True` 静默丢、读回空串 (029 §九 C29-8 前置测量那四条读数)。
//
// 本批范围 = 不吃 Nodes 集合的那几格:
//   LineStyle      0 = tvwTreeLines(默认) / 1 = tvwRootLines  → TVS_HASLINES ± TVS_LINESATROOT
//   CheckBoxes     Boolean, VB6 默认 False                   → TVS_CHECKBOXES
//   HotTracking    Boolean, VB6 默认 False                   → TVS_TRACKSELECT
//   HideSelection  Boolean, VB6 默认 True                     → TVS_SHOWSELALWAYS 取反
//   Indentation    缇 (Win32 侧是像素)                        → TVM_SETINDENT / TVM_GETINDENT
//   留 C29-8b (等 C29-3 立完成员对象机制): Style 四态(带图标)、PathSeparator、LabelEdit、
//   Sorted、SelectedItem、Nodes/Node 一族、TVN_* 事件。
//
// 真值放哪: 前四条的**真值就是窗口样式位本身** —— getter 直接读 GWL_STYLE, 读数与观感同源,
//   答不出"表里写着开、窗口上却没框"这种假话。Indentation 例外: VB6 侧存缇、Win32 侧是像素,
//   且 TVM_SETINDENT 会把超界值静默钳掉, 所以缇值另存窗口属性 (照 vb6forms_progress.c 的
//   SetProp +1 偏移手法: GetPropW 对"未设置"与"存了 0"都返回 NULL, 不偏移就把 `= 0` 当成没写)。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"

#ifdef _WIN32

// commctrl.h 在低 _WIN32_IE 下不给这几个样式位, 且 cgen 那侧也按字面量发 —— 两边都钉死数值。
#ifndef TVS_HASBUTTONS
#define TVS_HASBUTTONS     0x0001L
#endif
#ifndef TVS_HASLINES
#define TVS_HASLINES       0x0002L
#endif
#ifndef TVS_LINESATROOT
#define TVS_LINESATROOT    0x0004L
#endif
#ifndef TVS_SHOWSELALWAYS
#define TVS_SHOWSELALWAYS  0x0020L
#endif
#ifndef TVS_CHECKBOXES
#define TVS_CHECKBOXES     0x0100L
#endif
#ifndef TVS_TRACKSELECT
#define TVS_TRACKSELECT    0x0200L
#endif
#ifndef TVM_SETINDENT
#define TVM_SETINDENT      (WM_USER + 7)
#endif
#ifndef TVM_GETINDENT
#define TVM_GETINDENT      (WM_USER + 6)
#endif
#ifndef TVM_SETIMAGELIST
#define TVM_SETIMAGELIST   (WM_USER + 9)
#endif
#ifndef TVM_GETIMAGELIST
#define TVM_GETIMAGELIST   (WM_USER + 8)
#endif
#ifndef TVSIL_NORMAL
#define TVSIL_NORMAL       0
#endif

// TVM_SETINDENT 的合法域是 0..50 像素 (超界由控件钳掉), 缇换算后按同一上限钳,
// 免得"钳位"发生在控件里而 VB6 侧读数与显示不一致。
#define VB6_TV_INDENT_MAX_PX 50

static const wchar_t kTvIndentation[] = L"VB6_TV_Indentation";

static DWORD vb6_TvStyle(void* hwnd) {
    return (DWORD)(UINT_PTR)GetWindowLongPtrW((HWND)hwnd, GWL_STYLE);
}

// 改样式位后必须让控件重算非客户区/重绘: 复选框位改的是每个节点的度量, 只 SetWindowLong
// 的话旧度量还在用 (实测要 SWP_FRAMECHANGED 才立刻画得出方框)。
static void vb6_TvSetBit(void* hwnd, DWORD mask, int on) {
    if (!hwnd) return;
    DWORD cur = vb6_TvStyle(hwnd);
    DWORD next = on ? (cur | mask) : (cur & ~mask);
    if (next == cur) return;
    SetWindowLongPtrW((HWND)hwnd, GWL_STYLE, (LONG_PTR)next);
    SetWindowPos((HWND)hwnd, NULL, 0, 0, 0, 0,
                 SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
    InvalidateRect((HWND)hwnd, NULL, TRUE);
}

// VB6 的 True 是 -1, 不是 1 (与 vb6_GetControlVisible 同口径)。
static int32_t vb6_TvBitAsVbBool(void* hwnd, DWORD mask, int32_t defIfNoHwnd) {
    if (!hwnd) return defIfNoHwnd;
    return (vb6_TvStyle(hwnd) & mask) ? -1 : 0;
}

// ---------------- LineStyle ----------------
// 0 = tvwTreeLines: 画树线但不画根节点那一层; 1 = tvwRootLines: 根层也画线。
// 两种取值都带 TVS_HASLINES (区别只在 LINESATROOT), 所以 setter 得把 HASLINES 一并立住 ——
// 单独改 LINESATROOT 而 HASLINES 关着的话, 观感是"没有线", 与读数不符。
void vb6_TreeView_SetLineStyle(void* hwnd, int32_t val) {
    if (!hwnd) return;
    vb6_TvSetBit(hwnd, TVS_HASLINES, 1);
    vb6_TvSetBit(hwnd, TVS_LINESATROOT, val != 0);
}

int32_t vb6_TreeView_GetLineStyle(void* hwnd) {
    if (!hwnd) return 0;
    return (vb6_TvStyle(hwnd) & TVS_LINESATROOT) ? 1 : 0;
}

// ---------------- CheckBoxes ----------------
void vb6_TreeView_SetCheckBoxes(void* hwnd, int32_t val) {
    vb6_TvSetBit(hwnd, TVS_CHECKBOXES, val != 0);
}

int32_t vb6_TreeView_GetCheckBoxes(void* hwnd) {
    return vb6_TvBitAsVbBool(hwnd, TVS_CHECKBOXES, 0);
}

// ---------------- HotTracking ----------------
void vb6_TreeView_SetHotTracking(void* hwnd, int32_t val) {
    vb6_TvSetBit(hwnd, TVS_TRACKSELECT, val != 0);
}

int32_t vb6_TreeView_GetHotTracking(void* hwnd) {
    return vb6_TvBitAsVbBool(hwnd, TVS_TRACKSELECT, 0);
}

// ---------------- HideSelection ----------------
// VB6 语义与 Win32 那一位是**反**的: HideSelection=True (默认) = 失焦时不画选中项,
// 而 Win32 的 TVS_SHOWSELALWAYS = 失焦仍画选中。
void vb6_TreeView_SetHideSelection(void* hwnd, int32_t val) {
    vb6_TvSetBit(hwnd, TVS_SHOWSELALWAYS, val == 0);
}

int32_t vb6_TreeView_GetHideSelection(void* hwnd) {
    if (!hwnd) return -1;   // VB6 默认 True
    return (vb6_TvStyle(hwnd) & TVS_SHOWSELALWAYS) ? 0 : -1;
}

// ---------------- Indentation ----------------
void vb6_TreeView_SetIndentation(void* hwnd, int32_t twips) {
    if (!hwnd) return;
    if (twips < 0) twips = 0;
    SetPropW((HWND)hwnd, kTvIndentation, (HANDLE)(INT_PTR)(twips + 1));
    int px = vb6_TwipToX(twips);
    if (px > VB6_TV_INDENT_MAX_PX) px = VB6_TV_INDENT_MAX_PX;
    SendMessageW((HWND)hwnd, TVM_SETINDENT, 0, (LPARAM)px);
}

int32_t vb6_TreeView_GetIndentation(void* hwnd) {
    if (!hwnd) return 0;
    HANDLE stored = GetPropW((HWND)hwnd, kTvIndentation);
    if (stored) return (int32_t)(INT_PTR)stored - 1;
    // 没写过 ⇒ 读控件自己的当前值 (创建时 TVM_SETINDENT 尚未发过), 换算回缇。
    int px = (int)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETINDENT, 0, 0);
    int dpi = vb6_DpiX();
    return dpi > 0 ? (int32_t)MulDiv(px, 1440, dpi) : 0;
}

// ---------------- IconSize ----------------
// XtremeSuiteControls.TreeView 的扩展属性 (像素)。原生 SysTreeView32 没有"图标尺寸"这一格:
// 决定"文字从哪里起排"的是**挂没挂 ImageList**, 而不是某个尺寸属性 —— 挂了图列之后 comctl32
// 让每个条目的文字都从图宽之后起排 (即使 iImage = -1, 那一列也照留), 所以"预留图标列"的
// 唯一做法就是挂一张空图列。参考 exe 里 `tree.Icons = ...` + `tree.IconSize = 16` 留下的
// 正是这 16px 的空白列 (Outlook2013 左面板 "My Calendars" 树)。
//
// 图列按尺寸缓存: 空图列不含任何位图, 一张表可以被任意多棵树共用。
static const wchar_t kTvIconSize[] = L"VB6_TV_IconSize";

#define VB6_TV_ICON_CACHE 4
static HIMAGELIST g_tvIconList[VB6_TV_ICON_CACHE];
static int        g_tvIconPx[VB6_TV_ICON_CACHE];
static int        g_tvIconCount = 0;

static HIMAGELIST vb6_TvEmptyImageList(int px) {
    for (int i = 0; i < g_tvIconCount; i++)
        if (g_tvIconPx[i] == px) return g_tvIconList[i];
    if (g_tvIconCount >= VB6_TV_ICON_CACHE) return NULL;
    HIMAGELIST himl = ImageList_Create(px, px, ILC_COLOR32 | ILC_MASK, 1, 1);
    if (!himl) return NULL;
    g_tvIconPx[g_tvIconCount] = px;
    g_tvIconList[g_tvIconCount] = himl;
    g_tvIconCount++;
    return himl;
}

void vb6_TreeView_SetIconSize(void* hwnd, int32_t px) {
    if (!hwnd) return;
    if (px <= 0) px = 16;          // VB6 / Codejock 的默认
    if (px > 64) px = 64;
    SetPropW((HWND)hwnd, kTvIconSize, (HANDLE)(INT_PTR)(px + 1));
    // 已经挂了图列 (例如运行期先 Set Icons) 就不换 —— 那列宽度由控件自己说了算。
    HIMAGELIST himl = (HIMAGELIST)SendMessageW((HWND)hwnd, TVM_GETIMAGELIST, TVSIL_NORMAL, 0);
    if (!himl) himl = vb6_TvEmptyImageList(px);
    if (himl) {
        SendMessageW((HWND)hwnd, TVM_SETIMAGELIST, TVSIL_NORMAL, (LPARAM)himl);
        // 换图列改的是每个条目的文字起点: 与样式位那几条同样要重排重画。
        InvalidateRect((HWND)hwnd, NULL, TRUE);
    }
}

int32_t vb6_TreeView_GetIconSize(void* hwnd) {
    if (!hwnd) return 16;
    HANDLE stored = GetPropW((HWND)hwnd, kTvIconSize);
    if (stored) return (int32_t)(INT_PTR)stored - 1;
    HIMAGELIST himl = (HIMAGELIST)SendMessageW((HWND)hwnd, TVM_GETIMAGELIST, TVSIL_NORMAL, 0);
    if (himl) {
        int cx = 0, cy = 0;
        ImageList_GetIconSize(himl, &cx, &cy);
        if (cx > 0) return (int32_t)cx;
    }
    return 16;
}

// ---------------- 设计期初值 ----------------
// cgen 在建窗后按 .frm 里**写过**的属性发这一条; -999 = 未写 (不动窗口)。
// 为什么不用 ProgressBar 那套 -1 当哨兵: VB6 的布尔属性序列化成 `CheckBoxes = -1`,
// -1 恰好就是 True, 拿它当"未指定"等于设计期永远勾不上复选框。
void vb6_TreeView_Init(void* hwnd, int32_t lineStyle, int32_t indentation,
                       int32_t checkboxes, int32_t hotTracking, int32_t hideSelection) {
    if (!hwnd) return;
    if (lineStyle     != -999) vb6_TreeView_SetLineStyle(hwnd, lineStyle);
    if (indentation   != -999) vb6_TreeView_SetIndentation(hwnd, indentation);
    if (checkboxes    != -999) vb6_TreeView_SetCheckBoxes(hwnd, checkboxes);
    if (hotTracking   != -999) vb6_TreeView_SetHotTracking(hwnd, hotTracking);
    if (hideSelection != -999) vb6_TreeView_SetHideSelection(hwnd, hideSelection);
}

/* ============================================================================
 * C29-8b: Nodes / Node —— 集合面 (成员对象机制见 vb6forms_memberobj.c)
 *
 * 分工: **树的结构住在原生控件里** (父子/兄弟一律用 TVM_GETNEXTITEM 问), 本文件这张表只存
 * 原生不给的东西: Key / Text / Tag / 勾选 / 加粗 / 两个图索引。理由: SysTreeView32 存不了
 * 字符串 Key, 而自己再维护一份父子链就会与控件两真相 —— ListView 那族是"表当真相 + 整表重发",
 * 树这边不必: 五个 relationship 全部能一对一翻成 TVM_INSERTITEMW 的 (hParent, insertAfter)。
 *
 * 两个"序"要分清 (VB6 语义):
 *   - **集合序 = 插入序**: `Nodes(k)` / `Node.Index` / `For Each` 都按它。
 *     计划书 §四 原本写"For Each 顺序 = 显示序"，但 VB6 的 `Node.Index` 文档口径是
 *     "节点在集合中的位置"，让 `Nodes(k)` 与 `For Each` 用两个序会自相矛盾，所以统一按
 *     集合序 (= Index 本身)，并把这条记进手册与本文件；正常使用 (自上而下逐层 Add) 下
 *     两者本就重合。
 *   - 显示序只在导航读数里出现，且它问的是控件 (Parent/Child/Next/Previous/Root)。
 *
 * 表容量: 静态 256 条 / 每枚控件 (HeapAlloc 增长留作后续；到顶返回 0，Add 失败可观察)。
 * ========================================================================== */

#ifndef TVM_INSERTITEMW
#define TVM_INSERTITEMW   (WM_USER + 50)
#endif
#ifndef TVM_DELETEITEM
#define TVM_DELETEITEM    (WM_USER + 1)
#endif
#ifndef TVM_GETNEXTITEM
#define TVM_GETNEXTITEM   (WM_USER + 10)
#endif
#ifndef TVM_EXPAND
#define TVM_EXPAND        (WM_USER + 2)
#endif
#ifndef TVM_GETITEMSTATE
#define TVM_GETITEMSTATE  (WM_USER + 39)
#endif
#ifndef TVM_ENSUREVISIBLE
#define TVM_ENSUREVISIBLE (WM_USER + 20)
#endif
#ifndef TVM_GETITEMW
#define TVM_GETITEMW      (WM_USER + 62)
#endif
#ifndef TVM_SETITEMW
#define TVM_SETITEMW      (WM_USER + 63)
#endif
#ifndef TVIF_TEXT
#define TVIF_TEXT         0x0001
#define TVIF_IMAGE        0x0002
#define TVIF_PARAM        0x0004
#define TVIF_STATE        0x0008
#define TVIF_SELECTEDIMAGE 0x0020
#endif
#ifndef TVGN_ROOT
#define TVGN_ROOT     0
#define TVGN_NEXT     1
#define TVGN_PREVIOUS 2
#define TVGN_PARENT   3
#define TVGN_CHILD    4
#endif
#ifndef TVI_ROOT
#define TVI_ROOT    ((HTREEITEM)(INT_PTR)0xFFFF)
#define TVI_FIRST   ((HTREEITEM)(INT_PTR)0x0000)
#define TVI_LAST    ((HTREEITEM)(INT_PTR)0x0001)
#define TVI_SORT    ((HTREEITEM)(INT_PTR)0x0002)
#endif
#ifndef TVIS_STATEIMAGEMASK
#define TVIS_STATEIMAGEMASK 0xF000
#endif

// VB6 的 relationship (tvwFirst..tvwChild)，与 ai/内置控件/TreeView 控件.md 一致
#define VB6_TVG_FIRST     0
#define VB6_TVG_LAST      1
#define VB6_TVG_NEXT      2
#define VB6_TVG_PREVIOUS  3
#define VB6_TVG_CHILD     4

#define VB6_TV_NODE_MAX 256
#define VB6_TV_MAX      16      /* 一窗体里的 TreeView 枚数 */

typedef struct {
    HTREEITEM h;
    wchar_t*  key;
    wchar_t*  text;
    wchar_t*  tag;
    int       bold;
    int       image;       // 图索引 (未关联 ImageList 时只是存着, 不显形)
    int       selImage;
    int       used;
} Vb6TvNode;

typedef struct {
    HWND      hwnd;
    int       count;       // 在册条数
    Vb6TvNode n[VB6_TV_NODE_MAX];
} Vb6TvTree;

static Vb6TvTree g_tvt[VB6_TV_MAX];
static int g_tvtCount = 0;

static Vb6TvTree* vb6_TvTree(HWND h) {
    for (int i = 0; i < g_tvtCount; i++) if (g_tvt[i].hwnd == h) return &g_tvt[i];
    if (g_tvtCount >= VB6_TV_MAX) return NULL;
    Vb6TvTree* t = &g_tvt[g_tvtCount++];
    ZeroMemory(t, sizeof(*t));
    t->hwnd = h;
    return t;
}

static wchar_t* vb6_TvStr(const wchar_t* s) {
    if (!s || !*s) return NULL;
    int n = lstrlenW(s);
    wchar_t* p = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, (size_t)(n + 1) * sizeof(wchar_t));
    if (p) { CopyMemory(p, s, (size_t)n * sizeof(wchar_t)); p[n] = 0; }
    return p;
}

static void vb6_TvFree(wchar_t** pp) {
    if (*pp) { HeapFree(GetProcessHeap(), 0, *pp); *pp = NULL; }
}

static int vb6_TvIdxOfItem(Vb6TvTree* t, HTREEITEM h) {
    if (!h) return 0;
    for (int i = 0; i < VB6_TV_NODE_MAX; i++)
        if (t->n[i].used && t->n[i].h == h) return i + 1;
    return 0;
}

static HTREEITEM vb6_TvItemOfIdx(Vb6TvTree* t, int idx) {
    // 集合序 = 插入序：第 k 条在册项，不是第 k 号槽。
    if (!t || idx < 1) return NULL;
    int seen = 0;
    for (int i = 0; i < VB6_TV_NODE_MAX; i++) {
        if (!t->n[i].used) continue;
        if (++seen == idx) return t->n[i].h;
    }
    return NULL;
}

static int vb6_TvLiveCount(Vb6TvTree* t) {
    int c = 0;
    for (int i = 0; i < VB6_TV_NODE_MAX; i++) if (t->n[i].used) c++;
    return c;
}

// 原生侧的兄弟/父子关系查询 (0 = 没有)
static int vb6_TvRelIdx(Vb6TvTree* t, int idx, int flag) {
    HTREEITEM h = vb6_TvItemOfIdx(t, idx);
    if (!h) return 0;
    HTREEITEM o = (HTREEITEM)(INT_PTR)SendMessageW(t->hwnd, TVM_GETNEXTITEM, flag, (LPARAM)h);
    return vb6_TvIdxOfItem(t, o);
}

int32_t vb6_TreeView_NodeCount(void* hwnd) {
    Vb6TvTree* t = hwnd ? vb6_TvTree((HWND)hwnd) : NULL;
    return t ? (int32_t)vb6_TvLiveCount(t) : 0;
}

// Add([relative], [relationship], [key], [text]) → 1 基集合下标; 0 = 失败 (表满/参数错)
// relative 传 1 基下标 (memberobj 那侧已把"按 Key 的相对项"换成下标)，0/负 = 没有相对项。
int32_t vb6_TreeView_AddNode(void* hwnd, int32_t relative, int32_t relationship,
                             const wchar_t* key, const wchar_t* text,
                             int32_t image, int32_t selImage) {
    if (!hwnd) return 0;
    Vb6TvTree* t = vb6_TvTree((HWND)hwnd);
    if (!t || vb6_TvLiveCount(t) >= VB6_TV_NODE_MAX) return 0;

    int rel = (relative > 0) ? (int)relative : 0;
    HTREEITEM hRel = rel ? vb6_TvItemOfIdx(t, rel) : NULL;
    if (rel && !hRel) return 0;                 // 相对项不存在 → VB6 也是失败, 不猜

    HTREEITEM hParent = TVI_ROOT;
    HTREEITEM after   = TVI_LAST;
    if (!hRel) {
        // 没有相对项：一律挂在根上 (relationship 此时无意义)
        hParent = TVI_ROOT;
        after   = TVI_LAST;
    } else {
        switch (relationship) {
        case VB6_TVG_FIRST:
            hParent = (HTREEITEM)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETNEXTITEM,
                                                       TVGN_PARENT, (LPARAM)hRel);
            if (!hParent) hParent = TVI_ROOT;
            after = TVI_FIRST;
            break;
        case VB6_TVG_LAST:
            hParent = (HTREEITEM)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETNEXTITEM,
                                                       TVGN_PARENT, (LPARAM)hRel);
            if (!hParent) hParent = TVI_ROOT;
            after = TVI_LAST;
            break;
        case VB6_TVG_NEXT:    hParent = (HTREEITEM)(INT_PTR)SendMessageW(
                                   (HWND)hwnd, TVM_GETNEXTITEM, TVGN_PARENT, (LPARAM)hRel);
                              if (!hParent) hParent = TVI_ROOT;
                              after = hRel; break;
        case VB6_TVG_PREVIOUS: {
            hParent = (HTREEITEM)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETNEXTITEM,
                                                       TVGN_PARENT, (LPARAM)hRel);
            if (!hParent) hParent = TVI_ROOT;
            HTREEITEM prev = (HTREEITEM)(INT_PTR)SendMessageW(
                (HWND)hwnd, TVM_GETNEXTITEM, TVGN_PREVIOUS, (LPARAM)hRel);
            after = prev ? prev : TVI_FIRST;
            break;
        }
        case VB6_TVG_CHILD:
        default:                                 // VB6 的默认关系就是 tvwChild
            hParent = hRel;
            after   = TVI_LAST;
            break;
        }
    }

    int slot = -1;
    for (int i = 0; i < VB6_TV_NODE_MAX; i++) if (!t->n[i].used) { slot = i; break; }
    if (slot < 0) return 0;

    wchar_t txtBuf[512];
    if (text && *text) wcsncpy_s(txtBuf, 512, text, _TRUNCATE);
    else txtBuf[0] = 0;

    TVINSERTSTRUCTW ins;
    ZeroMemory(&ins, sizeof(ins));
    ins.hParent = hParent;
    ins.hInsertAfter = after;
    ins.item.mask = TVIF_TEXT | TVIF_STATE | TVIF_PARAM;
    ins.item.pszText = txtBuf;
    ins.item.cchTextMax = 0;
    // 复选框位开着才给状态图 (1 = 未勾)；否则留 0，免得往没勾选的控件上塞图号。
    ins.item.state = (vb6_TvStyle(hwnd) & TVS_CHECKBOXES) ? (DWORD)(1 << 12) : 0;
    ins.item.stateMask = TVIS_STATEIMAGEMASK;
    ins.item.lParam = 0;
    HTREEITEM h = (HTREEITEM)(INT_PTR)SendMessageW((HWND)hwnd, TVM_INSERTITEMW, 0, (LPARAM)&ins);
    if (!h) return 0;

    Vb6TvNode* nd = &t->n[slot];
    nd->h = h; nd->key = vb6_TvStr(key); nd->text = vb6_TvStr(txtBuf); nd->tag = NULL;
    nd->bold = 0; nd->image = (int)image; nd->selImage = (int)selImage; nd->used = 1;
    t->count++;
    // 集合序 = 插入序，槽位取"最小空位"⇒ 新条目前面那些槽必然全满，
    // 它的集合序就是 slot+1。别拿在册条数当返回值：删过项之后两者会分叉。
    return (int32_t)(slot + 1);
}

int32_t vb6_TreeView_NodeIndexByKey(void* hwnd, const wchar_t* key) {
    Vb6TvTree* t = hwnd ? vb6_TvTree((HWND)hwnd) : NULL;
    if (!t || !key || !*key) return 0;
    int seen = 0;
    for (int i = 0; i < VB6_TV_NODE_MAX; i++) {
        if (!t->n[i].used) continue;
        seen++;
        if (t->n[i].key && lstrcmpiW(t->n[i].key, key) == 0) return (int32_t)seen;
    }
    return 0;
}


// ---------- 条目读数 (字符串一律 BSTR，与 StatusBar/ListView 那两族同口径) ----------
static Vb6TvNode* vb6_TvNode(void* hwnd, int32_t idx) {
    Vb6TvTree* t = hwnd ? vb6_TvTree((HWND)hwnd) : NULL;
    if (!t || idx < 1) return NULL;
    int seen = 0;
    for (int i = 0; i < VB6_TV_NODE_MAX; i++) {
        if (!t->n[i].used) continue;
        if (++seen == (int)idx) return &t->n[i];
    }
    return NULL;
}

// 这三个 getter 返回**表内自有指针**，不是 BSTR：唯一的消费者是 memberobj 的 memSetStr，
// 它当场拷成自己的 BSTR；换成 SysAllocString 只会多一笔"谁负责 VariantClear"的悬账。
// 越界一律给空串而不是 NULL —— 成员对象读空字段在 VB6 里就是 ""。
const wchar_t* vb6_TreeView_GetNodeText(void* hwnd, int32_t idx) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    return (n && n->text) ? n->text : L"";
}
const wchar_t* vb6_TreeView_GetNodeKey(void* hwnd, int32_t idx) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    return (n && n->key) ? n->key : L"";
}
const wchar_t* vb6_TreeView_GetNodeTag(void* hwnd, int32_t idx) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    return (n && n->tag) ? n->tag : L"";
}

void vb6_TreeView_SetNodeText(void* hwnd, int32_t idx, const wchar_t* v) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return;
    vb6_TvFree(&n->text);
    n->text = vb6_TvStr(v);
    // 原生侧同步：不改控件的话读数是表里的新值、屏幕上还是旧字，正是本线一直防的分叉。
    TVITEMW it;
    ZeroMemory(&it, sizeof(it));
    it.mask = TVIF_HANDLE | TVIF_TEXT;
    it.hItem = n->h;
    it.pszText = (LPWSTR)(v ? v : (wchar_t*)L"");
    SendMessageW((HWND)hwnd, TVM_SETITEMW, 0, (LPARAM)&it);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}
void vb6_TreeView_SetNodeKey(void* hwnd, int32_t idx, const wchar_t* v) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return;
    vb6_TvFree(&n->key);
    n->key = vb6_TvStr(v);
}
void vb6_TreeView_SetNodeTag(void* hwnd, int32_t idx, const wchar_t* v) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return;
    vb6_TvFree(&n->tag);
    n->tag = vb6_TvStr(v);
}

// ---------- 勾选 / 展开 ----------
// 勾选住在原生 state image 里 (2 = 勾上)，所以读数与观感同源；没开 TVS_CHECKBOXES 时
// 控件不会画方框，但位照样存着 —— 与 VB6 "先设 CheckBoxes=True 才看得到" 的次序一致。
int32_t vb6_TreeView_GetNodeChecked(void* hwnd, int32_t idx) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return 0;
    DWORD st = (DWORD)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETITEMSTATE, (WPARAM)n->h,
                                            (LPARAM)TVIS_STATEIMAGEMASK);
    return (((st & TVIS_STATEIMAGEMASK) >> 12) == 2) ? -1 : 0;
}
void vb6_TreeView_SetNodeChecked(void* hwnd, int32_t idx, int32_t v) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return;
    TVITEMW it;
    ZeroMemory(&it, sizeof(it));
    // ⚠ 勾选态只能用 TVM_SETITEMW 写 —— **SDK 里根本没有 TVM_SETITEMSTATE 这条消息**
    // (commctrl.h 只有 TVM_GETITEMSTATE = TV_FIRST+39)。8b 一开始"照 GET 的样子"发明了
    // 一条 SETITEMSTATE, 编得过、跑得通、勾就是不亮 —— 未定义的 WM_USER+n 会被默认
    // 窗口过程吞掉并返回 0, 一点动静都没有。
    it.mask = TVIF_HANDLE | TVIF_STATE;
    it.hItem = n->h;
    it.stateMask = TVIS_STATEIMAGEMASK;
    it.state = (DWORD)((v != 0 ? 2 : 1) << 12);
    SendMessageW((HWND)hwnd, TVM_SETITEMW, 0, (LPARAM)&it);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

/* 账 #305 第三百四十九刀: Node.Bold —— 原生位 TVIS_BOLD, 与 Checked/Expanded
 * 同一条「读数与屏幕同源」口径 (GUISample frmClassView/frmSolutionView
 * `Root.Bold = True` 实测)。 */
int32_t vb6_TreeView_GetNodeBold(void* hwnd, int32_t idx) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return 0;
    DWORD st = (DWORD)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETITEMSTATE, (WPARAM)n->h,
                                            (LPARAM)TVIS_BOLD);
    return (st & TVIS_BOLD) ? -1 : 0;
}

void vb6_TreeView_SetNodeBold(void* hwnd, int32_t idx, int32_t v) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return;
    TVITEMW it;
    ZeroMemory(&it, sizeof(it));
    it.mask = TVIF_HANDLE | TVIF_STATE;
    it.hItem = n->h;
    it.stateMask = TVIS_BOLD;
    it.state = (v != 0) ? TVIS_BOLD : 0;
    SendMessageW((HWND)hwnd, TVM_SETITEMW, 0, (LPARAM)&it);
    InvalidateRect((HWND)hwnd, NULL, FALSE);
}

// 展开态同样住在原生位上 (TVIS_EXPANDED) —— 与 Checked 同一个理由：读数与屏幕同源。
#ifndef TVIS_EXPANDED
#define TVIS_EXPANDED 0x0020
#endif
#ifndef TVE_EXPAND
#define TVE_EXPAND   0x0002
#define TVE_COLLAPSE 0x0001
#endif

int32_t vb6_TreeView_GetNodeExpanded(void* hwnd, int32_t idx) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return 0;
    DWORD st = (DWORD)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETITEMSTATE, (WPARAM)n->h,
                                            (LPARAM)TVIS_EXPANDED);
    return (st & TVIS_EXPANDED) ? -1 : 0;
}

void vb6_TreeView_SetNodeExpanded(void* hwnd, int32_t idx, int32_t v) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return;
    SendMessageW((HWND)hwnd, TVM_EXPAND, (WPARAM)(v != 0 ? TVE_EXPAND : TVE_COLLAPSE),
                 (LPARAM)n->h);
}

// ---------- 导航 (结构住在控件里，一律现问) ----------
int32_t vb6_TreeView_GetNodeParent(void* hwnd, int32_t idx) { return vb6_TvRelIdx(vb6_TvTree((HWND)hwnd), idx, TVGN_PARENT); }
int32_t vb6_TreeView_GetNodeChild(void* hwnd, int32_t idx)  { return vb6_TvRelIdx(vb6_TvTree((HWND)hwnd), idx, TVGN_CHILD); }
int32_t vb6_TreeView_GetNodeNext(void* hwnd, int32_t idx)   { return vb6_TvRelIdx(vb6_TvTree((HWND)hwnd), idx, TVGN_NEXT); }
int32_t vb6_TreeView_GetNodePrev(void* hwnd, int32_t idx)   { return vb6_TvRelIdx(vb6_TvTree((HWND)hwnd), idx, TVGN_PREVIOUS); }

int32_t vb6_TreeView_GetNodeChildren(void* hwnd, int32_t idx) {
    Vb6TvTree* t = hwnd ? vb6_TvTree((HWND)hwnd) : NULL;
    if (!t) return 0;
    int32_t c = 0;
    for (int32_t k = vb6_TreeView_GetNodeChild(hwnd, idx); k; k = vb6_TreeView_GetNodeNext(hwnd, k)) c++;
    return c;
}

// VB6 的 Root = **所属树的根祖先**：顶层节点问自己就返回自己 (不是"没有")。
int32_t vb6_TreeView_GetNodeRoot(void* hwnd, int32_t idx) {
    int32_t cur = idx;
    if (!vb6_TvNode(hwnd, idx)) return 0;   // 越界不给"自己的下标"当根
    for (;;) {
        int32_t up = vb6_TreeView_GetNodeParent(hwnd, cur);
        if (!up) return cur;
        cur = up;
    }
}

void vb6_TreeView_NodeEnsureVisible(void* hwnd, int32_t idx) {
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return;
    SendMessageW((HWND)hwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)n->h);
}

// ---------- Remove / Clear ----------
// 原生一次删掉整棵子树，所以表里要把"它的后代"一起标成不用 —— 漏了这一步，
// 集合数与 Count 就会与屏幕上看到的不一致。
int32_t vb6_TreeView_RemoveNode(void* hwnd, int32_t idx) {
    Vb6TvTree* t = hwnd ? vb6_TvTree((HWND)hwnd) : NULL;
    if (!t) return 0;
    Vb6TvNode* n = vb6_TvNode(hwnd, idx);
    if (!n) return 0;

    // 后代：谁的祖先链上有 n->h 就一起完蛋。判据现问原生树，不另存一份父子。
    int doomed[VB6_TV_NODE_MAX]; int dn = 0;
    for (int i = 0; i < VB6_TV_NODE_MAX; i++) {
        if (!t->n[i].used) continue;
        HTREEITEM a = (HTREEITEM)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETNEXTITEM,
                                                       TVGN_PARENT, (LPARAM)t->n[i].h);
        while (a) {
            if (a == n->h) { doomed[dn++] = i; break; }
            a = (HTREEITEM)(INT_PTR)SendMessageW((HWND)hwnd, TVM_GETNEXTITEM,
                                                 TVGN_PARENT, (LPARAM)a);
        }
    }
    // 只发这一条删除：原生自己带走整棵子树。挨个删 doomed 会往已失效的 HTREEITEM
    // 上发消息 (槽位回收后父项下标可能比子项大，"按表序反着删"并不成立)。
    SendMessageW((HWND)hwnd, TVM_DELETEITEM, 0, (LPARAM)n->h);
    doomed[dn++] = (int)(n - t->n);
    for (int k = 0; k < dn; k++) {
        int i = doomed[k];
        if (!t->n[i].used) continue;
        vb6_TvFree(&t->n[i].key); vb6_TvFree(&t->n[i].text); vb6_TvFree(&t->n[i].tag);
        ZeroMemory(&t->n[i], sizeof(t->n[i]));
        t->count--;
    }
    return -1;
}

void vb6_TreeView_ClearNodes(void* hwnd) {
    Vb6TvTree* t = hwnd ? vb6_TvTree((HWND)hwnd) : NULL;
    if (!t) return;
    SendMessageW((HWND)hwnd, TVM_DELETEITEM, 0, (LPARAM)TVI_ROOT);
    for (int i = 0; i < VB6_TV_NODE_MAX; i++) {
        if (!t->n[i].used) continue;
        vb6_TvFree(&t->n[i].key); vb6_TvFree(&t->n[i].text); vb6_TvFree(&t->n[i].tag);
        ZeroMemory(&t->n[i], sizeof(t->n[i]));
    }
    t->count = 0;
}
/* ============================================================================
 * C29-8c: TreeView 的通知派发 (NodeClick / Expand / Collapse)
 *   通道与 C29-7 ListView、C29-4 StatusBar 同一条 (父窗的 WM_NOTIFY)。
 *   码值与本机的 10.0.19041 SDK commctrl.h 对过：TVN_FIRST = (0U-400U)，
 *   TVN_SELCHANGEDW = TVN_FIRST-51 = -451、TVN_ITEMEXPANDEDW = TVN_FIRST-55 = -455
 *   (那条 -400 系的是 **W** 版；老的 A 版号段不同，本仓库产物一律 Unicode)。
 *   通知负载是 NMTREEVIEWW：`itemNew.hItem` 就是"这次涉及的那个节点"，
 *   所以换算只问那张表 (句柄→1 基序号)，不另存一份节点身份。
 * ==========================================================================*/

#ifndef TVN_SELCHANGEDW
#define TVN_SELCHANGEDW   (-451)
#endif
#ifndef TVN_ITEMEXPANDEDW
#define TVN_ITEMEXPANDEDW (-455)
#endif

// vb6_TvTree 会**顺手开一格** (那是给控件首次使用者设计的)；通知换算只做查询，
// 绝不能因为收到一条别家的 WM_NOTIFY 就占掉一枚槽位。
static Vb6TvTree* vb6_TvFind(HWND h) {
    for (int i = 0; i < g_tvtCount; i++) if (g_tvt[i].hwnd == h) return &g_tvt[i];
    return NULL;
}
// 通知折成 1 基节点序号 (0 = 这个句柄不在册 / 不是本控件的节点)。
int32_t vb6_TreeView_NotifyNodeIndex(void* hwnd, void* nmTreeViewW) {
    if (!hwnd || !nmTreeViewW) return 0;
    NMTREEVIEWW* nv = (NMTREEVIEWW*)nmTreeViewW;
    Vb6TvTree* t = vb6_TvFind((HWND)hwnd);
    if (!t) return 0;
    return (int32_t)vb6_TvIdxOfItem(t, nv->itemNew.hItem);
}

// 展开还是折回。**只认这两个明确的 action 值**: 认不出的一律给 999 —— 派发那侧
// 两个分支分别比 == -1 / == 0，于是"认不出"两边都不接。
// (原来写成"带 TVE_COLLAPSE 就是折回，否则算展开"，结果 TVN_SELCHANGED 那条通知
//  的 action 字段没填 TVE_COLLAPSE ⇒ NodeClick 顺手把 Expand 也点了一次，实测读数
//  gExpands 白涨两条。布尔式地猜一个三态字段就是这种下场。)
int32_t vb6_TreeView_NotifyExpanded(void* nmTreeViewW) {
    if (!nmTreeViewW) return 999;
    NMTREEVIEWW* nv = (NMTREEVIEWW*)nmTreeViewW;
    if (nv->action == TVE_EXPAND) return -1;
    if (nv->action == TVE_COLLAPSE) return 0;
    return 999;
}

// ---------------- 判据辅助 (写进 ai/029 边界，不对应任何 VB6 语义) ----------------
// 无头环境点不了鼠标，而直接调 handler 会绕开整条派发链 —— 只有从真窗口消息进来，
// 才验得到"派发分支 + 通知换算 + NodeAt 造对象"三段都接上了。手法照 C29-4 的 SimClick。
// action 由调用方**原样**给 (NodeClick 那条填 0 ⇒ NotifyExpanded 认不出 = 999，
// 展开/折回两个分支都不接 —— 上一条注释里那个"布尔式猜三态"的坑不再复发)。
static void vb6_TvSendNotify(HWND hwnd, int32_t idx, int32_t code, UINT action) {
    NMTREEVIEWW nv;
    Vb6TvTree* t = hwnd ? vb6_TvFind(hwnd) : NULL;
    HTREEITEM h = t ? vb6_TvItemOfIdx(t, idx) : NULL;
    if (!h) return;
    ZeroMemory(&nv, sizeof(nv));
    nv.hdr.hwndFrom = hwnd;
    nv.hdr.idFrom   = (UINT_PTR)idx;
    nv.hdr.code     = (DWORD)code;
    nv.action       = action;
    nv.itemNew.mask = TVIF_HANDLE;
    nv.itemNew.hItem = h;
    HWND parent = GetParent(hwnd);
    if (!parent) parent = hwnd;
    SendMessageW(parent, WM_NOTIFY, 0, (LPARAM)&nv);
}

void vb6_TreeView_SimNodeClick(void* hwnd, int32_t nodeIdx) {
    if (nodeIdx < 1) return;
    vb6_TvSendNotify((HWND)hwnd, nodeIdx, TVN_SELCHANGEDW, 0);   /* 通知里没有 action 这一说 */
}

void vb6_TreeView_SimExpand(void* hwnd, int32_t nodeIdx, int32_t expanded) {
    if (nodeIdx < 1) return;
    vb6_TvSendNotify((HWND)hwnd, nodeIdx, TVN_ITEMEXPANDEDW,
                     (expanded ? (UINT)TVE_EXPAND : (UINT)TVE_COLLAPSE));
}


#endif // _WIN32
