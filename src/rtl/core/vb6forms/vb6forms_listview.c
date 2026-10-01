// vb6forms_listview.c - P20-45: VB6 ListView 控件的原生复刻
//
// 路线: **不加载 MSCOMCTL.OCX**, 用 comctl32 的 SysListView32 复刻等价功能与事件。
//   SysListView32 是 comctl32 注册好的窗口类 (x64/x86 实测 GetClassInfoW 直接成功),
//   所以**不需要** StatusBar 那套 RTL 自注册兜底。
//
// 设计要点: **自己那张表才是唯一真相**, SysListView32 只是显示。
//   理由: LVM_GETITEMTEXT 在 x64 跨进程/不同 comctl 版本下取文本有坑, 而设计期灌进去的
//   文本按 VB6 语义要能**原样读回**(Debug.Print ListView1.ListItems(1).Text)。所以
//   写入时同时改表 + 刷 UI; 读取一律读表。改动后 lvSyncAll() 按 syncedCols/syncedRows
//   做**增量**同步 (Fix 161f-extlist(3): 原先每次全量 LVM_DELETEALLITEMS 重建,
//   RefillList(500 行 x 10 列) 逐格灌数据触发 O(n^2) 千万次 SendMessage, 消息循环被拖死,
//   表现为"exe 打不开"), 再按表的 selected/checked 恢复状态。
//
// 下标口径 (VB6):
//   - ColumnHeaders(i) / ListItems(i)  **1 基**
//   - ListItem.SubItems(i)             **1 基, 且 i=1 对应第 2 列** (第 1 列是 Text)
//   - ColumnHeaders.Add 的 index 是**插入位** (省略/0 = 追加)
//
// 依赖: 只 include 同族头 (生成代码包含的那一族) + 平铺的 vb6rtl_bstr.h。

#include <windows.h>
#include <commctrl.h>
#include <stdbool.h>
#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6rtl_bstr.h"

#ifndef _WIN32
#error "vb6forms_listview.c 只支持 Win32"
#endif

#define VB6_LV_MAX 32

// ---------------- 列头 ----------------
typedef struct {
    wchar_t* key;
    wchar_t* text;
    int      width;
    int      align;      // 0=左 1=右 2=居中 (lvwColumnLeft/Right/Center)
    // Fix 161f-extlist(4): 这列的宽度是否来自"省略 width 的 Add"。
    // 留着它是因为 lvSyncAll 被 AddColumn 调多次 (每加一列一次), 需要知道
    // 哪些列该套默认宽 —— 显式给宽的列 (含 Adjust Widths 改过的) 不能被动。
    bool     defaultWidth;
} Vb6LVCol;

// Fix 161f-extlist(4): VB6 ListView ColumnHeader 省略 Width 时的**默认列宽**。
// 定标依据: 基准截图 (VB6 原版运行, 1536px 物理宽 = 0.8×1920 屏) 里表头分隔线
// 依次为 151,295,439,...,1445 —— 相邻差恒为 144px, 十列等宽。参照同一图里
// Option1 圆点在物理 x=365 (= 4 + 292 逻辑 × 1.25) 标定 **参考机是 125% DPI**,
// 故逻辑列宽 = 144 / 1.25 = 115.2 ⇒ 115px。
// (此前按"2x DPI"误标成 72px —— 同一张图里单选钮位置与 125% DPI 严丝合缝,
//  2x 假设下 Option 会落在 442px, 与实测矛盾。)
// 即: VB6 这里既不是 0、也不是"按表头文本自适应"(那是 LVSCW_AUTOSIZE_USEHEADER
// 的语义, 上一版照它实现, col0 被"单列铺满控件"规则污染成 389px), 而是一个
// **固定默认值** ≈ 1728 缇。
#define VB6_LV_DEFAULT_COL_W 115

// ---------------- 行 ----------------
typedef struct {
    wchar_t*  key;
    wchar_t*  text;      // 第 1 列
    wchar_t*  tag;
    wchar_t** subs;      // 第 2..n 列 (subs[0] = 第 2 列)
    int       subCount;
    int       subCap;
    int       icon;
    int       smallIcon;
    int       selected;
    int       checked;
} Vb6LVItem;

typedef struct {
    HWND hwnd;
    int  view;              // 0=lvwIcon 1=lvwSmallIcon 2=lvwList 3=lvwReport
    int  gridLines;
    int  fullRowSelect;
    int  multiSelect;
    int  checkBoxes;
    int  hideHeaders;
    int  allowColReorder;
    int  labelEdit;
    int  sorted;
    int  sortKey;
    int  sortOrder;
    Vb6LVCol*  cols;  int colCount,  colCap;
    Vb6LVItem* items; int itemCount, itemCap;
    void* cb[2];            // 0=ItemClick 1=ColumnClick (WM_NOTIFY 分发用)
    // Fix 161f-extlist(3): 已同步到 UI 的列数/行数 —— lvSyncAll 由"每次全删重建"
    // 改为**增量同步**。原实现下 RefillList(500 行 × 10 列) 每格 SetItemSub 都触发
    // 一次全量 DELETEALLITEMS + 重插全部行 ⇒ O(n²) ≈ 千万次 LVM_* ⇒ 窗口起来后
    // 消息循环被拖死 (表现为"打不开/无响应")。
    int syncedCols;
    int syncedRows;
} Vb6ListView;

static Vb6ListView g_lvs[VB6_LV_MAX];
static int g_lvCount = 0;

static Vb6ListView* lvFind(HWND h) {
    for (int i = 0; i < g_lvCount; i++) if (g_lvs[i].hwnd == h) return &g_lvs[i];
    return NULL;
}

static Vb6ListView* lvEnsure(HWND h) {
    Vb6ListView* v = lvFind(h);
    if (v) return v;
    if (g_lvCount >= VB6_LV_MAX) return NULL;
    v = &g_lvs[g_lvCount++];
    ZeroMemory(v, sizeof(*v));
    v->hwnd = h;
    v->view = 3;            // VB6 默认 lvwReport? 实际默认 0(lvwIcon); C3 取 0 与 VB6 一致
    v->view = 0;
    v->sortOrder = 0;
    return v;
}

// ---- 小工具: 宽串复制到堆 (自己管) ----
static wchar_t* lvDupW(const wchar_t* s) {
    if (!s) return NULL;
    int n = lstrlenW(s);
    wchar_t* p = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, (n + 1) * sizeof(wchar_t));
    if (!p) return NULL;
    memcpy(p, s, n * sizeof(wchar_t));
    p[n] = 0;
    return p;
}
static void lvFreeW(wchar_t* p) { if (p) HeapFree(GetProcessHeap(), 0, p); }
static void lvSetW(wchar_t** dst, const wchar_t* src) {
    if (*dst) { lvFreeW(*dst); *dst = NULL; }
    if (src) *dst = lvDupW(src);
}

// ---------------- UI 同步 ----------------
// 列: LVM_INSERTCOLUMNW; 行: LVM_INSERTITEMW + LVM_SETITEMW(iSubItem=1..)
//
// Fix 161f-extlist(3): **增量同步**。原实现每次调用都 LVM_DELETEALLITEMS + 重建全部
// 列与行 —— 在"逐格灌数据"的常规写法下 (RefillList: 500 行 × 每行 1 次 Add +
// 9 次 SubItems(i)=...) 总代价是 O(行数²) (= ~1250 万次 SendMessage), 窗口消息循环
// 被拖到几十秒不响应 (实测表现为"打不开")。改为:
//   ① 列/行数未减少 → 只补插**新增**的列与行 (已插过的不动);
//   ② 行数/列数**减少** (Clear / 重建) → 才走 DELETEALLITEMS 全量重建。
// 文本单格更新 (SetItemSub / SetItemText) 另走 lvSyncCell, 不碰整表。
static void lvSyncFull(Vb6ListView* v) {
    SendMessageW(v->hwnd, LVM_DELETEALLITEMS, 0, 0);
    // ⚠ 这里必须**同时清列**: lvSyncAll 之后会从 syncedCols(=0) 重插到 colCount,
    // 若控件里旧列还在, LVM_INSERTCOLUMN 是**追加**语义 → 列翻倍。
    // ClearColumns/ClearItems 各自会把 syncedCols/syncedRows 归零, 但"数量变少"
    // 这条路径 (直接 RemoveColumn / 删行) 只会走到这里, 不补会重复。
    HWND hdr = (HWND)SendMessageW(v->hwnd, LVM_GETHEADER, 0, 0);
    int curCols = hdr ? (int)SendMessageW(hdr, HDM_GETITEMCOUNT, 0, 0) : 0;
    for (int i = curCols - 1; i >= 0; i--)
        SendMessageW(v->hwnd, LVM_DELETECOLUMN, (WPARAM)i, 0);
    v->syncedCols = 0;
    v->syncedRows = 0;
}
static void lvInsertColUI(Vb6ListView* v, int i) {
    LVCOLUMNW col;
    ZeroMemory(&col, sizeof(col));
    col.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM | LVCF_FMT;
    col.pszText = v->cols[i].text ? v->cols[i].text : (wchar_t*)L"";
    /* Fix 161f-extlist(4): 省略 width 的列在 AddColumn 里已填好固定默认宽
     * (VB6_LV_DEFAULT_COL_W), 这里直接下发即可。
     * ⚠ 不能改成下发 LVSCW_AUTOSIZE_USEHEADER: LVM_INSERTCOLUMNW 不认该常量
     * (传 -2 被当负宽 → 0), 且它的语义 (按表头文本算 / 单列时铺满控件) 与
     * VB6 实际表现的"固定 72px 等宽"不符 (两者都实测过)。 */
    col.cx = v->cols[i].width;
    col.iSubItem = i;
    col.fmt = (v->cols[i].align == 1) ? LVCFMT_RIGHT
            : (v->cols[i].align == 2) ? LVCFMT_CENTER : LVCFMT_LEFT;
    SendMessageW(v->hwnd, LVM_INSERTCOLUMNW, (WPARAM)i, (LPARAM)&col);
}
static void lvInsertRowUI(Vb6ListView* v, int r) {
    Vb6LVItem* it = &v->items[r];
    LVITEMW li;
    ZeroMemory(&li, sizeof(li));
    li.mask = LVIF_TEXT;
    li.iItem = r;
    li.iSubItem = 0;
    li.pszText = it->text ? it->text : (wchar_t*)L"";
    SendMessageW(v->hwnd, LVM_INSERTITEMW, 0, (LPARAM)&li);
    for (int s = 0; s < it->subCount; s++) {
        LVITEMW si;
        ZeroMemory(&si, sizeof(si));
        si.mask = LVIF_TEXT;
        si.iItem = r;
        si.iSubItem = s + 1;                 // SubItems(1) = 第 2 列
        si.pszText = it->subs[s] ? it->subs[s] : (wchar_t*)L"";
        SendMessageW(v->hwnd, LVM_SETITEMW, 0, (LPARAM)&si);
    }
    if (it->selected || it->checked) {
        SendMessageW(v->hwnd, LVM_SETITEMSTATE, (WPARAM)r, (LPARAM)&(LVITEMW){
            .mask = LVIF_STATE,
            .state = (UINT)((it->selected ? LVIS_SELECTED : 0) | (it->checked ? LVIS_STATEIMAGEMASK : 0)),
            .stateMask = LVIS_SELECTED | LVIS_STATEIMAGEMASK });
    }
}
static void lvSyncAll(Vb6ListView* v) {
    if (!v) return;
    if (!v || !v->hwnd || !IsWindow(v->hwnd)) return;
    // 减少 (Clear / 大幅重建) → 全量重建
    if (v->colCount < v->syncedCols || v->itemCount < v->syncedRows) {
        lvSyncFull(v);
    }
    // 补插新列
    for (int i = v->syncedCols; i < v->colCount; i++) lvInsertColUI(v, i);
    v->syncedCols = v->colCount;
    // 补插新行
    for (int r = v->syncedRows; r < v->itemCount; r++) lvInsertRowUI(v, r);
    v->syncedRows = v->itemCount;
}
// 单格文本更新 (不改行列数) — 避免 SetItemSub 走整表同步。
static void lvSyncCell(Vb6ListView* v, int r, int sub) {
    if (!v || !v->hwnd || !IsWindow(v->hwnd)) return;
    if (r < 0 || r >= v->itemCount || sub < 0 || sub > v->syncedCols) return;
    Vb6LVItem* it = &v->items[r];
    LVITEMW si;
    ZeroMemory(&si, sizeof(si));
    si.mask = LVIF_TEXT;
    si.iItem = r;
    si.iSubItem = sub;                       // sub=0 → 第 1 列 (Text)
    si.pszText = (sub == 0)
        ? (it->text ? it->text : (wchar_t*)L"")
        : ((sub - 1 < it->subCount && it->subs[sub - 1]) ? it->subs[sub - 1] : (wchar_t*)L"");
    if (sub == 0) SendMessageW(v->hwnd, LVM_SETITEMTEXTW, (WPARAM)r, (LPARAM)&si);
    else          SendMessageW(v->hwnd, LVM_SETITEMW, 0, (LPARAM)&si);
}

// 把标量属性刷到窗口 (GWL_STYLE 位 + LVM_SETEXTENDEDLISTVIEWSTYLE)
static void lvApplyStyle(Vb6ListView* v) {
    if (!v || !v->hwnd) return;
    LONG_PTR st = GetWindowLongPtrW(v->hwnd, GWL_STYLE);
    st &= ~(LONG_PTR)(LVS_TYPEMASK | LVS_SINGLESEL | LVS_NOCOLUMNHEADER);
    switch (v->view) {
        case 0:  st |= LVS_ICON;       break;
        case 1:  st |= LVS_SMALLICON;  break;
        case 2:  st |= LVS_LIST;       break;
        default: st |= LVS_REPORT;     break;
    }
    if (!v->multiSelect) st |= LVS_SINGLESEL;
    if (v->hideHeaders)  st |= LVS_NOCOLUMNHEADER;
    SetWindowLongPtrW(v->hwnd, GWL_STYLE, st);
    DWORD ex = (DWORD)SendMessageW(v->hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0);
    ex &= ~(LVS_EX_GRIDLINES | LVS_EX_FULLROWSELECT | LVS_EX_CHECKBOXES | LVS_EX_HEADERDRAGDROP);
    if (v->gridLines)      ex |= LVS_EX_GRIDLINES;
    if (v->fullRowSelect)  ex |= LVS_EX_FULLROWSELECT;
    if (v->checkBoxes)     ex |= LVS_EX_CHECKBOXES;
    if (v->allowColReorder) ex |= LVS_EX_HEADERDRAGDROP;
    SendMessageW(v->hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, (LPARAM)ex);
    SendMessageW(v->hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, ex, (LPARAM)ex);  // 生效 (双调用模式)
}

// ---------------- 入口: 设计期一次性初始化 ----------------
// 参数顺序 = .frm 里各个设计期属性的常见集合; 生成器按名字取好再传进来。
void vb6_ListView_Init(void* hwnd, int32_t view, int32_t gridLines, int32_t fullRowSelect,
                       int32_t multiSelect, int32_t checkBoxes, int32_t hideHeaders,
                       int32_t allowColReorder, int32_t labelEdit) {
    if (!hwnd) return;
    Vb6ListView* v = lvEnsure((HWND)hwnd);
    if (!v) return;
    v->view = (int)view;
    // 账 #128-b: 这六位是布尔 ⇒ **存进去就先归化成 VB6 的 -1/0**。以前是"写什么存什么"，
    // 于是设计期发的 1 会原样读回 1 —— 类型面一旦说自己是 Boolean，`X = True` 这条比较
    // 就恒假（1 ≠ -1）。归一化放在存这一头，getter 仍是直答控件里的那个数，不多折一层。
    v->gridLines = gridLines ? -1 : 0;
    v->fullRowSelect = fullRowSelect ? -1 : 0;
    v->multiSelect = multiSelect ? -1 : 0;
    v->checkBoxes = checkBoxes ? -1 : 0;
    v->hideHeaders = hideHeaders ? -1 : 0;
    v->allowColReorder = allowColReorder ? -1 : 0;
    v->labelEdit = (int)labelEdit;
    lvApplyStyle(v);
}

// ---------------- 标量属性 读 ----------------
int32_t vb6_ListView_GetView(void* hwnd)            { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->view : 0; }
int32_t vb6_ListView_GetGridLines(void* hwnd)       { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->gridLines : 0; }
int32_t vb6_ListView_GetFullRowSelect(void* hwnd)   { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->fullRowSelect : 0; }
int32_t vb6_ListView_GetMultiSelect(void* hwnd)     { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->multiSelect : 0; }
int32_t vb6_ListView_GetCheckBoxes(void* hwnd)      { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->checkBoxes : 0; }
int32_t vb6_ListView_GetHideColumnHeaders(void* hwnd) { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->hideHeaders : 0; }
int32_t vb6_ListView_GetAllowColumnReorder(void* hwnd) { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->allowColReorder : 0; }
int32_t vb6_ListView_GetLabelEdit(void* hwnd)       { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->labelEdit : 0; }
int32_t vb6_ListView_GetSorted(void* hwnd)          { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->sorted : 0; }
int32_t vb6_ListView_GetSortKey(void* hwnd)         { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->sortKey : 0; }
int32_t vb6_ListView_GetSortOrder(void* hwnd)       { Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->sortOrder : 0; }

// ---------------- 标量属性 写 ----------------
void vb6_ListView_SetView(void* hwnd, int32_t val) {
    Vb6ListView* v = lvFind((HWND)hwnd); if (!v) return;
    v->view = (int)val; lvApplyStyle(v);
}
void vb6_ListView_SetGridLines(void* hwnd, int32_t val)         { Vb6ListView* v = lvFind((HWND)hwnd); if (v) { v->gridLines = val ? -1 : 0; lvApplyStyle(v); } }
void vb6_ListView_SetFullRowSelect(void* hwnd, int32_t val)     { Vb6ListView* v = lvFind((HWND)hwnd); if (v) { v->fullRowSelect = val ? -1 : 0; lvApplyStyle(v); } }
void vb6_ListView_SetMultiSelect(void* hwnd, int32_t val)       { Vb6ListView* v = lvFind((HWND)hwnd); if (v) { v->multiSelect = val ? -1 : 0; lvApplyStyle(v); } }
void vb6_ListView_SetCheckBoxes(void* hwnd, int32_t val)        { Vb6ListView* v = lvFind((HWND)hwnd); if (v) { v->checkBoxes = val ? -1 : 0; lvApplyStyle(v); lvSyncAll(v); } }
void vb6_ListView_SetHideColumnHeaders(void* hwnd, int32_t val) { Vb6ListView* v = lvFind((HWND)hwnd); if (v) { v->hideHeaders = val ? -1 : 0; lvApplyStyle(v); } }
void vb6_ListView_SetAllowColumnReorder(void* hwnd, int32_t val) { Vb6ListView* v = lvFind((HWND)hwnd); if (v) { v->allowColReorder = val ? -1 : 0; lvApplyStyle(v); } }
void vb6_ListView_SetLabelEdit(void* hwnd, int32_t val)         { Vb6ListView* v = lvFind((HWND)hwnd); if (v) v->labelEdit = (int)val; }
void vb6_ListView_SetSorted(void* hwnd, int32_t val) {
    Vb6ListView* v = lvFind((HWND)hwnd); if (!v) return;
    v->sorted = (int)val;
    // VB6: Sorted=True 立即按 SortKey/SortOrder 排序。用表排, 再重建 UI, 保持一致。
    if (v->sorted && v->itemCount > 1) {
        for (int i = 1; i < v->itemCount; i++) {
            Vb6LVItem tmp = v->items[i];
            int j = i - 1;
            while (j >= 0) {
                const wchar_t* a = NULL; const wchar_t* b = NULL;
                if (v->sortKey <= 0) { a = v->items[j].text; b = tmp.text; }
                else {
                    int si = v->sortKey - 1;   // SortKey 1 => SubItems(1)
                    a = (si < v->items[j].subCount) ? v->items[j].subs[si] : L"";
                    b = (si < tmp.subCount)         ? tmp.subs[si]         : L"";
                }
                int cmp = lstrcmpW(a ? a : L"", b ? b : L"");
                if (v->sortOrder == 1) cmp = -cmp;
                if (cmp <= 0) break;
                v->items[j + 1] = v->items[j];
                j--;
            }
            v->items[j + 1] = tmp;
        }
    }
    lvSyncAll(v);
}
void vb6_ListView_SetSortKey(void* hwnd, int32_t val)   { Vb6ListView* v = lvFind((HWND)hwnd); if (v) v->sortKey = (int)val; }
void vb6_ListView_SetSortOrder(void* hwnd, int32_t val) { Vb6ListView* v = lvFind((HWND)hwnd); if (v) v->sortOrder = (int)val; }

// ---------------- 列头集合 ----------------
int32_t vb6_ListView_GetColumnCount(void* hwnd) {
    Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->colCount : 0;
}

// ColumnHeaders.Add [index], [key], [text], [width], [alignment]
// idx 是**插入位** (1 基); <=0 或 > count 时追加 (VB6 语义)。
int32_t vb6_ListView_AddColumn(void* hwnd, int32_t idx, void* keyBstr, void* textBstr,
                               int32_t width, int32_t align) {
    Vb6ListView* v = lvEnsure((HWND)hwnd);
    if (!v) return 0;
    if (v->colCount + 1 > v->colCap) {
        int nc = v->colCap ? v->colCap * 2 : 8;
        Vb6LVCol* na;
        if (v->cols) na = (Vb6LVCol*)HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                                 v->cols, sizeof(Vb6LVCol) * nc);
        else         na = (Vb6LVCol*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                                sizeof(Vb6LVCol) * nc);
        if (!na) return 0;
        v->cols = na; v->colCap = nc;
    }
    int at = (idx >= 1 && idx <= v->colCount) ? (int)idx - 1 : v->colCount;
    if (at < v->colCount)
        memmove(&v->cols[at + 1], &v->cols[at], sizeof(Vb6LVCol) * (v->colCount - at));
    ZeroMemory(&v->cols[at], sizeof(Vb6LVCol));
    lvSetW(&v->cols[at].key, (const wchar_t*)keyBstr);
    lvSetW(&v->cols[at].text, (const wchar_t*)textBstr);
    // Fix 161f-extlist(4): `ColumnHeaders.Add(, , "text")` 省略 width (Missing) 时,
    // VB6 用**固定默认列宽 115px** (见 VB6_LV_DEFAULT_COL_W 的定标说明),
    // 既不是原先折成的 0、也不是"表头文本自适应"。
    // (0 ⇒ 10 列全 0 像素宽、500 行不可见; 表头自适应 ⇒ 各列宽参差且第 0 列被
    //  "单列铺满控件"规则污染成 389px —— 两种都试过且都实测不符。)
    int w0 = (int)width;
    v->cols[at].defaultWidth = (w0 <= 0);
    v->cols[at].width = (w0 > 0) ? w0 : VB6_LV_DEFAULT_COL_W;
    v->cols[at].align = (int)align;
    v->colCount++;
    lvSyncAll(v);
    return at + 1;   // 返回 1 基 Index
}

// 列头文本/宽度 (1 基)
void* vb6_ListView_GetColumnText(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount || !v->cols[idx - 1].text) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(v->cols[idx - 1].text);
}
void vb6_ListView_SetColumnText(void* hwnd, int32_t idx, void* bstr) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount) return;
    lvSetW(&v->cols[idx - 1].text, (const wchar_t*)bstr);
    lvSyncAll(v);
}
int32_t vb6_ListView_GetColumnWidth(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount) return 0;
    return v->cols[idx - 1].width;
}
void vb6_ListView_SetColumnWidth(void* hwnd, int32_t idx, int32_t w) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount) return;
    // Fix 161f-extlist(4): 列已建好时改宽必须发 LVM_SETCOLUMNWIDTH ——
    // lvSyncAll 只在**新增**列时 LVM_INSERTCOLUMN, 已同步的列不会被重发,
    // 只改结构体等于改了个寂寞 (Flip Columns / Adjust Widths 按钮点击无效)。
    int i0 = (int)idx - 1;
    // 显式给宽 = 放弃"默认宽"身份 (Adjust Widths 传 LVSCW_AUTOSIZE_USEHEADER
    // 进来的除外 —— 那是显式要求按表头自适应, 由 comctl32 算)。
    bool wantAuto = ((int)w == LVSCW_AUTOSIZE_USEHEADER || (int)w == LVSCW_AUTOSIZE);
    v->cols[i0].defaultWidth = false;
    v->cols[i0].width = wantAuto ? VB6_LV_DEFAULT_COL_W : (int)w;
    if (!v->hwnd || !IsWindow(v->hwnd)) return;
    if (i0 < v->syncedCols) {
        SendMessageW(v->hwnd, LVM_SETCOLUMNWIDTH, (WPARAM)i0, (LPARAM)(int)w);
        if (wantAuto)
            v->cols[i0].width = (int)SendMessageW(v->hwnd, LVM_GETCOLUMNWIDTH, (WPARAM)i0, 0);
    } else {
        lvSyncAll(v);
    }
}
// 列对齐 (0=左 1=右 2=居中 —— VB6 的 lvwColumnLeft/Right/Center)
void vb6_ListView_SetColumnAlign(void* hwnd, int32_t idx, int32_t val) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount) return;
    v->cols[idx - 1].align = (int)val;
    lvSyncAll(v);
}

int32_t vb6_ListView_GetColumnAlign(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount) return 0;
    return v->cols[idx - 1].align;
}
// 列头 Key / Index (反查)
void* vb6_ListView_GetColumnKey(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount || !v->cols[idx - 1].key) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(v->cols[idx - 1].key);
}
int32_t vb6_ListView_GetColumnIndexByKey(void* hwnd, void* keyBstr) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || !keyBstr) return 0;
    const wchar_t* k = (const wchar_t*)keyBstr;
    for (int i = 0; i < v->colCount; i++)
        if (v->cols[i].key && !lstrcmpW(v->cols[i].key, k)) return i + 1;
    return 0;
}
// ColumnHeaders.Remove(i) —— 删掉第 i 列 (1 基)。与 IlRemoveAt 同一条纪律:
// 尾巴上那份的内存已经搬走/释放过了, 只能补一个空壳, **不能再 free 一遍**。
void vb6_ListView_RemoveColumn(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->colCount) return;
    int at = (int)idx - 1;
    lvFreeW(v->cols[at].key);
    lvFreeW(v->cols[at].text);
    if (at + 1 < v->colCount)
        memmove(&v->cols[at], &v->cols[at + 1],
                sizeof(Vb6LVCol) * (v->colCount - at - 1));
    ZeroMemory(&v->cols[v->colCount - 1], sizeof(Vb6LVCol));
    v->colCount--;
    lvSyncAll(v);
}

void vb6_ListView_ClearColumns(void* hwnd) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v) return;
    for (int i = 0; i < v->colCount; i++) { lvFreeW(v->cols[i].key); lvFreeW(v->cols[i].text); }
    v->colCount = 0;
    SendMessageW(v->hwnd, LVM_DELETEALLITEMS, 0, 0);
    // Fix 161f-extlist(3): UI 已清空 → 同步计数归零, 否则后续增量同步会以为还在。
    v->syncedCols = 0;
    v->syncedRows = 0;
}
// ---------------- ListItems (行) ----------------
int32_t vb6_ListView_GetItemCount(void* hwnd) {
    Vb6ListView* v = lvFind((HWND)hwnd); return v ? v->itemCount : 0;
}

// 行内 sub 槽扩容: 首次 HeapAlloc (HeapReAlloc 传 NULL 无效 —— 与 StatusBar 同一个坑)
static int lvItemEnsureSub(Vb6LVItem* it, int need) {
    if (need <= it->subCap) return 1;
    int nc = it->subCap ? it->subCap : 4;
    while (nc < need) nc *= 2;
    wchar_t** na;
    if (it->subs) {
        na = (wchar_t**)HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                    it->subs, sizeof(wchar_t*) * nc);
    } else {
        na = (wchar_t**)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                  sizeof(wchar_t*) * nc);
    }
    if (!na) return 0;
    it->subs = na;
    it->subCap = nc;
    return 1;
}

// ListItems.Add [index], [key], [text], [icon], [smallicon] → 返回 1 基 Index
int32_t vb6_ListView_AddItem(void* hwnd, int32_t idx, void* keyBstr, void* textBstr,
                             int32_t icon, int32_t smallIcon) {
    Vb6ListView* v = lvEnsure((HWND)hwnd);
    if (!v) return 0;
    if (v->itemCount + 1 > v->itemCap) {
        int nc = v->itemCap ? v->itemCap * 2 : 8;
        Vb6LVItem* na;
        if (v->items) {
            na = (Vb6LVItem*)HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                         v->items, sizeof(Vb6LVItem) * nc);
        } else {
            na = (Vb6LVItem*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                       sizeof(Vb6LVItem) * nc);
        }
        if (!na) return 0;
        v->items = na;
        v->itemCap = nc;
    }
    int at = (idx >= 1 && idx <= v->itemCount) ? (int)idx - 1 : v->itemCount;
    if (at < v->itemCount)
        memmove(&v->items[at + 1], &v->items[at], sizeof(Vb6LVItem) * (v->itemCount - at));
    ZeroMemory(&v->items[at], sizeof(Vb6LVItem));
    lvSetW(&v->items[at].key, (const wchar_t*)keyBstr);
    lvSetW(&v->items[at].text, (const wchar_t*)textBstr);
    v->items[at].icon = (int)icon;
    v->items[at].smallIcon = (int)smallIcon;
    v->itemCount++;
    lvSyncAll(v);
    return at + 1;
}

void* vb6_ListView_GetItemText(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount || !v->items[idx - 1].text)
        return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(v->items[idx - 1].text);
}
void vb6_ListView_SetItemText(void* hwnd, int32_t idx, void* bstr) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return;
    lvSetW(&v->items[idx - 1].text, (const wchar_t*)bstr);
    lvSyncCell(v, (int)idx - 1, 0);   // Fix 161f-extlist(3): 单格更新
}
void* vb6_ListView_GetItemKey(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount || !v->items[idx - 1].key)
        return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(v->items[idx - 1].key);
}
void vb6_ListView_SetItemKey(void* hwnd, int32_t idx, void* bstr) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return;
    lvSetW(&v->items[idx - 1].key, (const wchar_t*)bstr);
}
int32_t vb6_ListView_GetItemIndexByKey(void* hwnd, void* keyBstr) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || !keyBstr) return 0;
    const wchar_t* k = (const wchar_t*)keyBstr;
    if (!*k) return 0;
    for (int i = 0; i < v->itemCount; i++)
        if (v->items[i].key && !lstrcmpW(v->items[i].key, k)) return i + 1;
    return 0;
}

// SubItems(i): **i 从 1 开始, i=1 就是第 2 列** (第 1 列是 Text)
void* vb6_ListView_GetItemSub(void* hwnd, int32_t idx, int32_t sub) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount || sub < 1) return (void*)vb6_BSTR_Empty();
    Vb6LVItem* it = &v->items[idx - 1];
    if (sub > it->subCount || !it->subs[sub - 1]) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(it->subs[sub - 1]);
}
void vb6_ListView_SetItemSub(void* hwnd, int32_t idx, int32_t sub, void* bstr) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount || sub < 1) return;
    Vb6LVItem* it = &v->items[idx - 1];
    if (!lvItemEnsureSub(it, sub)) return;
    lvSetW(&it->subs[sub - 1], (const wchar_t*)bstr);
    if (sub > it->subCount) it->subCount = sub;
    // Fix 161f-extlist(3): 单格更新, 不走整表同步 (见 lvSyncCell 注释)。
    lvSyncCell(v, (int)idx - 1, (int)sub);
}
int32_t vb6_ListView_GetItemSubCount(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return 0;
    return v->items[idx - 1].subCount;
}
int32_t vb6_ListView_GetItemSelected(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return 0;
    return v->items[idx - 1].selected ? -1 : 0;      // VB6 True = -1
}
void vb6_ListView_SetItemSelected(void* hwnd, int32_t idx, int32_t val) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return;
    v->items[idx - 1].selected = val ? 1 : 0;
    LVITEMW st;
    ZeroMemory(&st, sizeof(st));
    st.mask = LVIF_STATE;
    st.state = (UINT)(val ? LVIS_SELECTED : 0);
    st.stateMask = LVIS_SELECTED;
    SendMessageW(v->hwnd, LVM_SETITEMSTATE, (WPARAM)(idx - 1), (LPARAM)&st);
}
int32_t vb6_ListView_GetItemChecked(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return 0;
    return v->items[idx - 1].checked ? -1 : 0;
}
void vb6_ListView_SetItemChecked(void* hwnd, int32_t idx, int32_t val) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return;
    v->items[idx - 1].checked = val ? 1 : 0;
    lvSyncAll(v);
}
// SelectedItem.Index: 第一个选中行 (无则 0)
int32_t vb6_ListView_GetSelectedIndex(void* hwnd) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v) return 0;
    for (int i = 0; i < v->itemCount; i++) if (v->items[i].selected) return i + 1;
    return 0;
}
void vb6_ListView_RemoveItem(void* hwnd, int32_t idx) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v || idx < 1 || idx > v->itemCount) return;
    int at = (int)idx - 1;
    Vb6LVItem* it = &v->items[at];
    lvFreeW(it->key);
    lvFreeW(it->text);
    lvFreeW(it->tag);
    for (int s = 0; s < it->subCount; s++) lvFreeW(it->subs[s]);
    if (it->subs) HeapFree(GetProcessHeap(), 0, it->subs);
    // 向上搬: 尾部那份已被搬到 at, 只能 memset 空壳, 绝不能再 free (StatusBar 的教训)
    if (at < v->itemCount - 1)
        memmove(&v->items[at], &v->items[at + 1], sizeof(Vb6LVItem) * (v->itemCount - at - 1));
    ZeroMemory(&v->items[v->itemCount - 1], sizeof(Vb6LVItem));
    v->itemCount--;
    lvSyncAll(v);
}
void vb6_ListView_ClearItems(void* hwnd) {
    Vb6ListView* v = lvFind((HWND)hwnd);
    if (!v) return;
    for (int i = 0; i < v->itemCount; i++) {
        Vb6LVItem* it = &v->items[i];
        lvFreeW(it->key);
        lvFreeW(it->text);
        lvFreeW(it->tag);
        for (int s = 0; s < it->subCount; s++) lvFreeW(it->subs[s]);
        if (it->subs) HeapFree(GetProcessHeap(), 0, it->subs);
        ZeroMemory(it, sizeof(*it));
    }
    v->itemCount = 0;
    SendMessageW(v->hwnd, LVM_DELETEALLITEMS, 0, 0);
    // Fix 161f-extlist(3): UI 已清空 → 同步计数归零 (增量同步前提)。
    v->syncedRows = 0;
}

// ImageList 关联 (把 #9 ImageList 的真 HIMAGELIST 喂给控件)
// ===================== C29-7: 事件通知换算 =====================
// 把一条 WM_NOTIFY 的 LVN_* 换算成 VB6 语义的 **1 基下标** (0 = 这条通知与我们无关)。
// 生成代码在 WM_NOTIFY 分支里按控件句柄匹配 hwndFrom, 再拿这个下标去取成员对象,
// 最后调 <控件名>_ItemClick / _ColumnClick。
// 两个通知都带 NMLISTVIEW:
//   LVN_ITEMACTIVATE (-114): iItem    = 被双击/回车激活的那一行
//   LVN_COLUMNCLICK  (-108): iSubItem = 被点击的那一列
// **不用** LVN_ITEMCHANGED: 它在"程序化 SetItemSelected/Checked"时也会发, 会被当成
// 用户点击 (VB6 的 ItemClick 只在用户动作时发), 判据里容易假绿。
int32_t vb6_ListView_OnNotify(void* hwnd, int32_t code, void* lParam) {
    if (!hwnd || !lParam) return 0;
    NMLISTVIEW* nmlv = (NMLISTVIEW*)lParam;
    // hwndFrom 必须就是本控件 —— 父窗会收到所有子控件的 WM_NOTIFY。
    if ((HWND)nmlv->hdr.hwndFrom != (HWND)hwnd) return 0;
    switch (code) {
    case LVN_ITEMACTIVATE:
        return (nmlv->iItem >= 0) ? nmlv->iItem + 1 : 0;
    case LVN_COLUMNCLICK:
        return (nmlv->iSubItem >= 0) ? nmlv->iSubItem + 1 : 0;
    default:
        break;
    }
    return 0;
}

void vb6_ListView_SetImageList(void* hwnd, void* himl, int32_t which) {
    if (!hwnd) return;
    // which: 0=Icons 1=SmallIcons 2=ColumnHeaderIcons
    // ⚠ 没有 `LVSIL_HEADER` 这个宏 (commctrl.h 只给 NORMAL/SMALL/STATE) —— 列头图标的
    // ImageList 挂在 **header 控件**上, 所以 which==2 要先 LVM_GETHEADER 再 HDM_SETIMAGELIST。
    // (一开始写成 LVSIL_HEADER -> C2065, 而这条只有**夹具编译 RTL** 时才暴露,
    //  build.bat 只编 C++ 看不到。)
    if (which == 2) {
        HWND hh = (HWND)SendMessageW((HWND)hwnd, LVM_GETHEADER, 0, 0);
        if (hh) SendMessageW(hh, HDM_SETIMAGELIST, 0, (LPARAM)himl);
        return;
    }
    SendMessageW((HWND)hwnd, LVM_SETIMAGELIST,
                 (which == 1) ? LVSIL_SMALL : LVSIL_NORMAL, (LPARAM)himl);
}
