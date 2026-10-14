// vb6forms_sstab.c — VB6 SSTab 控件 (TabDlg.SSTab) 复刻
//
// 用 comctl32 的 **SysTabControl32** 复刻 VB6 选项卡容器, **不加载 TABCTL32.OCX**
// (本机未注册; 且 32 位 inproc OCX 无法进 x64 进程)。
//
// 与 StatusBar(#10) 的一处重要差别: msctls_status32 两个版本的 comctl32 都不注册
// (要 RTL 自注册兜底), 而 **SysTabControl32 是注册好的** —— 实测 x64/x86 两个进程里
// GetClassInfoW 都直接成功, 连 InitCommonControlsEx 之前就已经在。所以这里不需要那个兜底。
//
// 属性表 (照 ai/内置控件/SSTab 控件（选项卡容器).md 全量实现):
//   Tabs            页总数
//   Tab             当前活动页索引 (**0 基**, 与其它 VB6 集合的 1 基不同)
//   TabCaption(i)   第 i 页标题          —— 带下标的索引属性
//   TabVisible(i)   第 i 页是否可见       —— 带下标的索引属性
//   TabOrientation  0=上(默认) 1=下 2=左 3=右
//   TabStyle        0=选项卡对话框式(默认) 1=属性页式
//   TabsPerRow      每行显示的标签数
//   WordWrap        标题过长时是否换行
//
// ===================== 容器语义: 页码从 Left 反推 =====================
//
// 这是本控件与已做的 ProgressBar/ImageList/StatusBar 最大的不同: SSTab 是**容器**。
// VB6 把每一页的子控件都摆在同一份坐标里, 靠把非活动页的子控件 Left **减 75000 缇**
// 推出可见区来隐藏。微软 KB Q187562 / Q206904 都明确写了这一点:
//   "The SSTab control hides the controls present in the inactive tabs by setting
//    their Left property values less by 75000."
// 所以 .frm 里子控件是顺序平铺在 SSTab 的 Begin..End 之内的, **没有任何显式的页号字段**,
// 页码只能从 Left 反推:
//       storedLeft = realLeft - 75000 * page
// 取让 realLeft 落回 [0, 75000) 的 page 即为页号 (阈值用 -37500, 免得把本来就有
// 小负值 Left 的控件误判到下一页)。这段反推在 codegen 里做, RTL 只认页码。
//
// 为什么不照抄 VB6 的"挪 Left"做法? 因为那样 `Debug.Print Label1.Left` 会读出一个被内部
// hack 挪过的数, 而 VB6 里这个属性就是设计期的那个值。这里改为**给子控件登记页号 +
// 切页时 ShowWindow 开关**, Left 保持原值不动, 对外口径干净。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6rtl_bstr.h"

#ifdef _WIN32

// ===================== VB6 枚举常量 =====================
#define VB6_TAB_ORIENT_TOP     0
#define VB6_TAB_ORIENT_BOTTOM  1
#define VB6_TAB_ORIENT_LEFT    2
#define VB6_TAB_ORIENT_RIGHT   3

#define VB6_TAB_STYLE_DIALOG    0   // 选项卡对话框式 (Win32 TCS_TABS)
#define VB6_TAB_STYLE_PROPPAGE  1   // 属性页式 (扁平按钮式)

// ===================== 实例 =====================

typedef struct {
    void* hwnd;      // 子控件句柄
    int   page;      // 所属页 (0 基)
} Vb6TabChild;

typedef struct {
    HWND         hwnd;         // SysTabControl32
    wchar_t**    captions;     // 平行表: 每页标题
    wchar_t**    tips;         // 平行表: 每页 ToolTipText (SSTabEx TabToolTipText(i))
    int*         visible;      // 平行表: 每页可见与否
    int*         images;       // 平行表: 每页图标在 himl 里的下标, -1=无
                                // (Task #44: SSTabEx TabPic16/20/24(i) —— 参考截图
                                //  Theme/Frame/Other/Cmd 四页带 16px 图标)
    HIMAGELIST   himl;         // 挂到 TabCtrl 上的真 image list (首次 SetTabPicture 惰性建)
    int          count;        // 页数 (逻辑页总数, 含被 TabVisible 藏掉的)
    int          cur;          // 当前活动页 (逻辑号, 0 基)
    int          prev;         // 上一次活动页 —— Click(PreviousTab) 要这个值
    int          orientation;
    int          tabStyle;
    int          tabsPerRow;
    int          wordWrap;
    Vb6TabChild* children;
    int          childCount;
    int          childCap;
} Vb6SSTab;

// 实例表: 窗体数量级, 线性查找足够 (与 vb6forms_statusbar.c 同口径)。
#define VB6_SSTAB_MAX 64
static Vb6SSTab* g_tabs[VB6_SSTAB_MAX];
static int       g_tabCount = 0;

static Vb6SSTab* sstabFind(HWND h) {
    if (!h) return NULL;
    for (int i = 0; i < g_tabCount; i++)
        if (g_tabs[i] && g_tabs[i]->hwnd == h) return g_tabs[i];
    return NULL;
}

static Vb6SSTab* sstabEnsure(HWND h) {
    Vb6SSTab* t = sstabFind(h);
    if (t) return t;
    if (g_tabCount >= VB6_SSTAB_MAX) return NULL;
    t = (Vb6SSTab*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(Vb6SSTab));
    if (!t) return NULL;
    t->hwnd = h;
    t->cur = 0;
    t->prev = 0;
    t->orientation = VB6_TAB_ORIENT_TOP;
    t->tabStyle = VB6_TAB_STYLE_DIALOG;
    g_tabs[g_tabCount++] = t;
    return t;
}

// ---- 逻辑页 ↔ 视图项 ------------------------------------------------
//
// TabVisible(i)=False 的页在 VB6 里是连标签带内容一起消失, 而 Win32 TabCtrl **没有
// per-item 可见位**, 只能"不插入"。代价是 view 下标与逻辑页号错位, 所以下面这两张
// 方向的换算是唯一入口 —— 凡是碰 GetCurSel/SetCurSel 的返回值都不能直接拿去当页号用。
static int sstabLogicalToView(const Vb6SSTab* t, int logical) {
    if (!t || logical < 0 || logical >= t->count) return -1;
    if (!t->visible || !t->visible[logical]) return -1;
    int v = 0;
    for (int i = 0; i < logical; i++)
        if (t->visible[i]) v++;
    return v;
}

static int sstabViewToLogical(const Vb6SSTab* t, int view) {
    if (!t || !t->visible || view < 0) return -1;
    for (int i = 0; i < t->count; i++) {
        if (!t->visible[i]) continue;
        if (view == 0) return i;
        view--;
    }
    return -1;
}

// 重建全部可见标签项。改了任何影响布局的样式之后必须走一遍 —— 否则旧 item 尺寸不重算。
// Task #44: 有 image list 时 item 要带 TCIF_IMAGE —— 没有 image 的页不设 iImage,
// comctl32 就不画图标 (不能填 0, 0 是第一张图的合法下标)。
static void sstabRebuildVisibleTabs(Vb6SSTab* t) {
    if (!t || !t->hwnd) return;
    TabCtrl_DeleteAllItems(t->hwnd);
    if (!t->captions) return;
    for (int i = 0; i < t->count; i++) {
        if (t->visible && !t->visible[i]) continue;
        TCITEMW ti;
        ZeroMemory(&ti, sizeof(ti));
        ti.mask = TCIF_TEXT;
        ti.pszText = t->captions[i] ? t->captions[i] : (wchar_t*)L"";
        if (t->himl && t->images && t->images[i] >= 0) {
            ti.mask |= TCIF_IMAGE;
            ti.iImage = t->images[i];
        }
        TabCtrl_InsertItem(t->hwnd, TabCtrl_GetItemCount(t->hwnd), &ti);
    }
}

// VB6 的样位 → TCS_*。位置类样式 (TCS_BOTTOM / TCS_VERTICAL / TCS_RIGHT) 只在
// 创建时给最稳, 运行期改靠 SetWindowLongPtr + 重建 item 强制重排。
static void sstabApplyRowMetrics(Vb6SSTab* t);   // 定义在下面, sstabReapplyStyle 用得到
static void sstabPostSelChange(Vb6SSTab* t);  // SetTab 末尾补发 TCN_SELCHANGE, 定义在下方

static LONG sstabStyleBits(const Vb6SSTab* t, LONG base) {
    LONG s = base;
    if (t->orientation == VB6_TAB_ORIENT_BOTTOM) s |= TCS_BOTTOM;
    if (t->orientation == VB6_TAB_ORIENT_LEFT)   s |= TCS_VERTICAL;
    if (t->orientation == VB6_TAB_ORIENT_RIGHT)  s |= TCS_VERTICAL | TCS_RIGHT;
    if (t->tabStyle == VB6_TAB_STYLE_PROPPAGE)   s |= TCS_BUTTONS | TCS_FLATBUTTONS;
    if (t->wordWrap)                             s |= TCS_MULTILINE;
    return s;
}

#define SST_STYLE_MASK (TCS_BOTTOM | TCS_VERTICAL | TCS_RIGHT | TCS_BUTTONS \
                        | TCS_FLATBUTTONS | TCS_MULTILINE)

static void sstabReapplyStyle(Vb6SSTab* t) {
    if (!t || !t->hwnd) return;
    LONG st = (LONG)GetWindowLongPtr(t->hwnd, GWL_STYLE);
    SetWindowLongPtr(t->hwnd, GWL_STYLE, sstabStyleBits(t, st & ~SST_STYLE_MASK));
    sstabRebuildVisibleTabs(t);
    sstabApplyRowMetrics(t);
}

// TabsPerRow: Win32 没有"第行几个标签"的直接位, 折成最小标签宽
// (可视宽 / 每行个数), 让换行按 VB6 的意图走。0 或非法值 = 不限制。
static void sstabApplyRowMetrics(Vb6SSTab* t) {
    if (!t || !t->hwnd) return;
    if (t->tabsPerRow <= 0) { TabCtrl_SetMinTabWidth(t->hwnd, 0); return; }
    RECT rc;
    GetClientRect(t->hwnd, &rc);
    int w = rc.right - rc.left;
    if (w <= 0) return;
    // Task #44: VB6 SSTab 的行分组是**固定每行 tabsPerRow 个** (4+3), 不随
    // caption 宽度变。comctl v6 (主题开) 换行按自然宽 + 主题内边距算, minTabWidth
    // 只是显示下限 —— 调它换行点不动 (实测 3+4 纹丝不动)。多行时改用
    // TCS_FIXEDWIDTH + 显式项宽 = 可用宽/perRow: 每行**恰好**容 perRow 个
    // (4×W ≤ avail < 5×W 两侧余量都足), 单行 (count ≤ perRow) 保持自然宽。
    LONG st = (LONG)GetWindowLongPtrW(t->hwnd, GWL_STYLE);
    if (t->wordWrap && t->tabsPerRow > 0 && t->count > t->tabsPerRow) {
        if (!(st & TCS_FIXEDWIDTH)) {
            SetWindowLongPtrW(t->hwnd, GWL_STYLE, st | TCS_FIXEDWIDTH);
        }
        int minW = (w - 24) / t->tabsPerRow;   // 扣横向内边距 (DPI 余量)
        if (minW < 8) minW = 8;
        int h = 0;
        RECT ir;
        if (TabCtrl_GetItemRect(t->hwnd, 0, &ir) && ir.bottom > ir.top)
            h = (int)(ir.bottom - ir.top);
        TabCtrl_SetItemSize(t->hwnd, (LPARAM)minW, (LPARAM)h);
    } else {
        if (st & TCS_FIXEDWIDTH) {
            SetWindowLongPtrW(t->hwnd, GWL_STYLE, st & ~TCS_FIXEDWIDTH);
            TabCtrl_SetItemSize(t->hwnd, 0, 0);   // 回归自然宽
        }
        int minW = w / t->tabsPerRow;
        if (minW < 8) minW = 8;
        TabCtrl_SetMinTabWidth(t->hwnd, minW);
    }
}

// 切页后按"登记的页号"开关子控件 —— 见文件头: 不动 Left, 只动可见性。
static void sstabRefreshChildren(Vb6SSTab* t) {
    if (!t) return;
    for (int i = 0; i < t->childCount; i++) {
        HWND c = (HWND)t->children[i].hwnd;
        if (!c) continue;
        int p = t->children[i].page;
        // 两种隐藏: 不是当前页 / 该页被 TabVisible=False 藏了
        int show = (p == t->cur && p >= 0 && p < t->count
                    && (!t->visible || t->visible[p]));
        ShowWindow(c, show ? SW_SHOW : SW_HIDE);
    }
}

// ===================== 初始化 =====================

// Task #44: SSTab 行宽 (TabsPerRow → 最小标签宽) 依赖**客户区宽度**, 而 Init 时
// 窗口尚未布局 (GetClientRect 宽 0) → 折行按 caption 自然宽走, 7 页折成 3+4 而
// 非 VB6 的 4+3 (SSTabEx frmTest 实证)。挂一个内部子类过程, WM_SIZE 时重算。
static LRESULT CALLBACK sstabSubclassProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    if (msg == WM_ERASEBKGND) {
        // 防闪烁 (第三步): 上一版"直接 return 1 跳过擦除"治好了页内容, 却暴露了
        // 新问题 —— 多行标签 (TCS_MULTILINE, 3+4 两行) 时**与显示区相邻的那一行**
        // 主题绘制要先擦"行间缝隙/标签背后"再画按钮, 跳过统一擦底后新旧像素交替,
        // 用户实测: 上面 3 个不闪、下面 4 个闪 (闪的正是紧贴显示区那行)。
        // 解法: 不跳过擦除, 改为**用父窗体注册类的背景刷擦全客户区**。主题标签区
        // 底色本就是窗体背景色 (非主题路径 SysTabControl32 也是向父窗口要
        // WM_CTLCOLORSTATIC 拿底色), 擦除色与绘制后的底色一致 → 擦/画跳变不可见
        // → 不闪; 页内子控件仍由 WS_CLIPCHILDREN 保护, 不会被这层擦除波及。
        HDC hdc = (HDC)wp;
        RECT rc;
        GetClientRect(hwnd, &rc);
        HBRUSH br = NULL;
        HWND parent = GetParent(hwnd);
        if (parent) br = (HBRUSH)GetClassLongPtrW(parent, GCLP_HBRBACKGROUND);
        if (!br) br = (HBRUSH)GetSysColorBrush(COLOR_BTNFACE);
        FillRect(hdc, &rc, br);
        return 1;
    }
    if (msg == WM_CTLCOLORSTATIC) {
        // Fix 187: 页内 STATIC 类子控件 (picTabBackColor 等色块, parent=SSTab
        // 窗口) 通过 WM_CTLCOLORSTATIC 向本容器要底色 — 此前不处理走原生默认,
        // vb6_SetControlBackColor 写的 VB6_BackColor 属性无人消费。与主窗体
        // WndProc 同款: 仅在子控件显式 Set 过 VB6_BackColor 时接管。
        // Fix <c3-menu3d-labelbg>: 与主窗体 WndProc 同步放行 VB6_BackStyle0
        // (Label.BackStyle=0 Transparent) —— 这种 Label 一定没有 BackColor,
        // 只按 BackColorSet 放行就会落到 DefWindowProcW 被 STATIC 类背景刷成白块。
        HWND child = (HWND)lp;
        if (child && (GetPropW(child, L"VB6_BackColorSet") || GetPropW(child, L"VB6_BackStyle0"))) {
            LRESULT br187 = vb6_ApplyCtlColorStatic((HDC)wp, child);
            if (br187) return br187;
        }
    }
    if (msg == WM_SIZE) {
        Vb6SSTab* t = sstabFind(hwnd);
        if (t) sstabApplyRowMetrics(t);
    }
    WNDPROC orig = (WNDPROC)GetPropW(hwnd, L"VB6_SSTab_OrigProc");
    return orig ? CallWindowProcW(orig, hwnd, msg, wp, lp)
                : DefWindowProcW(hwnd, msg, wp, lp);
}

void vb6_SSTab_Init(void* tabHwnd, int tabs, int curTab, int orientation, int tabStyle,
                    int tabsPerRow, int wordWrap) {
    if (!tabHwnd) return;
    Vb6SSTab* t = sstabEnsure((HWND)tabHwnd);
    if (!t) return;

    // 内部子类 (只装一次): WM_SIZE → sstabApplyRowMetrics
    if (!GetPropW(t->hwnd, L"VB6_SSTab_OrigProc")) {
        WNDPROC op = (WNDPROC)SetWindowLongPtrW(t->hwnd, GWLP_WNDPROC,
                                                (LONG_PTR)sstabSubclassProc);
        if (op) SetPropW(t->hwnd, L"VB6_SSTab_OrigProc", (HANDLE)op);
    }

    // 防闪烁 (SSTabEx frmTest 实测: 内容最多的左页最明显):
    // 页内子控件是 SysTabControl32 的**子窗口** (cgen_form_frame_menu.inc 用
    // containerHwnd 当父)。tab 控件自身没 WS_CLIPCHILDREN 时, 它每次重绘显示区
    // 背景会把子控件区域一并擦掉再让子控件重画 → 切页/悬停时页面内容闪烁。
    // 补上裁剪样式即可 (SST_STYLE_MASK 不含这两位, 后面 sstabReapplyStyle 的
    // SetWindowLongPtr 不会把它洗掉)。
    {
        LONG cs = (LONG)GetWindowLongPtrW(t->hwnd, GWL_STYLE);
        if (!(cs & WS_CLIPCHILDREN)) {
            SetWindowLongPtrW(t->hwnd, GWL_STYLE, cs | WS_CLIPCHILDREN | WS_CLIPSIBLINGS);
            SetWindowPos(t->hwnd, NULL, 0, 0, 0, 0,
                         SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER |
                         SWP_FRAMECHANGED | SWP_NOACTIVATE);
        }
    }

    // 防闪烁 (第三步遗留结论, 已移除 WS_EX_COMPOSITED):
    // 曾给顶层窗体加 WS_EX_COMPOSITED 治"窗体↔控件合成时序闪", 但它把整棵窗口树
    // 转入重定向合成 —— DWM 合成时向子控件要内容走 WM_PRINTCLIENT 而非 WM_PAINT,
    // 自绘 Frame 等不响应 WM_PRINTCLIENT 的控件被画成未初始化黑块 (用户实测回归)。
    // 且闪烁真根因是外部定时器同值赋值触发全量重建 (见各 setter 的幂等早退),
    // 与合成时序无关 → COMPOSITED 弊大于利, 移除。

    t->orientation = orientation;
    t->tabStyle = tabStyle;
    t->tabsPerRow = tabsPerRow;
    t->wordWrap = wordWrap;

    // 清掉旧平行表, 重新按 tabs 建
    if (t->captions) {
        for (int i = 0; i < t->count; i++)
            if (t->captions[i]) HeapFree(GetProcessHeap(), 0, t->captions[i]);
        HeapFree(GetProcessHeap(), 0, t->captions);
        t->captions = NULL;
    }
    if (t->tips) {
        for (int i = 0; i < t->count; i++)
            if (t->tips[i]) HeapFree(GetProcessHeap(), 0, t->tips[i]);
        HeapFree(GetProcessHeap(), 0, t->tips);
        t->tips = NULL;
    }
    if (t->visible) { HeapFree(GetProcessHeap(), 0, t->visible); t->visible = NULL; }
    if (t->images)  { HeapFree(GetProcessHeap(), 0, t->images);  t->images = NULL; }
    t->count = 0;

    if (tabs > 0) {
        t->captions = (wchar_t**)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(wchar_t*) * tabs);
        t->tips = (wchar_t**)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(wchar_t*) * tabs);
        t->visible = (int*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(int) * tabs);
        t->images = (int*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(int) * tabs);
        // images 缺省 0 会指向 image list 的**第一张图** (合法下标), 必须 -1 = 无图标
        if (t->captions && t->tips && t->visible && t->images) {
            for (int i = 0; i < tabs; i++) {
                t->captions[i] = NULL;
                t->tips[i] = NULL;
                t->visible[i] = 1;
                t->images[i] = -1;
            }
            t->count = tabs;
        }
    }

    sstabReapplyStyle(t);

    t->cur = 0;
    if (curTab > 0 && curTab < t->count) t->cur = curTab;
    if (t->visible && t->cur < t->count && !t->visible[t->cur]) {
        // VB6 不会让活动页落在隐藏页上; 顺延到第一个可见页。
        int first = -1;
        for (int i = 0; i < t->count; i++) if (t->visible[i]) { first = i; break; }
        t->cur = (first < 0) ? 0 : first;
    }
    t->prev = t->cur;
    int view = sstabLogicalToView(t, t->cur);
    if (view >= 0) TabCtrl_SetCurSel(t->hwnd, view);
    sstabRefreshChildren(t);
}

// --- 运行期样式重算出口: 改了 orientation/style/wordWrap 之后统一走这里 ---
static void sstabAfterVisualChange(Vb6SSTab* t) {
    sstabReapplyStyle(t);
    int view = sstabLogicalToView(t, t->cur);
    if (view >= 0) TabCtrl_SetCurSel(t->hwnd, view);
    sstabRefreshChildren(t);
}

// ===================== 属性: Tabs / Tab =====================

// Tabs 变多时补空页, 变少时截掉 —— 与 StatusBar 的 Panels 不同, TabCtrl 的 item 数
// 就是页本身, 没有平行集合对象, 所以 count 与视图 item 数必须一起动。
static void sstabResize(Vb6SSTab* t, int n) {
    if (!t || n < 0 || n == t->count) return;
    wchar_t** nc = NULL;
    wchar_t** nt = NULL;
    int* nv = NULL;
    int* ni = NULL;
    if (n > 0) {
        nc = (wchar_t**)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(wchar_t*) * n);
        nt = (wchar_t**)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(wchar_t*) * n);
        nv = (int*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(int) * n);
        ni = (int*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(int) * n);
        if (!nc || !nt || !nv || !ni) {
            if (nc) HeapFree(GetProcessHeap(), 0, nc);
            if (nt) HeapFree(GetProcessHeap(), 0, nt);
            if (nv) HeapFree(GetProcessHeap(), 0, nv);
            if (ni) HeapFree(GetProcessHeap(), 0, ni);
            return;
        }
        int keep = (n < t->count) ? n : t->count;
        for (int i = 0; i < keep; i++) {
            nc[i] = t->captions[i];
            nt[i] = t->tips ? t->tips[i] : NULL;
            nv[i] = t->visible[i];
            ni[i] = t->images ? t->images[i] : -1;
        }
        // 缩容时**必须**显式释放被截掉那几页的字符串 (否则泄漏), 且只能释放
        // 超出新长度的那些 —— 保留段已经搬进 nc/nt, 不能碰。
        for (int i = keep; i < t->count; i++) {
            if (t->captions[i]) HeapFree(GetProcessHeap(), 0, t->captions[i]);
            if (t->tips && t->tips[i]) HeapFree(GetProcessHeap(), 0, t->tips[i]);
        }
        for (int i = keep; i < n; i++) { nv[i] = 1; ni[i] = -1; }
    } else {
        for (int i = 0; i < t->count; i++) {
            if (t->captions[i]) HeapFree(GetProcessHeap(), 0, t->captions[i]);
            if (t->tips && t->tips[i]) HeapFree(GetProcessHeap(), 0, t->tips[i]);
        }
    }
    if (t->captions) HeapFree(GetProcessHeap(), 0, t->captions);
    if (t->tips) HeapFree(GetProcessHeap(), 0, t->tips);
    if (t->visible) HeapFree(GetProcessHeap(), 0, t->visible);
    if (t->images) HeapFree(GetProcessHeap(), 0, t->images);
    t->captions = nc;
    t->tips = nt;
    t->visible = nv;
    t->images = ni;
    t->count = n;
    if (t->cur >= n) t->cur = (n > 0) ? n - 1 : 0;
}

int32_t vb6_SSTab_GetTabs(void* tabHwnd) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    return t ? (int32_t)t->count : 0;
}

void vb6_SSTab_SetTabs(void* tabHwnd, int32_t n) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t) return;
    // VB6/OCX 语义: 同值赋值是 no-op —— SSTabEx frmTest 的 tmrUpdate 每 1.5s
    // 执行 Tabs = 7 (值不变), OCX 直接返回; 这里若照常 sstabAfterVisualChange
    // 会把整个 tab 控件 DeleteAllItems+重建+子控件整轮隐藏/显示 → 带图标的那行
    // 标签每 1.5s 闪一次 (用户实测 Theme/Frame/Other/Cmd 闪)。必须幂等早退。
    if ((int)n == t->count) return;
    sstabResize(t, (int)n);
    sstabAfterVisualChange(t);
}

int32_t vb6_SSTab_GetTab(void* tabHwnd) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    return t ? (int32_t)t->cur : 0;
}

void vb6_SSTab_SetTab(void* tabHwnd, int32_t idx) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t) return;
    if (idx < 0 || idx >= t->count) return;
    if (t->visible && !t->visible[idx]) return;   // VB6: 不能切到隐藏页
    t->prev = t->cur;
    t->cur = (int)idx;
    int view = sstabLogicalToView(t, t->cur);
    if (view >= 0) TabCtrl_SetCurSel(t->hwnd, view);
    sstabRefreshChildren(t);

    // VB6 的 SSTab 在**页被切换**时就触发 Click(PreviousTab), 程序化改 Tab 也算。
    // 而 comctl32 只在**用户点标签**时才自己发 TCN_SELCHANGE —— 实测程序化
    // TabCtrl_SetCurSel 一点通知都不发, 于是 VB 侧的 Click 永远不触发。
    // 所以这里补发一条: 让 SetTab 与真实点击走同一条 WM_NOTIFY 路径。
    // 不会递归: OnSelChange 不调 SetTab。
    sstabPostSelChange(t);
}

// 给父窗口发一条 TCN_SELCHANGE —— 与 comctl32 在用户点击时发的那条同构。
static void sstabPostSelChange(Vb6SSTab* t) {
    if (!t || !t->hwnd) return;
    HWND parent = GetParent(t->hwnd);
    if (!parent) return;
    NMHDR nm;
    ZeroMemory(&nm, sizeof(nm));
    nm.hwndFrom = t->hwnd;
    nm.idFrom = (UINT_PTR)GetWindowLongPtr(t->hwnd, GWLP_ID);
    nm.code = TCN_SELCHANGE;
    SendMessageW(parent, WM_NOTIFY, (WPARAM)nm.idFrom, (LPARAM)&nm);
}

// ===================== 属性: TabCaption(i) / TabVisible(i) =====================
//
// 这两个是"带下标的索引属性", 不是集合对象 (VB6 里写作 SSTab1.TabCaption(0))。
// 与 StatusBar 的 Panels(i).Text 不同, 后者是集合成员的属性, 前者是控件自身的
// 参数化属性 —— codegen 侧也得走不同的分支。

void* vb6_SSTab_GetTabCaption(void* tabHwnd, int32_t idx) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || idx < 0 || idx >= t->count) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(t->captions[idx] ? t->captions[idx] : L"");
}

void vb6_SSTab_SetTabCaption(void* tabHwnd, int32_t idx, void* bstr) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || idx < 0 || idx >= t->count) return;
    const wchar_t* s = (const wchar_t*)bstr;
    // 同文本 no-op: TabCtrl_SetItem 即使文字没变也会重绘该项,
    // 演示代码循环写同一标题时会闪标签。
    if ((s && *s) == (t->captions[idx] != NULL) &&
        (!s || !*s || wcscmp(s, t->captions[idx]) == 0)) return;
    if (t->captions[idx]) { HeapFree(GetProcessHeap(), 0, t->captions[idx]); t->captions[idx] = NULL; }
    if (s && *s) {
        int n = (int)lstrlenW(s);
        wchar_t* cp = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, sizeof(wchar_t) * (n + 1));
        if (cp) { memcpy(cp, s, sizeof(wchar_t) * n); cp[n] = 0; t->captions[idx] = cp; }
    }
    // 只改文字, 不必重建整个样式, 但视图里的 item 要跟着更新
    int view = sstabLogicalToView(t, idx);
    if (view >= 0) {
        TCITEMW ti;
        ZeroMemory(&ti, sizeof(ti));
        ti.mask = TCIF_TEXT;
        ti.pszText = t->captions[idx] ? t->captions[idx] : (wchar_t*)L"";
        TabCtrl_SetItem(t->hwnd, view, &ti);
    }
}

// TabToolTipText(i): SSTabEx 的每页悬停提示。属性语义先保住 (读回设计/运行所存的值);
// 真正的鼠标悬停气泡要 TCS_TOOLTIPS + TTN_GETDISPINFO 派发, 见文件尾 TODO。
void* vb6_SSTab_GetTabToolTipText(void* tabHwnd, int32_t idx) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || idx < 0 || idx >= t->count) return (void*)vb6_BSTR_Empty();
    return (void*)vb6_BSTR_FromStr(t->tips && t->tips[idx] ? t->tips[idx] : L"");
}

void vb6_SSTab_SetTabToolTipText(void* tabHwnd, int32_t idx, void* bstr) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || idx < 0 || idx >= t->count || !t->tips) return;
    const wchar_t* s = (const wchar_t*)bstr;
    if (t->tips[idx]) { HeapFree(GetProcessHeap(), 0, t->tips[idx]); t->tips[idx] = NULL; }
    if (s && *s) {
        int n = (int)lstrlenW(s);
        wchar_t* cp = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, sizeof(wchar_t) * (n + 1));
        if (cp) { memcpy(cp, s, sizeof(wchar_t) * n); cp[n] = 0; t->tips[idx] = cp; }
    }
}

int32_t vb6_SSTab_GetTabVisible(void* tabHwnd, int32_t idx) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || idx < 0 || idx >= t->count) return 0;
    return t->visible ? (int32_t)(t->visible[idx] ? -1 : 0) : 0;   // VB6 True = -1
}

void vb6_SSTab_SetTabVisible(void* tabHwnd, int32_t idx, int32_t v) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || idx < 0 || idx >= t->count || !t->visible) return;
    int nv = v ? 1 : 0;
    if (t->visible[idx] == nv) return;      // 同值赋值 no-op
    t->visible[idx] = nv;
    if (t->cur == idx && !t->visible[idx]) {
        int first = -1;
        for (int i = 0; i < t->count; i++) if (t->visible[i]) { first = i; break; }
        t->cur = (first < 0) ? 0 : first;
    }
    // 隐藏页会改变"视图项"的个数与编号 → 必须整表重建, 不能只改一个 item
    sstabAfterVisualChange(t);
}

// ===================== 属性: 外观四件套 =====================

int32_t vb6_SSTab_GetTabOrientation(void* tabHwnd) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    return t ? (int32_t)t->orientation : 0;
}

void vb6_SSTab_SetTabOrientation(void* tabHwnd, int32_t v) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || v < 0 || v > 3) return;
    if (t->orientation == (int)v) return;   // 同值赋值 no-op (防演示代码循环触发重绘)
    t->orientation = (int)v;
    sstabAfterVisualChange(t);
}

int32_t vb6_SSTab_GetTabStyle(void* tabHwnd) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    return t ? (int32_t)t->tabStyle : 0;
}

void vb6_SSTab_SetTabStyle(void* tabHwnd, int32_t v) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || v < 0 || v > 1) return;
    if (t->tabStyle == (int)v) return;      // 同值赋值 no-op
    t->tabStyle = (int)v;
    sstabAfterVisualChange(t);
}

int32_t vb6_SSTab_GetTabsPerRow(void* tabHwnd) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    return t ? (int32_t)t->tabsPerRow : 0;
}

void vb6_SSTab_SetTabsPerRow(void* tabHwnd, int32_t v) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || v < 0) return;
    if (t->tabsPerRow == (int)v) return;    // 同值赋值 no-op
    t->tabsPerRow = (int)v;
    sstabApplyRowMetrics(t);
}

int32_t vb6_SSTab_GetWordWrap(void* tabHwnd) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    return t ? (int32_t)(t->wordWrap ? -1 : 0) : 0;
}

void vb6_SSTab_SetWordWrap(void* tabHwnd, int32_t v) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t) return;
    int nv = v ? 1 : 0;
    if (t->wordWrap == nv) return;          // 同值赋值 no-op
    t->wordWrap = nv;
    sstabAfterVisualChange(t);
}

// ===================== 容器: 子控件登记 =====================

// codegen 对每个 .frm 里落在 SSTab Begin..End 内的子控件调一次, page 由 Left 反推。
void vb6_SSTab_RegisterChild(void* tabHwnd, void* childHwnd, int32_t page) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || !childHwnd) return;
    if (t->childCount + 1 > t->childCap) {
        int nc = t->childCap ? t->childCap * 2 : 8;
        // **首次分配必须走 HeapAlloc**: HeapReAlloc 要求 lpMem 是先前由
        // HeapAlloc/HeapReAlloc 返回的指针, 传 NULL 直接失败返回 NULL —— 于是
        // 子控件一个都登记不上, 切页时谁也不会显示 (表现是 Visible 恒 0,
        // 而 Left 看着正常, 极易误判成"页码反推错了")。
        Vb6TabChild* na;
        if (t->children) {
            na = (Vb6TabChild*)HeapReAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                           t->children, sizeof(Vb6TabChild) * nc);
        } else {
            na = (Vb6TabChild*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY,
                                         sizeof(Vb6TabChild) * nc);
        }
        if (!na) return;
        t->children = na;
        t->childCap = nc;
    }
    t->children[t->childCount].hwnd = childHwnd;
    t->children[t->childCount].page = (int)page;
    t->childCount++;
    sstabRefreshChildren(t);
}

// ===================== 事件: Click(PreviousTab) =====================

// 用户在窗体 WndProc 收到 TCN_SELCHANGE 后调这个: 把 Win32 的视图下标换算成逻辑页号,
// 刷新子控件可见性, 并返回**切换前**的页号 (VB6 的 Click 参数就是它)。
int32_t vb6_SSTab_OnSelChange(void* tabHwnd) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t) return 0;
    int view = (int)TabCtrl_GetCurSel(t->hwnd);
    int logical = sstabViewToLogical(t, view);
    if (logical < 0) return 0;
    int prev = t->prev;
    t->prev = logical;
    t->cur = logical;
    sstabRefreshChildren(t);
    return (int32_t)prev;
}

// ===================== 属性: TabPicture(i) — SSTabEx 每页图标 (Task #44) =====================
//
// .frm 的 TabPic16/20/24(i) = "xxx.frx":偏移 是 SSTabEx 的三档设计期位图 (16/20/24px),
// codegen 读出字节后逐页喂进来。这里建一把**真 HIMAGELIST** 挂到 SysTabControl32 上
// (TabCtrl_SetImageList), sstabRebuildVisibleTabs 插 item 时带 TCIF_IMAGE ——
// 参考截图 (sstabex.png) 里 Theme/Frame/Other/Cmd 四页标签左侧的图标就是这么来的。
//
// mask 口径: SSTabEx 的 MaskColor 缺省 = 品红 (&HFF00FF, 参考截图 MaskColor 色块同色),
// frmTest.frx 的 TabPic 位图四角实测就是 (255,0,255) → ImageList_AddMasked 用品红抠背景。
// 图标区固定 16x16 (96DPI 走 TabPic16 档; codegen 优先取 16, 缺档顺延 20/24, 大图被缩)。
void vb6_SSTab_SetTabPicture(void* tabHwnd, int32_t idx, const void* data, int32_t size) {
    Vb6SSTab* t = sstabFind((HWND)tabHwnd);
    if (!t || !t->hwnd || !t->images || idx < 0 || idx >= t->count) return;
    if (!data || size <= 0) return;

    if (!t->himl) {
        t->himl = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 4, 8);
        if (!t->himl) return;
        TabCtrl_SetImageList(t->hwnd, t->himl);
    }

    // vb6_LoadPictureFromMemory: BMP → 复制出的 HBITMAP (所有权归调用方); ICO → HICON。
    void* pic = vb6_LoadPictureFromMemory(data, size);
    if (!pic) return;
    BITMAP bm;
    if (GetObjectW((HBITMAP)pic, sizeof(bm), &bm) == 0) {
        // 不是 HBITMAP (ICO/CURSOR 等) → ImageList_AddMasked 吃不了。TabPic 实际只存
        // BMP (.frx 实证), 这里丢弃并按语义清理句柄 (DestroyIcon 对位图句柄无害失败)。
        DestroyIcon((HICON)pic);
        return;
    }
    int imgIdx = ImageList_AddMasked(t->himl, (HBITMAP)pic, RGB(255, 0, 255));
    DeleteObject((HBITMAP)pic);
    if (imgIdx < 0) return;
    t->images[idx] = imgIdx;

    // 图标改变标签内容区尺寸 → 重建 item (带 TCIF_IMAGE) + 重算行宽/行高。
    // 不走 SetTab: 选中页不该因为贴图变化而跳。
    sstabAfterVisualChange(t);
}

// ===================== 属性: SSTabEx 颜色族 (Task #44) =====================
// 参照 ctlSSTabEx.ctl (SSTabEx-main/control-source) 的真实缺省:
//   MaskColor       = &HFF00FF 品红常量 (ctl:2661)
//   TabBackColor    = Ambient.BackColor (ctl:2667) — 宿主容器背景色
//   TabSelBackColor = Ambient.BackColor (ctl:2668)
//   TabSelForeColor = ForeColor 缺省 = Ambient.ForeColor (ctl:2646) — 宿主容器前景色
// (BackColor/ForeColor 本身是通用控件属性, cgen 直接走 vb6_Get/SetControlBackColor
//  窗口属性链不落本表; 其缺省 BTNFACE/黑 恰与 Ambient 一致, 实测吻合。)
// Ambient 实取: GetParent(宿主窗体) 的控件级颜色 —— frmTest 窗体无设计色时
// vb6_GetControlBackColor 返回 GetSysColor(COLOR_BTNFACE), 与 VB6 &H8000000F 等效。
// 不接这批表的话 cgen 发 vb6_ComGetStringProp/ComSetProp(裸 HWND, L"TabBackColor")
// → 读静默答空 / 写静默丢 (029 §九同款, 生成 C 实证)。
static int32_t sstabAmbientColor(HWND hwnd, int fore) {
    HWND par = GetParent(hwnd);
    if (!par) par = GetAncestor(hwnd, GA_ROOT);
    return fore ? vb6_GetControlForeColor(par) : vb6_GetControlBackColor(par);
}
static int32_t sstabGetColorProp(HWND hwnd, const wchar_t* prop, int32_t dflt) {
    HANDLE h = GetPropW(hwnd, prop);
    return h ? (int32_t)(INT_PTR)h : dflt;
}
int32_t vb6_SSTab_GetMaskColor(void* tabHwnd) {
    if (!tabHwnd) return 0;
    return sstabGetColorProp((HWND)tabHwnd, L"VB6_SS_MaskColor", 0xFF00FF);
}
void vb6_SSTab_SetMaskColor(void* tabHwnd, int32_t c) {
    if (!tabHwnd) return;
    SetPropW((HWND)tabHwnd, L"VB6_SS_MaskColor", (HANDLE)(INT_PTR)c);
}
int32_t vb6_SSTab_GetTabBackColor(void* tabHwnd) {
    if (!tabHwnd) return 0;
    return sstabGetColorProp((HWND)tabHwnd, L"VB6_SS_TabBackColor",
                             sstabAmbientColor((HWND)tabHwnd, 0));
}
void vb6_SSTab_SetTabBackColor(void* tabHwnd, int32_t c) {
    if (!tabHwnd) return;
    SetPropW((HWND)tabHwnd, L"VB6_SS_TabBackColor", (HANDLE)(INT_PTR)c);
}
int32_t vb6_SSTab_GetTabSelBackColor(void* tabHwnd) {
    if (!tabHwnd) return 0;
    return sstabGetColorProp((HWND)tabHwnd, L"VB6_SS_TabSelBackColor",
                             sstabAmbientColor((HWND)tabHwnd, 0));
}
void vb6_SSTab_SetTabSelBackColor(void* tabHwnd, int32_t c) {
    if (!tabHwnd) return;
    SetPropW((HWND)tabHwnd, L"VB6_SS_TabSelBackColor", (HANDLE)(INT_PTR)c);
}
int32_t vb6_SSTab_GetTabSelForeColor(void* tabHwnd) {
    if (!tabHwnd) return 0;
    return sstabGetColorProp((HWND)tabHwnd, L"VB6_SS_TabSelForeColor",
                             sstabAmbientColor((HWND)tabHwnd, 1));
}
void vb6_SSTab_SetTabSelForeColor(void* tabHwnd, int32_t c) {
    if (!tabHwnd) return;
    SetPropW((HWND)tabHwnd, L"VB6_SS_TabSelForeColor", (HANDLE)(INT_PTR)c);
}

#endif  // _WIN32
