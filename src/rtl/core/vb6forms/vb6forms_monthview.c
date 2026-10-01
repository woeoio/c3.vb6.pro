// vb6forms_monthview.c — ai/029 C29-MV-a: VB6 MonthView 的窗口 + 样式 + 标量属性面
//
// 复刻口径（ai/内置控件/MonthView 控件（月历面板）.md + 029 §三 D6）：
//   **不加载 MSCOMCT2.OCX** —— 那是 32 位 inproc，x64 进程里 CoCreateInstance 直接失败。
//   原生等价类 = comctl32 注册的 SysMonthCal32（头 6288 行 MONTHCAL_CLASSW），类已由
//   vb6forms.c 那次 InitCommonControlsEx 的 ICC_DATE_CLASSES 请求过 ⇒ 与 DTPicker/TreeView
//   同型，不需要 RTL 自注册兜底。
//
// 1..22 行那些函数 = MV-a 的标量面；Date 型那三格（Value / SelStart / SelEnd）在下面的
// MV-b 段里；MCN_SELCHANGE(-749) 那条事件（VB6 的 DateClick）留 C29-MV-c。
//
// VB6 属性 → 原生落点（数值全部抄自 D:\Windows Kits\10\Include\10.0.19041.0\um\CommCtrl.h）：
//   MultiSelect      0x0002 MCS_MULTISELECT   （样式位；副作用可被控件侧读数问到：
//                                                没这两位时 MCM_SETMAXSELCOUNT 直接失败）
//   ShowWeekNumbers  0x0004 MCS_WEEKNUMBERS   （样式位；控件侧：MCM_GETMINREQRECT 变宽）
//   ShowToday        0x0010 MCS_NOTODAY 取反  （样式位；控件侧：MCM_GETMINREQRECT 变高）
//   MaxSelCount      MCM_SETMAXSELCOUNT / MCM_GETMAXSELCOUNT（**真往返过控件**，不是自存）
//   MonthRows / MonthColumns
//                    原生没有"行列"这一条消息 —— 多月平铺完全由**矩形多大**决定（头 6361 行
//                    那段注释就是这么教的）。所以 Init 按 rows×cols 摆矩形，再让
//                    MCM_SIZERECTTOMIN 把它撑到真装得下，最后用 MCM_GETCALENDARCOUNT 问控件
//                    "你眼下画了几个月" —— 这一格第一次有了控件侧读数。
//   BackColor / ForeColor / TitleBackColor / TitleForeColor / TrailingForeColor
//                    MCM_SETCOLOR / MCM_GETCOLOR 的 MCSC_* 六格（0/1/2/3/4/5，六个名字全是
//                    VB6 真名：MonthBackColor 在官方那页的配色段里就有）
//
// "窗口真建起来了"怎么证：不另开读数口子 —— 样式类属性走"写进 GWL_STYLE 再读回来"，
//   句柄为 NULL 时那条写是空转、读必回默认，写读不一致就翻红；类名字面串由 emitc 形状针钉住
//   （vb6_CreateControl("SysMonthCal32"…)），两头夹住中间那段"建错类还能自洽"的空子。
//
// 为什么样式判据一律配一条控件侧读数（与 DT-a 同一条纪律）：GWL_STYLE 的往返只证明
//   "我们写进去了"，不证明"控件按那位在画"。这里更危险一层 —— MCS_ 那三位里
//   MCS_MULTISELECT(0x2) 与 MCS_WEEKNUMBERS(0x4) 挨着，只对自己的掩码读数会自洽地假绿。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include <string.h>

#ifdef _WIN32

// commctrl.h 在低 _WIN32_IE 下不给这些，且 cgen 那侧按字面量发 —— 两边都钉死数值。
#ifndef MCS_MULTISELECT
#define MCS_MULTISELECT          0x0002L
#endif
#ifndef MCS_WEEKNUMBERS
#define MCS_WEEKNUMBERS          0x0004L
#endif
#ifndef MCS_NOTODAY
#define MCS_NOTODAY              0x0010L
#endif
#ifndef MCM_GETMAXSELCOUNT
#define MCM_GETMAXSELCOUNT       (0x1000 + 3)
#endif
#ifndef MCM_SETMAXSELCOUNT
#define MCM_SETMAXSELCOUNT       (0x1000 + 4)
#endif
#ifndef MCM_GETMINREQRECT
#define MCM_GETMINREQRECT        (0x1000 + 9)
#endif
#ifndef MCM_SETCOLOR
#define MCM_SETCOLOR             (0x1000 + 10)
#endif
#ifndef MCM_GETCOLOR
#define MCM_GETCOLOR             (0x1000 + 11)
#endif
#ifndef MCM_GETCALENDARCOUNT
#define MCM_GETCALENDARCOUNT     (0x1000 + 23)
#endif
#ifndef MCM_SIZERECTTOMIN
#define MCM_SIZERECTTOMIN        (0x1000 + 29)
#endif
// MCSC_* 五格在有的 SDK 版本里根本没定义（头里只有 #define 的注释文字），全部按数值给。
#define VB6_MCSC_BACKGROUND   0   /* VB6 BackColor（月与月之间的那层） */
#define VB6_MCSC_TEXT         1   /* VB6 ForeColor（日期字） */
#define VB6_MCSC_TITLEBK      2   /* VB6 TitleBackColor */
#define VB6_MCSC_TITLETEXT    3   /* VB6 TitleForeColor */
#define VB6_MCSC_MONTHBK      4   /* 原生多出来的一格：月历内部底色 */
#define VB6_MCSC_TRAILINGTEXT 5   /* VB6 TrailingForeColor（非本月字） */

static DWORD vb6_MvStyle(HWND hwnd) {
    return (DWORD)(DWORD_PTR)GetWindowLongPtrW(hwnd, GWL_STYLE);
}

static void vb6_MvSetStyleBits(HWND hwnd, DWORD on, DWORD off) {
    DWORD s = vb6_MvStyle(hwnd);
    SetWindowLongPtrW(hwnd, GWL_STYLE, (LONG_PTR)((s & ~off) | on));
}

// ---------------- 样式三格 ----------------
// 读的是窗口样式位本身（与观感同源）；写是"尽力而为"—— 这三位在原生里是**创建参数**，
// 运行期写下去控件抹不抹由夹具的读数说，不由这里宣称。
int32_t vb6_MV_GetMultiSelect(void* hwnd) {
    if (!hwnd) return 0;
    return (vb6_MvStyle((HWND)hwnd) & MCS_MULTISELECT) ? -1 : 0;
}
void vb6_MV_SetMultiSelect(void* hwnd, int32_t on) {
    if (!hwnd) return;
    vb6_MvSetStyleBits((HWND)hwnd, (DWORD)(on ? MCS_MULTISELECT : 0),
                       (DWORD)(on ? 0 : MCS_MULTISELECT));
}
int32_t vb6_MV_GetShowWeekNumbers(void* hwnd) {
    if (!hwnd) return 0;
    return (vb6_MvStyle((HWND)hwnd) & MCS_WEEKNUMBERS) ? -1 : 0;
}
void vb6_MV_SetShowWeekNumbers(void* hwnd, int32_t on) {
    if (!hwnd) return;
    vb6_MvSetStyleBits((HWND)hwnd, (DWORD)(on ? MCS_WEEKNUMBERS : 0),
                       (DWORD)(on ? 0 : MCS_WEEKNUMBERS));
}
// VB6 的 ShowToday 与原生 MCS_NOTODAY 是**反**的：ShowToday=False ⇒ 挂上 NOTODAY 那位。
int32_t vb6_MV_GetShowToday(void* hwnd) {
    if (!hwnd) return 0;
    return (vb6_MvStyle((HWND)hwnd) & MCS_NOTODAY) ? 0 : -1;
}
void vb6_MV_SetShowToday(void* hwnd, int32_t on) {
    if (!hwnd) return;
    vb6_MvSetStyleBits((HWND)hwnd, (DWORD)(on ? 0 : MCS_NOTODAY),
                       (DWORD)(on ? MCS_NOTODAY : 0));
}

// ---------------- 范围选择的上限（真往返过控件） ----------------
int32_t vb6_MV_GetMaxSelCount(void* hwnd) {
    if (!hwnd) return 0;
    return (int32_t)(INT_PTR)SendMessageW((HWND)hwnd, MCM_GETMAXSELCOUNT, 0, 0);
}
void vb6_MV_SetMaxSelCount(void* hwnd, int32_t n) {
    if (!hwnd || n < 1) return;
    SendMessageW((HWND)hwnd, MCM_SETMAXSELCOUNT, (WPARAM)n, 0);
}

// ---------------- 控件侧读数 ----------------
// MCM_GETMINREQRECT 问的是"装得下**一个**月"的最小矩形（头 6361 行原文）：装 week numbers
// 会加宽、要不要 today 那一条会改高 —— 于是样式位第一次有了不经过 GWL_STYLE 的读数。
static void vb6_MvMinReq(HWND hwnd, int32_t* w, int32_t* h) {
    RECT rc;
    if (w) *w = 0;
    if (h) *h = 0;
    if (!hwnd) return;
    memset(&rc, 0, sizeof(rc));
    if (!SendMessageW((HWND)hwnd, MCM_GETMINREQRECT, 0, (LPARAM)&rc)) return;
    if (w) *w = (int32_t)(rc.right - rc.left);
    if (h) *h = (int32_t)(rc.bottom - rc.top);
}
int32_t vb6_MV_MinReqWidth(void* hwnd)  { int32_t w, h; vb6_MvMinReq(hwnd, &w, &h); return w; }
int32_t vb6_MV_MinReqHeight(void* hwnd) { int32_t w, h; vb6_MvMinReq(hwnd, &w, &h); return h; }

// "眼下真画了几个月" —— 控件自己的答案，不是我们那张表。
int32_t vb6_MV_GetMonthCount(void* hwnd) {
    if (!hwnd) return 0;
    return (int32_t)(INT_PTR)SendMessageW((HWND)hwnd, MCM_GETCALENDARCOUNT, 0, 0);
}

// ---------------- 设计期初值 ----------------
// rows/cols 两条 <=0 表示 .frm 没写 ⇒ 一动不动（多月平铺是"给多大矩形"这件事，
// 1×1 时连问都不必问）。撑矩形按头里教的那条：先按行列把矩形乘出来，再让
// MCM_SIZERECTTOMIN 把它撑到真装得下 rows*cols 个月，最后 MoveWindow 落回去。
// maxSelCount 传 -999 = .frm 没写这一格 ⇒ 不发 MCM_SETMAXSELCOUNT（原生默认是 7，
// 发不发是可被 GetMaxSelCount 问出来的差别，判据就靠这条区分"设计期真到了控件"）。
// 客户区尺寸由 RTL 算，边框/标题不参与（子窗口没有非客户区）。
void vb6_MV_Init(void* hwnd, int32_t rows, int32_t cols, int32_t maxSelCount) {
    RECT rc, base;
    int32_t w, h;
    if (!hwnd) return;
    if (maxSelCount > 0) vb6_MV_SetMaxSelCount(hwnd, maxSelCount);
    if (rows < 1) rows = 1;
    if (cols < 1) cols = 1;
    if (rows == 1 && cols == 1) return;
    memset(&base, 0, sizeof(base));
    if (!SendMessageW((HWND)hwnd, MCM_GETMINREQRECT, 0, (LPARAM)&base)) return;
    w = (int32_t)(base.right - base.left);
    h = (int32_t)(base.bottom - base.top);
    if (w <= 0 || h <= 0) return;
    SetRect(&rc, 0, 0, w * cols, h * rows);
    SendMessageW((HWND)hwnd, MCM_SIZERECTTOMIN, (WPARAM)(rows * cols), (LPARAM)&rc);
    GetWindowRect((HWND)hwnd, &base);          /* 只换尺寸，位置原样 */
    MoveWindow((HWND)hwnd, base.left, base.top,
               (int)(rc.right - rc.left), (int)(rc.bottom - rc.top), TRUE);
}

// ---------------- 底色 / 字色 ----------------
static int32_t vb6_MvGetColor(void* hwnd, WPARAM which) {
    if (!hwnd) return 0;
    return (int32_t)(INT_PTR)SendMessageW((HWND)hwnd, MCM_GETCOLOR, which, 0);
}
static void vb6_MvSetColor(void* hwnd, WPARAM which, int32_t val) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, MCM_SETCOLOR, which, (LPARAM)val);
}
int32_t vb6_MV_GetBackColor(void* hwnd)        { return vb6_MvGetColor(hwnd, VB6_MCSC_BACKGROUND); }
void    vb6_MV_SetBackColor(void* hwnd, int32_t v)  { vb6_MvSetColor(hwnd, VB6_MCSC_BACKGROUND, v); }
int32_t vb6_MV_GetForeColor(void* hwnd)        { return vb6_MvGetColor(hwnd, VB6_MCSC_TEXT); }
void    vb6_MV_SetForeColor(void* hwnd, int32_t v)  { vb6_MvSetColor(hwnd, VB6_MCSC_TEXT, v); }
int32_t vb6_MV_GetTitleBackColor(void* hwnd)   { return vb6_MvGetColor(hwnd, VB6_MCSC_TITLEBK); }
void    vb6_MV_SetTitleBackColor(void* hwnd, int32_t v)  { vb6_MvSetColor(hwnd, VB6_MCSC_TITLEBK, v); }
int32_t vb6_MV_GetTitleForeColor(void* hwnd)   { return vb6_MvGetColor(hwnd, VB6_MCSC_TITLETEXT); }
void    vb6_MV_SetTitleForeColor(void* hwnd, int32_t v)  { vb6_MvSetColor(hwnd, VB6_MCSC_TITLETEXT, v); }
int32_t vb6_MV_GetTrailingForeColor(void* hwnd){ return vb6_MvGetColor(hwnd, VB6_MCSC_TRAILINGTEXT); }
void    vb6_MV_SetTrailingForeColor(void* hwnd, int32_t v) { vb6_MvSetColor(hwnd, VB6_MCSC_TRAILINGTEXT, v); }
// VB6 的官方属性面里就有 MonthBackColor（《MonthView 控件》那页的配色段：MonthBackColor /
// TitleBackColor / TitleForeColor / TrailingForeColor）⇒ 这一格不是 C3 扩展，是第六条真名。
int32_t vb6_MV_GetMonthBackColor(void* hwnd)   { return vb6_MvGetColor(hwnd, VB6_MCSC_MONTHBK); }
void    vb6_MV_SetMonthBackColor(void* hwnd, int32_t v)  { vb6_MvSetColor(hwnd, VB6_MCSC_MONTHBK, v); }

// ---------------- Value / SelStart / SelEnd（C29-MV-b）----------------
// Date <-> SYSTEMTIME 走 vb6forms_internal.h 那组共用助手（与 DTPicker 同一对换算）。
// 原生两条消息正好对上 VB6 的两格：MCM_GET/SETCURSEL = Value（单选/焦点那一格），
// MCM_GET/SETSELRANGE = SelStart/SelEnd（**一张两端表** ⇒ 改一端必须像 DTPicker 的范围端点
// 那样先读回整张表、换掉那一格、两格一起发回去，只发一端会把另一端拆掉）。
#ifndef MCM_GETCURSEL
#define MCM_GETCURSEL      (0x1000 + 1)
#endif
#ifndef MCM_SETCURSEL
#define MCM_SETCURSEL      (0x1000 + 2)
#endif
#ifndef MCM_GETSELRANGE
#define MCM_GETSELRANGE    (0x1000 + 5)
#endif
#ifndef MCM_SETSELRANGE
#define MCM_SETSELRANGE    (0x1000 + 6)
#endif

// VB6 的 Value / SelStart / SelEnd 都是**纯日期**（月历没有"时分"这一格），而原生 comctl
// **6** 在 `MCM_GETCURSEL` 回来的 SYSTEMTIME 里把**当前挂钟时间**填进四个时间字段 ——
// 本机实测（`.build\mcsel6b_out.txt`，同一份源码编两份、只改挂不挂 v6 manifest）：
// 发 43894 回来 `2020-03-04 18:21:47.128` = 序列号 43894.765127，而不挂 manifest 的那份
// 回 `2020-03-04 00:00:00.000` = 43894.000000。 ⇒ 不抹平就是**下午红、上午绿**：
// `CLng(mv.Value)` 在 0.765 那一头进位成 43895，而 CI 那批跑在上午（分数 < 0.5）照旧绿，
// 所以这道门一直拦不住它。日期面一律归到"那一天"。
static double vb6_MvDaySerial(const SYSTEMTIME* st) {
    SYSTEMTIME d;
    d = *st;
    d.wHour = 0; d.wMinute = 0; d.wSecond = 0; d.wMilliseconds = 0;
    return vb6_DateToSerial(&d);
}

double vb6_MV_GetValue(void* hwnd) {
    SYSTEMTIME st;
    if (!hwnd) return 0.0;
    vb6_DateZero(&st);
    if (!SendMessageW((HWND)hwnd, MCM_GETCURSEL, 0, (LPARAM)&st)) return 0.0;
    return vb6_MvDaySerial(&st);
}

void vb6_MV_SetValue(void* hwnd, double serial) {
    SYSTEMTIME st;
    if (!hwnd) return;
    if (!vb6_DateFromSerial(serial, &st)) return;
    SendMessageW((HWND)hwnd, MCM_SETCURSEL, 0, (LPARAM)&st);
}

// 整张范围表读回来。问不出来（没挂 MCS_MULTISELECT 的控件答不答，由判据读数说）
// 就把两格留 0 并回 ok=0，让调用方区分"问不出"与"答案是 1900 年那天"。
static void vb6_MvGetSelRange(void* hwnd, SYSTEMTIME* rg, int32_t* ok) {
    if (ok) *ok = 0;
    vb6_DateZero(&rg[0]);
    vb6_DateZero(&rg[1]);
    if (!hwnd) return;
    if (SendMessageW((HWND)hwnd, MCM_GETSELRANGE, 0, (LPARAM)rg)) *ok = 1;
}

double vb6_MV_GetSelStart(void* hwnd) {
    SYSTEMTIME rg[2];
    int32_t ok;
    vb6_MvGetSelRange(hwnd, rg, &ok);
    return ok ? vb6_MvDaySerial(&rg[0]) : 0.0;
}

double vb6_MV_GetSelEnd(void* hwnd) {
    SYSTEMTIME rg[2];
    int32_t ok;
    vb6_MvGetSelRange(hwnd, rg, &ok);
    if (!ok) return 0.0;
    /* 止端这一格实测过两版 comctl（同一份探针挂/不挂 v6 manifest 各跑一遍，读数
       `.build\mcsel6b_out.txt`，写进 lo=46268 / hi=46272）：**v6 回 `2026-09-07 23:59:59.9`**
       （序列号 46272.999988），**v5 回 `2026-09-07 00:00`**（46272.000000）⇒ 两版存的都正好是
       **被选中的最后一天**，差别只在那截时间。台账与这文件旧注释里那句"控件内部存成
       [起, 止+1) 的半开区间、止端恒比写入值多一天"是**误读** —— 那是把 23:59 那截经 `CLng`
       进位之后看到的 46273 当成了控件里的数。所以这里不折天，只**抹时间**（同 `Value`
       那一条助手）：旧写法 `-1.0` 在 v6 上得到 46270.999988 这种数，只有靠 `CLng` 四舍五入
       才碰巧读对，`If mv.SelEnd = d` 按数值比恒差一天（MV40 就是钉这条）。 */
    return vb6_MvDaySerial(&rg[1]);
}

// 改一端 = 读回整张表 → 换掉那一格 → 归一化 → 两格一起发回去。
//
// 归一化口径 (2026-10-01, ctrlmonthview MV26/MV33/MV40 三条一起定):
//
//   原生 MCM_SETSELRANGE 是"全对或全否"的接口: 要求 rg[0] <= rg[1] 且
//   宽度 <= MaxSelCount, 否则整条 FALSE 拒收。VB 侧一次只写一端, 拿原生当"改一格"
//   用是**接不上**的:
//     · 控件刚建起来时初值是一张同日 [Today, Today]; 用户先写 `SelStart = Today+1`
//       就会算出 [Today+1, Today] 反序 → 原生吞掉 → 下一次 GET 还是 Today, 上层
//       看不见这次写入 (MV26 跨月那天翻红的根因)。
//     · 已经贴着 MaxSelCount 的窗口被往外推 → 原生同样吞掉 → 用户按代码字面看,
//       "另一端明明没让我改, 怎么读的旧值还在" (MV33 之前把这条钉成"两端都不动",
//       实际上是把原生的沉默当成了语义)。
//   VB6 MSComCtl2.MonthView 的文档在这两处都规定了"另一端跟着调":
//     SelStart > SelEnd ⇒ SelEnd := SelStart; SelEnd < SelStart ⇒ SelStart := SelEnd;
//     宽度超过 MaxSelCount ⇒ 另一端被挤到 SelStart+Max-1 或 SelEnd-Max+1。
//   这版 RTL 把这两条归一化都在**发**之前做完, 让上层看到的 SELRANGE 属性接口
//   符合它字面承诺 ("我写的这一格一定读得出"), 不再靠原生沉默拒收的巧合。
//   MV33/MV40 各自改到自己对应的落点。
static void vb6_MvSetSelEnd(void* hwnd, int which, double serial) {
    SYSTEMTIME rg[2], next;
    int32_t ok;
    if (!hwnd) return;
    if (!vb6_DateFromSerial(serial, &next)) return;
    vb6_MvGetSelRange(hwnd, rg, &ok);
    rg[which] = next;
    if (ok) {
        /* (a) 顺序: 哪一格被写了, 另一端翻过它就被拉到同一格。 */
        if (vb6_MvDaySerial(&rg[0]) > vb6_MvDaySerial(&rg[1])) {
            rg[1 - which] = next;
        }
        /* (b) 宽度: 贴着 MaxSelCount 时把另一端挤回合法窗口。 */
        LRESULT maxc = SendMessageW((HWND)hwnd, MCM_GETMAXSELCOUNT, 0, 0);
        if (maxc > 0) {
            double lo = vb6_MvDaySerial(&rg[0]);
            double hi = vb6_MvDaySerial(&rg[1]);
            double width = hi - lo + 1.0;
            if (width > (double)maxc) {
                if (which == 0) {
                    /* SelStart 定住, SelEnd 收到 Start + max - 1 那天 */
                    double target = lo + (double)maxc - 1.0;
                    if (!vb6_DateFromSerial(target, &rg[1])) { /* 转换失败就退回原 rg[1] */ }
                } else {
                    /* SelEnd 定住, SelStart 收到 End - max + 1 那天 */
                    double target = hi - (double)maxc + 1.0;
                    if (!vb6_DateFromSerial(target, &rg[0])) { /* 同上 */ }
                }
            }
        }
    }
    SendMessageW((HWND)hwnd, MCM_SETSELRANGE, 0, (LPARAM)rg);
}

void vb6_MV_SetSelStart(void* hwnd, double serial) { vb6_MvSetSelEnd(hwnd, 0, serial); }
void vb6_MV_SetSelEnd(void* hwnd, double serial)   { vb6_MvSetSelEnd(hwnd, 1, serial); }

// ---------------- 通知换算与判据助手 (C29-MV-c) ----------------
// VB6 的 `DateClick(ByVal DateSelected As Date)` 对应原生一条 MCN_SELCHANGE(-749)，
// 负载 = NMSELCHANGE{nmhdr, stSelStart, stSelEnd}（头 6577 行）。派发那头只要**那一天**，
// 所以这里把负载折成一个 Date（double 序列号）；问不出来回 0（同 DT-b 的口径：
// 0 = "没有值"，绝不回负数冒充一个怪日期）。
#ifndef MCN_SELCHANGE
#define MCN_SELCHANGE      (-749L)
#endif

double vb6_MV_NotifyDate(void* nmSelChange) {
    NMSELCHANGE* sc = (NMSELCHANGE*)nmSelChange;
    if (!sc) return 0.0;
    return vb6_MvDaySerial(&sc->stSelStart);
}

// 判据专用（不对应任何 VB6 语义，见 029 §九 本格）：无头环境点不了鼠标，而直接调 handler
// 会绕开整条派发链 —— 只有从真 WM_NOTIFY 进父窗，才验得到"case WM_NOTIFY + code 分流 +
// hwndFrom 认来源 + 负载折算"四段都接上了。手法照 C29-DT-c 的 vb6_DTP_SimChange。
// 负载按**调用方给的那一天**填，两端填同一个值（= VB6 单点选中时 SelStart == SelEnd 那一态）。
void vb6_MV_SimDateClick(void* hwnd, double serial) {
    NMSELCHANGE sc;
    SYSTEMTIME st;
    HWND parent;
    if (!hwnd) return;
    if (!vb6_DateFromSerial(serial, &st)) return;
    memset(&sc, 0, sizeof(sc));
    sc.nmhdr.hwndFrom = (HWND)hwnd;
    sc.nmhdr.idFrom   = (UINT_PTR)GetWindowLongPtrW((HWND)hwnd, GWLP_ID);
    sc.nmhdr.code     = (DWORD)MCN_SELCHANGE;
    sc.stSelStart = st;
    sc.stSelEnd   = st;
    parent = GetParent((HWND)hwnd);
    if (!parent) parent = (HWND)hwnd;
    SendMessageW(parent, WM_NOTIFY, 0, (LPARAM)&sc);
}

#endif /* _WIN32 */
