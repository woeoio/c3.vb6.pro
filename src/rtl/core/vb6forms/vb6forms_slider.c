// vb6forms_slider.c — VB6 Slider 控件 (ai/内置控件/Slider 控件（滑杆）.md, ai/029 C29-SL)
//
// 复刻口径: 用 comctl32 的 msctls_trackbar32 复刻 VB6 Slider 的等价行为,
// **不加载任何 OCX** (mscomctl.ocx 在本机未注册, 且 32 位 inproc 无法进 x64 进程)。
//
// 本文件按格长厚：**SL-a** = 窗口 + 创建样式 + 标量属性面，**SL-b** = 值面
// (Min/Max/Value/Small·LargeChange/Sel*)，**SL-c** = Change / Scroll 两条事件（文末那一节），
// **SL-h** = TickStyle 四档 + GetNumTicks，**SL-i** = SelLength / ClearSel（类型库读出来的
// VB6 那一面，见文末两段）。
//
// 实测口径 (.build/slprobe/slmeasure*.c, x64 真跑; 全文抄在 ai/029 §九 C29-SL-0 那一格):
//   1) **方向的正解**: TBM_GETCHANNELRECT 的 rect 永远把行程长度放在 **x 分量** ——
//      300x40 的水平杆与 40x300 的竖杆都答 (8,10)-(292,14)，它压根不随方向旋转，
//      所以"量 channel 的长短边"等于什么都没量（本格第一发就这么错过一次，见下第 4 条）。
//      能用的证人是 **TBM_GETTHUMBRECT 在 min/max 两点之间的位移轴** (TravelIsVert)。
//   2) Orientation **运行期改是有效的**: 创建水平之后写 TBS_VERT + SetWindowPos(SWP_FRAMECHANGED)，
//      滑块矩形与"创建时就是垂直"逐字相同（(2,8)-(24,19)）；对照组同样 SetWindowPos 但不翻样式，
//      读数不变 ⇒ 翻样式这一步才是因。**第一发探针量出"只能创建时定"是错的**，错因就是第 1 条
//      那条假证人 —— 样式位读回来一直是"写了"，而布局到底听不听，得推滑块才知道。
//   3) TickFrequency 原生**问不出**: 没有 TBM_GETTICFREQ；TBM_GETTIC(pos) 在 freq=10 下对
//      pos=15 也返回非 0（返回值 = pos+1，答的是"在不在量程内"）；TBM_GETTICPOS(i) 在 freq=10
//      与 freq=25 下读数逐字相同（16,18,21,24,27,29），只有"有没有刻度"这一维能证
//      （TBM_CLEARTICS 之后全变 -1）。⇒ 频率只能自存读回 + 真下发一次 TBM_SETTICFREQ；
//      "有刻度"这条用 TickPresent 当控件侧证人。
//   4) Value 越界 = 原生钳位（range 10..100 时 SETPOS(150) 读回 100、(-5) 读回 10）—— SL-b 用。
//      TBM_GETCHANNELRECT / GETTHUMBRECT 的**返回值不是成功标志**（实测返回 0 而 rect 有效），
//      一律只看 rect 内容。
//
// VB6 枚举（本文件用到的这一档）：
//   Orientation: 0 = sldHorizontal（默认），1 = sldVertical
//   TickStyle:   0 = sldBottomRight, 1 = sldTopLeft, 2 = sldBoth, 3 = sldNoTicks
//     （SL-h 起这不再是我猜的：四档与它的数值是从 **MSCOMCTL.OCX 自带的那张类型库**读出来的，
//      探针 `.build/slprobe/sltlb.cpp` 走 `LoadTypeLibEx(路径, REGKIND_NONE)` —— 不需要注册，
//      OCX 在 System32 里就有。对应关系与实测证人见下面 C29-SL-h 那一格。）

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <commctrl.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"

#ifdef _WIN32

#define VB6_SLD_HORZ 0
#define VB6_SLD_VERT 1

// 刻意**不写** `#ifndef TBS_xxx / #define ...` 那一族兜底宏：上一版手写的
// `#define TBM_GETTHUMBRECT (WM_USER + 17)` 与 SDK 头对不上（头里 GETTHUMBRECT 是 +25，
// +17 是 TBM_GETSELSTART）—— 只因 commctrl.h 先定义了它，#ifndef 才没盖上去。
// 本文件全程 #include <commctrl.h>，消息名一律用头里的，不再抄第二份编号表。
#ifndef VB6_SLD_UNSET
#define VB6_SLD_UNSET (-999)     // 设计期"没写过"那一档的哨兵（同 vb6_TreeView_Init）
#endif
#ifndef VB6_SLD_I16MIN
#define VB6_SLD_I16MIN (-32768)
#endif
#ifndef VB6_SLD_I16MAX
#define VB6_SLD_I16MAX 32767
#endif

// --- 自存槽（+1 偏移，同 vb6forms_progress.c 那条：SetPropW(hw,name,NULL) 看着像"存了个空值"，
//     而 GetPropW 对"未设置"与"存了 0"都返回 NULL —— 不偏移就会把 `TickFrequency = 0`
//     这种合法值当成未设置，静默答回默认）---
static const wchar_t kSdTickFreq[] = L"VB6_SD_TickFrequency";
// C29-SL-c：Change 的基准值（上一次派发时控件是多少）。与 TickFrequency 同一档自存理由：
// 原生没有任何"上次值"可问，而 GetPropW 对未设置与存了 0 都答 NULL ⇒ 走 +1 偏移那套。
static const wchar_t kSdLastPos[] = L"VB6_SD_LastPos";

static void vb6_SdSetProp(void* hwnd, const wchar_t* name, LONG val) {
    if (!hwnd) return;
    SetPropW((HWND)hwnd, name, (HANDLE)(INT_PTR)((LONG)val + 1));
}

static LONG vb6_SdGetProp(void* hwnd, const wchar_t* name, LONG def) {
    if (!hwnd) return def;
    HANDLE h = GetPropW((HWND)hwnd, name);
    return h ? ((LONG)(INT_PTR)h - 1) : def;
}

// 换一次帧：实测样式位（Orientation / EnableSelRange）写完要配这一脚，控件才真按新位重排
// （SetWindowPos 传原尺寸，只挂 NOZORDER|NOMOVE|FRAMECHANGED）。
static void vb6_SdReframe(HWND hw) {
    RECT wr;
    GetWindowRect(hw, &wr);
    SetWindowPos(hw, NULL, 0, 0, (int)(wr.right - wr.left), (int)(wr.bottom - wr.top),
                 SWP_NOZORDER | SWP_NOMOVE | SWP_FRAMECHANGED);
}

// 设计期下发（SL-a 只有 tickFrequency 一参，SL-b 把它扩成完整值面）。
// 哨兵 -999 = .frm 里**没写过**这一项 ⇒ 那一条消息都不发，保持控件默认
// （同 vb6_TreeView_Init：拿 -1 当"未写"会让设计期永远设不上合法的 0/-1 那一档）。
// 顺序是量出来的，不能换：
//   1) range 先立 —— 实测 `TBM_SETRANGE(10,100)` 会把 `pos` 顶到下限 10、把 `page` 从 20
//      重算成 18、把 `selstart` 跟到 10；先设值再设 range 就会被这三处连带改动吃掉；
//   2) line/page 再设 —— 它们同样会被 range 重算，必须在 range 之后；
//   3) pos 第三（量程已定，越界由控件钳位 —— 实测 SETPOS(150)→100、(-5)→10）；
//   4) 刻度与 Sel 最后。
// Orientation 不走这里 —— 它是创建参数，cgen 立进 vb6_CreateControl 的样式位。
void vb6_Slider_Init(void* hwnd, long min, long max, long value,
                     long smallChange, long largeChange, long tickFrequency,
                     long selStart, long selEnd, long selectRange) {
    HWND hw;
    long lo, hi, curLo, curHi;
    if (!hwnd) return;
    hw = (HWND)hwnd;
    // 只有 .frm **真写过** Min/Max 才发 range：控件当前值要先读出来当另一端。
    // （踩过的坑：早先写成"没写就从当前值取"再比 `lo != min` 判要不要发 —— 那比较拿
    //  当前值与哨兵 -999 比，永远不等 ⇒ 每枚 Slider 都被重发一次 range，而重发 range
    //  会把 page 从 20 重算成 18、把刻度画出来 ⇒ 什么都没写的控件也"自己长出刻度"。）
    curLo = (long)SendMessageW(hw, TBM_GETRANGEMIN, 0, 0);
    curHi = (long)SendMessageW(hw, TBM_GETRANGEMAX, 0, 0);
    if (min != VB6_SLD_UNSET || max != VB6_SLD_UNSET) {
        lo = (min == VB6_SLD_UNSET) ? curLo : min;
        hi = (max == VB6_SLD_UNSET) ? curHi : max;
        // VB6 的 Min/Max 是 Long，原生这条消息的 lParam 是两个 16 位半字（实测
        // MAKELONG(40000,50000) 读回 -25536/-15536 = 截断）⇒ 先钳到 i16 再下发，
        // 这样"我们答出去的数"与"控件真走得动的数"是同一个。超界那一档 VB6 怎么办：
        // 本机 OCX 未注册、拿不到真值 ⇒ 押后（见 029）。
        if (lo < VB6_SLD_I16MIN) lo = VB6_SLD_I16MIN;
        if (lo > VB6_SLD_I16MAX) lo = VB6_SLD_I16MAX;
        if (hi < VB6_SLD_I16MIN) hi = VB6_SLD_I16MIN;
        if (hi > VB6_SLD_I16MAX) hi = VB6_SLD_I16MAX;
        if (hi < lo) hi = lo;
        if (lo != curLo || hi != curHi) {
            SendMessageW(hw, TBM_SETRANGE, TRUE, MAKELONG((WORD)(SHORT)lo, (WORD)(SHORT)hi));
        }
    }
    if (smallChange != VB6_SLD_UNSET && smallChange >= 0) {
        SendMessageW(hw, TBM_SETLINESIZE, TRUE, (WPARAM)smallChange);
    }
    if (largeChange != VB6_SLD_UNSET && largeChange >= 0) {
        SendMessageW(hw, TBM_SETPAGESIZE, TRUE, (WPARAM)largeChange);
    }
    if (value != VB6_SLD_UNSET) {
        SendMessageW(hw, TBM_SETPOS, TRUE, (WPARAM)value);
    }
    if (tickFrequency != VB6_SLD_UNSET) {
        vb6_SdSetProp(hwnd, kSdTickFreq, (LONG)tickFrequency);
        if (tickFrequency > 0) {
            // 先清再设：控件自带的默认刻度不清掉的话，新旧两套刻度会叠在一起画。
            SendMessageW(hw, TBM_CLEARTICS, TRUE, 0);
            SendMessageW(hw, TBM_SETTICFREQ, (WPARAM)tickFrequency, 0);
        }
    }
    if (selectRange != VB6_SLD_UNSET && selectRange != 0) {
        LONG st = (LONG)GetWindowLongPtrW(hw, GWL_STYLE);
        if (!(st & TBS_ENABLESELRANGE)) {
            SetWindowLongPtrW(hw, GWL_STYLE, st | TBS_ENABLESELRANGE);
            vb6_SdReframe(hw);
        }
    }
    if (selStart != VB6_SLD_UNSET || selEnd != VB6_SLD_UNSET) {
        long a = (selStart == VB6_SLD_UNSET)
                     ? vb6_Slider_GetSelStart(hwnd) : selStart;
        long b = (selEnd == VB6_SLD_UNSET)
                     ? vb6_Slider_GetSelEnd(hwnd) : selEnd;
        if (a > b) a = b;                     // 原生不接受反向区段
        SendMessageW(hw, TBM_SETSEL, TRUE, MAKELONG((WORD)(SHORT)a, (WORD)(SHORT)b));
    }
    // C29-SL-c：把设计期落定之后的那一个值记成 Change 的基准（后面每次派发都跟它比）。
    vb6_SdSetProp(hwnd, kSdLastPos, (LONG)(int)SendMessageW(hw, TBM_GETPOS, 0, 0));
}

// VB6 属性面：读样式位那一档（运行期写也走样式位，实测有效 —— 见文件头第 2 条）。
int vb6_Slider_GetOrientation(void* hwnd) {
    if (!hwnd) return VB6_SLD_HORZ;
    return ((GetWindowLongPtrW((HWND)hwnd, GWL_STYLE) & TBS_VERT) != 0)
               ? VB6_SLD_VERT : VB6_SLD_HORZ;
}

void vb6_Slider_SetOrientation(void* hwnd, int orientation) {
    LONG st;
    if (!hwnd) return;
    st = (LONG)GetWindowLongPtrW((HWND)hwnd, GWL_STYLE);
    if (orientation != 0) st |= TBS_VERT;
    else                  st &= ~TBS_VERT;
    SetWindowLongPtrW((HWND)hwnd, GWL_STYLE, st);
    // 光写样式不换帧不重排（实测对照组：只做 SetWindowPos 不翻样式 ⇒ 方向不变；
    // 翻了样式再做同一次 SetWindowPos ⇒ 方向跟着变）。尺寸传原值，只挂 FRAMECHANGED。
    {
        RECT wr;
        GetWindowRect((HWND)hwnd, &wr);
        SetWindowPos((HWND)hwnd, NULL, 0, 0,
                     (int)(wr.right - wr.left), (int)(wr.bottom - wr.top),
                     SWP_NOZORDER | SWP_NOMOVE | SWP_FRAMECHANGED);
    }
}

// C3 扩展（不是 VB6 属性，判据专用）：这杆到底**横着走还是竖着走**。
//
// 三条候选证人量过（`.build/slprobe/slmeasure3/4/5.c`），只有一条站得住：
//   * TBM_GETCHANNELRECT 的"长短边" **废** —— 它的 rect 永远把行程长度放在 x 分量：
//     300x40 的横杆与 40x300 的竖杆都答 (8,10)-(292,14)。第一发探针拿它判方向，
//     于是得出"Orientation 运行期改不动"的**错结论**（已订正，见文件头第 2 条）。
//   * "把滑块推到量程两端看位移轴" 也 **不单独够用** —— 133x27 的窗口里挂 TBS_VERT 时
//     竖直行程被压成 0（实测 Δx=Δy=0），Δ 退化成一潭死水，答"横"。
//   * **滑块自己的形状** 才是稳的：横杆的滑块是"竖块" (11x22)、竖杆的是"横块" (22x11)，
//     因为那张滑块图是随控件方向一起转的。实测四种情形（创建横 / 创建竖 / 运行期翻竖 /
//     窄高竖）里 w>h ⟺ 竖直 **四条全对**，而 Δ 判在"太矮的竖杆"上失手。
// 所以这里用形状判；位移量留在注释里当反例（别再有人拿 channel 矩形判方向）。
int vb6_Slider_TravelIsVert(void* hwnd) {
    RECT th;
    if (!hwnd) return 0;
    th.left = th.top = th.right = th.bottom = 0;
    SendMessageW((HWND)hwnd, TBM_GETTHUMBRECT, 0, (LPARAM)&th);
    return ((th.right - th.left) > (th.bottom - th.top)) ? 1 : 0;
}

// C3 扩展（证人）：控件当前**有没有**画刻度。实测这条只能证"有/无"，证不了频率
// （freq=10 与 freq=25 下 GETTICPOS(0..5) 读数逐字相同；清完刻度才全变 -1）。
int vb6_Slider_TickPresent(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, TBM_GETTICPOS, 0, 0) != -1 ? -1 : 0;  // VB6: True=-1
}

// `TickFrequency`：原生没有回读消息 ⇒ 自存读回；写侧真下发 TBM_SETTICFREQ。
int vb6_Slider_GetTickFrequency(void* hwnd) {
    return (int)vb6_SdGetProp(hwnd, kSdTickFreq, 0);
}

void vb6_Slider_SetTickFrequency(void* hwnd, int freq) {
    if (!hwnd) return;
    vb6_SdSetProp(hwnd, kSdTickFreq, (LONG)freq);
    if (freq > 0) {
        SendMessageW((HWND)hwnd, TBM_CLEARTICS, TRUE, 0);
        SendMessageW((HWND)hwnd, TBM_SETTICFREQ, (WPARAM)freq, 0);
    }
}

/* ======================= C29-SL-b: 值面 ======================= *
 * 全部直问直发控件，**一格自存都没有**（自存的只有 TickFrequency —— 原生问不出，见上）。
 * 四条实测口径（.build/slprobe/slmeasure6.c）：
 *   · 默认档：min=0 max=100、pos=0、**line=1 page=20**、selstart=0 selend=0。
 *     （VB6 文档给的是 SmallChange=1 / LargeChange=5：小那条对得上，大那条原生是 20，
 *      "没写过 LargeChange 时读回 20 还是折成 5"本机 OCX 未注册、拿不到 VB6 真值 ⇒ 押后，
 *      这里一律**照原生答 20**，不自作主张改成 5。）
 *   · 改 range 会把三样东西一起重算：pos 顶到新下限、page 从 20 变 18、selstart 跟到下限
 *     ⇒ 所以 Init 的顺序是 range → line/page → pos → 刻度 → Sel（见 vb6_Slider_Init 上面）。
 *   · 收窄 range 时 pos 由控件自己钳位（实测 80 → 收到 0..50 之后读回 50）。
 *   · `TBM_GETSELSTART/END` 在**清掉/没设过**时答 **(UINT)-1**（实测 CLEARSEL 之后）
 *     ⇒ 本 getter 把那一个折算成 0（"无区段"），免得 VB6 侧冒出个 -1 的起点。
 */

// VB6 的 Min/Max 是 Long，原生这条消息只有 16 位（实测超界被截成别的数）。
// 下发前钳到 i16，读回也是那一个钳过的值 ⇒ "答出去的"与"控件真走得动的"始终是同一个数。
int vb6_Slider_GetMin(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, TBM_GETRANGEMIN, 0, 0);
}

int vb6_Slider_GetMax(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, TBM_GETRANGEMAX, 0, 0);
}

static void vb6_SdSetRange(void* hwnd, LONG lo, LONG hi) {
    LONG a, b;
    if (!hwnd) return;
    a = (LONG)SendMessageW((HWND)hwnd, TBM_GETRANGEMIN, 0, 0);
    b = (LONG)SendMessageW((HWND)hwnd, TBM_GETRANGEMAX, 0, 0);
    if (lo == a && hi == b) return;          // 幂等: 不重发, 免得 range 一动就连带重算
    SendMessageW((HWND)hwnd, TBM_SETRANGE, TRUE,
                 MAKELONG((WORD)(SHORT)lo, (WORD)(SHORT)hi));
}

void vb6_Slider_SetMin(void* hwnd, int v) {
    LONG hi;
    if (!hwnd) return;
    if (v < VB6_SLD_I16MIN) v = VB6_SLD_I16MIN;
    if (v > VB6_SLD_I16MAX) v = VB6_SLD_I16MAX;
    hi = (LONG)SendMessageW((HWND)hwnd, TBM_GETRANGEMAX, 0, 0);
    if (v > hi) hi = v;                 // 原生不接受 min > max
    vb6_SdSetRange(hwnd, (LONG)v, hi);
}

void vb6_Slider_SetMax(void* hwnd, int v) {
    LONG lo;
    if (!hwnd) return;
    if (v < VB6_SLD_I16MIN) v = VB6_SLD_I16MIN;
    if (v > VB6_SLD_I16MAX) v = VB6_SLD_I16MAX;
    lo = (LONG)SendMessageW((HWND)hwnd, TBM_GETRANGEMIN, 0, 0);
    if (v < lo) lo = v;
    vb6_SdSetRange(hwnd, lo, (LONG)v);
}

int vb6_Slider_GetValue(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, TBM_GETPOS, 0, 0);
}

void vb6_Slider_SetValue(void* hwnd, int v) {
    if (!hwnd) return;
    // 越界**不自己钳**：实测控件就钳（150→100、-5→10），让控件答这一档。
    SendMessageW((HWND)hwnd, TBM_SETPOS, TRUE, (WPARAM)v);
    // C29-SL-c：程序化赋值**不发**通知（实测 TBM_SETPOS / TBM_SETRANGE 一条都不给父窗），
    // 所以这一趟也不该欠下一条 Change：基准跟着推进，后面来一条同值的通知才不会误报"变了"。
    vb6_SdSetProp(hwnd, kSdLastPos, (LONG)(int)SendMessageW((HWND)hwnd, TBM_GETPOS, 0, 0));
}

int vb6_Slider_GetSmallChange(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, TBM_GETLINESIZE, 0, 0);
}

void vb6_Slider_SetSmallChange(void* hwnd, int v) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, TBM_SETLINESIZE, TRUE, (WPARAM)v);
}

int vb6_Slider_GetLargeChange(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, TBM_GETPAGESIZE, 0, 0);
}

void vb6_Slider_SetLargeChange(void* hwnd, int v) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, TBM_SETPAGESIZE, TRUE, (WPARAM)v);
}

// SelectRange = 原生样式位 TBS_ENABLESELRANGE。实测**运行期改这一位有效**（补挂之后
// SETSEL 才答得回来、且换帧后区段画出来）⇒ 写口登记；挂上时按当前 Sel 重发一次 SETSEL，
// 否则位是挂上了、区段还是空的。
int vb6_Slider_GetSelectRange(void* hwnd) {
    if (!hwnd) return 0;
    return ((GetWindowLongPtrW((HWND)hwnd, GWL_STYLE) & TBS_ENABLESELRANGE) != 0) ? -1 : 0;
}

void vb6_Slider_SetSelectRange(void* hwnd, int on) {
    LONG st;
    if (!hwnd) return;
    st = (LONG)GetWindowLongPtrW((HWND)hwnd, GWL_STYLE);
    if (on) {
        if (st & TBS_ENABLESELRANGE) return;
        SetWindowLongPtrW((HWND)hwnd, GWL_STYLE, st | TBS_ENABLESELRANGE);
        vb6_SdReframe((HWND)hwnd);
        // 补挂之后按当前两端重发一次：原生那两位在没挂时 SETSEL 是不生效的（实测）。
        {
            LONG a = vb6_Slider_GetSelStart(hwnd);
            LONG b = vb6_Slider_GetSelEnd(hwnd);
            if (a > b) a = b;
            SendMessageW((HWND)hwnd, TBM_SETSEL, TRUE, MAKELONG((WORD)(SHORT)a, (WORD)(SHORT)b));
        }
    } else {
        if (!(st & TBS_ENABLESELRANGE)) return;
        SetWindowLongPtrW((HWND)hwnd, GWL_STYLE, st & ~(LONG)TBS_ENABLESELRANGE);
        vb6_SdReframe((HWND)hwnd);
    }
}

int vb6_Slider_GetSelStart(void* hwnd) {
    LONG v;
    if (!hwnd) return 0;
    v = (LONG)SendMessageW((HWND)hwnd, TBM_GETSELSTART, 0, 0);
    return (v == -1) ? 0 : (int)v;      // 实测"没设过/已 CLEARSEL"答 (UINT)-1 ⇒ 折成 0
}

int vb6_Slider_GetSelEnd(void* hwnd) {
    LONG v;
    if (!hwnd) return 0;
    v = (LONG)SendMessageW((HWND)hwnd, TBM_GETSELEND, 0, 0);
    return (v == -1) ? 0 : (int)v;
}

// VB6 那侧 SelStart 不能大于 SelEnd（区段是闭区间）。两端各自 setter 时**互相顶**：
// 设起点超过终点 ⇒ 终点跟上来（原生 SETSELSTART 会自己夹，实测分开发 15/60 正常往返）。
void vb6_Slider_SetSelStart(void* hwnd, int v) {
    LONG hi;
    if (!hwnd) return;
    hi = vb6_Slider_GetSelEnd(hwnd);
    if (v > hi) hi = v;
    SendMessageW((HWND)hwnd, TBM_SETSEL, TRUE, MAKELONG((WORD)(SHORT)v, (WORD)(SHORT)hi));
}

void vb6_Slider_SetSelEnd(void* hwnd, int v) {
    LONG lo;
    if (!hwnd) return;
    lo = vb6_Slider_GetSelStart(hwnd);
    if (v < lo) lo = v;
    SendMessageW((HWND)hwnd, TBM_SETSEL, TRUE, MAKELONG((WORD)(SHORT)lo, (WORD)(SHORT)v));
}

/* ======================= C29-SL-c: 事件面 ======================= *
 * 通道实测（.build/slprobe/slmeasure7.c / 8.c，x64 真跑，父窗 WndProc 逐条记 wParam）：
 *   · Slider 与 ScrollBar 同一条道 —— 控件给**父窗**发 WM_HSCROLL（横杆）/ WM_VSCROLL（竖杆），
 *     不走 WM_NOTIFY；`lParam` 就是控件句柄（派发按它认来源），`LOWORD(wParam)` 是 TB_* 码。
 *   · **值在高字**：一次真拖收到 5×N（TB_THUMBTRACK，高字 30,32,35,37,40,42,45,47,49 一路跟着
 *     控件走）→ 4（TB_THUMBPOSITION，高字 = 落点 49）→ 8（TB_ENDTRACK，高字 0）。
 *     所以"拖拽中 Change 连续触发"在原生这边就是那一串 5。
 *   · 方向键（VK_LEFT）收到 0（TB_LINEUP）→ 8，**控件先动**（70→69，动的正是 line 那一格）再发。
 *   · 程序化 `TBM_SETPOS` / `TBM_SETRANGE` **一条都不发**（实测 delta=0）⇒ 与 DT-c 的 `DT41` 同型：
 *     VB6 那颗 OCX 里 Value 赋值会 raise Change，这里刻意照原生、不伪造，夹具 SC 那一条是哨兵。
 * 两条 VB6 事件的分法：
 *   · `Change` 按"**值真的变了**"发（文档原话是 Value 改变即触发、拖拽中连续触发）。拿码表硬枚举
 *     也能凑出同样的序列，但 8（ENDTRACK）与 4（POSITION）都不带新值，用基准比一次就够，
 *     不必记一张"哪些码算变"的表 —— 何况 TB_THUMBTRACK 连发两个同值（拖动不足一像素）时，
 *     按码表会多发一次 Change，按值比不会。
 *   · `Scroll` 只认滑块那两档（4 / 5），与仓里 HScrollBar/VScrollBar 已发货的口径**同一档**
 *     （那两条用的就是 `scrollCode == 5 || == 4`）；同一原生通道不允许两套分法，
 *     "点轨道/按方向键算不算 Scroll"本机拿不到 VB6 真值 ⇒ 押后，见 029。
 */

int32_t vb6_Slider_FireChange(void* hwnd) {
    LONG pos, last;
    if (!hwnd) return 0;
    pos = (LONG)SendMessageW((HWND)hwnd, TBM_GETPOS, 0, 0);
    // 没有基准（Init 之前就来通知）时把当前值当基准 ⇒ 这一条不算变，下一口才开始比。
    last = vb6_SdGetProp(hwnd, kSdLastPos, pos);
    if (pos == last) return 0;
    vb6_SdSetProp(hwnd, kSdLastPos, pos);
    return -1;                       // VB6: True = -1
}

void vb6_Slider_SimNotify(void* hwnd, int32_t code, int32_t pos) {
    HWND hw, parent;
    if (!hwnd) return;
    hw = (HWND)hwnd;
    // 真手势（拖、键、点轨道）都是**控件先动、再发通知**，所以先把值推到 pos ——
    // 高字里带的那个数与控件当时的值必须是同一个，否则判据就在考我们自己的表。
    SendMessageW(hw, TBM_SETPOS, TRUE, (WPARAM)pos);
    parent = GetParent(hw);
    if (!parent) return;
    SendMessageW(parent,
                 (vb6_Slider_GetOrientation(hwnd) != VB6_SLD_HORZ) ? WM_VSCROLL : WM_HSCROLL,
                 MAKEWPARAM((WORD)(SHORT)code, (WORD)(SHORT)pos),
                 (LPARAM)hw);
}

/* 判据专用（不对应 VB6 语义，与上面 SimNotify / DT-c 的 SimChange 同先例）：
 * 把一条**常规事件**的原生消息放进控件自己的队列。为什么只能这样：无头环境点不了鼠标，
 * 而真手势在本机也不稳（.build/slprobe/slmeasure10.c 里 SetForegroundWindow 需要
 * 先敲一次 ALT 才生效，且合成点击偶尔落在激活切换上被吃掉）。
 * 实测（同一目录 slmeasure11.c，逐条 PostMessage 后看子类过程与父窗）：
 *   · WM_LBUTTONUP     → 子类过程收到，父窗**一条通知都不发**、值不动 ⇒ Click 不牵连 Change/Scroll；
 *   · WM_LBUTTONDBLCLK → 同上，只有这一条消息 ⇒ DblClick 与 Click 是两档，不互带；
 *   · WM_KEYDOWN(VK_LEFT=37) → 子类过程收到，**控件真的动**（50→48，line 档 = 2）
 *     并发父窗 code=0；跟着的 WM_KEYUP 发 code=8、值不再动
 *     ⇒ 按键走的就是 SL-c 那条原生通道，Change 会因为值变了而跟一次（夹具按增量数）。
 * kind: 0=Click 1=DblClick 2=KeyDown 3=KeyUp；wParam 只有按键那两档有意义（VB6 的
 * KeyCode 就是它，VK_LEFT 与 vbKeyLeft 同为 37）。
 * 用 SendMessage 而不是 PostMessage：两条实测读数逐字相同（同一枚探针各走一遍，
 * .build/slprobe/sl11_out.txt 的 P* 与 S* 两段），而同步那一条**判据不用夹 DoEvents**
 * —— 夹具在 SimStdEvent 之后紧跟着就读计数器，异步的那条会读到旧值。 */
void vb6_Slider_SimStdEvent(void* hwnd, int32_t kind, int32_t wParam) {
    UINT m;
    if (!hwnd) return;
    /* C29-SL-g: kind=4 不是"发一条消息"，而是**真把焦点给这枚控件**（`SetFocus`）——
     * 焦点事件的原生来源就是窗口管理器自己发的 `WM_SETFOCUS` / `WM_KILLFOCUS`，我们伪造
     * 那两条消息反而验不到真链路（尤其验不到"焦点从 sld6 移到 sld3 时两边各发一次"）。
     * 探针实测这条路是通的：slmeasure9.c 的 Q5a（SetFocus 之后子类过程收到 SETFOCUS）、
     * slmeasure10.c 的 R3（真按下时轨道条自己就抢焦点，紧跟一条 SETFOCUS）。 */
    if (kind == 4) { SetFocus((HWND)hwnd); return; }
    switch (kind) {
    case 0:  m = WM_LBUTTONUP;     break;
    case 1:  m = WM_LBUTTONDBLCLK; break;
    case 2:  m = WM_KEYDOWN;       break;
    default: m = WM_KEYUP;         break;
    }
    SendMessageW((HWND)hwnd, m, (WPARAM)wParam, MAKELPARAM(40, 12));
}

/* ======================= C29-SL-h: TickStyle 四档 + GetNumTicks ======================= *
 * VB6 那一档的真值不是猜的 —— 从 MSCOMCTL.OCX 自带的类型库直接读出来的
 * （探针 .build/slprobe/sltlb.cpp：LoadTypeLibEx(路径, REGKIND_NONE, &tl)，不用注册）：
 *     TickStyleConstants { sldBottomRight = 0, sldTopLeft = 1, sldBoth = 2, sldNoTicks = 3 }
 *     ISlider::TickStyle 的 dispid = 0x0009，型别是 VT_USERDEFINED（就是这张枚举）
 *     ISlider::GetNumTicks 的 dispid = 0x000f，型别 VT_I4，**只有 propget**（VB6 也是只读）
 *
 * 与原生样式位是一张 1:1 的表，不是凑的：commctrl.h 里
 *     TBS_BOTTOM == TBS_RIGHT == 0x0000   TBS_TOP == TBS_LEFT == 0x0004
 *     TBS_BOTH   == 0x0008               TBS_NOTICKS == 0x0010
 * VB6 那两条合名（"BottomRight"/"TopLeft"）念的就是"水平看下半、垂直看右半"，
 * 与上面那两行同值异名的样式位正好对上。
 *
 * 四档在控件侧各有证人，实测于 .build/slprobe/slmeasure12.c（x64 真跑，150x40、range 0..100、
 * freq=10、pos=50；chan=通道矩形、thumb=滑块矩形）：
 *     ts=0  chan.top=10  thumb.top=2   numTics=11
 *     ts=1  chan.top=20  thumb.top=10  numTics=11
 *     ts=2  chan.top=19  thumb.top=10  numTics=11
 *     ts=3  chan.top=10  thumb.top=2   numTics=0     <- 与 ts=0 只差 numTics 这一维
 *   · 所以 **ts=3 的证人是 GetNumTicks**，另外三档靠 chan.top 的**相对高低**分（1 > 2 > 0）。
 *     ⚠ 但 **1 与 2 谁高谁低不稳**：同一枚探针在 40px 高答 20/19，夹具那枚 400 缇（26~27px）
 *     答 18/19 —— 顺序反过来了。所以判据只用「与 ts=0 不同」这一维，1/2 之间不写几何断言，
 *     两档的区分只靠样式位读回（本文件那对 getter/setter）。
 *   · 另一条要记住：`TickPresent`（GETTICPOS(0) != -1）在 NOTICKS 下**照旧答"有刻度"**
 *     （实测仍是 14，刻度只是不画、表还在）⇒ 它看不见 ts=3，别拿它当四档的判据。
 *   · **运行期写样式位 + 换帧**之后的每一条形与"创建时就带那一位"**逐字相同**（12.c 的
 *     Q2 那一段，含来回切四档）⇒ 这一格与 Orientation 同一条口径：不是只能创建时给。
 *
 * 越界值（4、-1…）怎么办：写侧只认 1/2/3 那三位，其余一律落成"下半/右半"那一档（=0），
 * 读侧读的是窗口当前的样式位 ⇒ 与 Orientation 同一口径，**答出去的数就是窗口真在走的那一档**，
 * 不自存、不猜 VB6 会不会报错（OCX 没注册、跑不起来，那条真值本机拿不到）。
 */

// VB6: Slider.TickStyle（读写）。四档就是样式位的三种组合 + "无刻度"。
int vb6_Slider_GetTickStyle(void* hwnd) {
    LONG st;
    if (!hwnd) return 0;
    st = (LONG)GetWindowLongPtrW((HWND)hwnd, GWL_STYLE);
    if (st & TBS_NOTICKS) return 3;   // sldNoTicks（这一位压倒其余两位）
    if (st & TBS_BOTH) return 2;      // sldBoth
    if (st & TBS_TOP) return 1;       // sldTopLeft（竖杆时同一位念作 TBS_LEFT）
    return 0;                          // sldBottomRight
}

void vb6_Slider_SetTickStyle(void* hwnd, int tickStyle) {
    LONG st;
    RECT wr;
    if (!hwnd) return;
    st = (LONG)GetWindowLongPtrW((HWND)hwnd, GWL_STYLE);
    st &= ~(LONG)(TBS_TOP | TBS_BOTH | TBS_NOTICKS);
    if (tickStyle == 1)      st |= TBS_TOP;
    else if (tickStyle == 2) st |= TBS_BOTH;
    else if (tickStyle == 3) st |= TBS_NOTICKS;
    SetWindowLongPtrW((HWND)hwnd, GWL_STYLE, st);
    // 与 Orientation 同一脚：光写样式不重排（实测对照组在 12.c —— 不 FRAMECHANGED 就停在旧布局）
    GetWindowRect((HWND)hwnd, &wr);
    SetWindowPos((HWND)hwnd, NULL, 0, 0, (int)(wr.right - wr.left), (int)(wr.bottom - wr.top),
                 SWP_NOZORDER | SWP_NOMOVE | SWP_FRAMECHANGED);
}

// VB6: Slider.GetNumTicks（只读，dispid 0x000f）—— "当前看得见几条刻度"。
// 原生这条消息在 NOTICKS 下真答 0（实测），所以它同时是 ts=3 那一档的控件侧证人。
int vb6_Slider_GetNumTicks(void* hwnd) {
    if (!hwnd) return 0;
    return (int)SendMessageW((HWND)hwnd, TBM_GETNUMTICS, 0, 0);
}

// C3 扩展（不是 VB6 属性，判据专用）：通道矩形的**上边**。
// ts=0/1/2 三档的差别只在刻度画在哪一侧，而那一侧一换，通道就被顶下去几像素（实测
// 10 / 20 / 19）。夹具用的是**相对高低**而不是绝对像素 —— 绝对值会随主题与控件高度变。
int vb6_Slider_ChannelTop(void* hwnd) {
    RECT ch;
    if (!hwnd) return -1;
    ch.left = ch.top = ch.right = ch.bottom = 0;
    SendMessageW((HWND)hwnd, TBM_GETCHANNELRECT, 0, (LPARAM)&ch);
    return (int)ch.top;
}

// 设计期那一档**只走创建样式位**这一条路（cgen 里两条创建路共用 sliderStyleBits 那一处折算），
// 所以这里刻意不再发一次 setter —— 同一件事不留第二份表。

/* ======================= C29-SL-i: SelLength + ClearSel ======================= *
 * 类型库读数（探针 .build/slprobe/sltlb.cpp）：ISlider 的选区那一对是 **SelStart(0x0007) +
 * SelLength(0x0008)**，两条都是 VT_I4，另有一条方法 **ClearSel(0x000e)**，文档原话
 * "Sets the SelLength to 0"。VB6 那一面**没有 SelEnd 这个名字** —— SelEnd 是 SL-b 按原生
 * TBM_SETSELEND 自己加的口，留着不撤（存量夹具在用），本格把 VB6 真的那一对照着补上。
 *
 * 原生读数（.build/slprobe/slmeasure13.c，x64 真跑，量程 10..100）：
 *   · **什么都没设过**：GETSELSTART=10（就是量程下限！）、GETSELEND=0 ⇒ 终点比起点还小，
 *     拿 `end - start` 会算出负数 ⇒ 长度一律折成 0（"空选区"）。
 *   · SETSEL(20,60) → (20,60)，长度 40。
 *   · SETSEL(20,400) → (20,100)：**远端由量程上限夹住**，不报错、不拒绝。
 *   · SETSEL(80,30)（起点大于终点）→ (80,80)：终点跟上来 ⇒ 长度 0。
 *   · SETSEL(-50,200) → (10,100)：两端都夹进量程。
 *   · CLEARSEL → 两端都答 -1 ⇒ 两个 getter 各折成 0，于是 SelLength 也就是 0
 *     —— 与 VB6 文档那句"把 SelLength 置 0"同形（原生清完之后连起点都问不出来）。
 *   · **没挂 TBS_ENABLESELRANGE 时 SETSEL 整条不生效**（照旧 (10,0)），与 SL-b 那条一致
 *     ⇒ 这一格的 setter 也不伪造：写不进去就是写不进去，读回来还是 0。
 */

// VB6: Slider.SelLength（读写）。起点保持不变，终点 = 起点 + 长度，越界由控件夹进量程。
int vb6_Slider_GetSelLength(void* hwnd) {
    LONG a, b;
    if (!hwnd) return 0;
    a = vb6_Slider_GetSelStart(hwnd);
    b = vb6_Slider_GetSelEnd(hwnd);
    return (b > a) ? (int)(b - a) : 0;   // 默认态 (下限, 0) 与 CLEARSEL 之后都是空选区 ⇒ 0
}

void vb6_Slider_SetSelLength(void* hwnd, int len) {
    LONG a, b;
    if (!hwnd) return;
    // 起点用**裸读数**判："没有选区"时原生答 -1（CLEARSEL 之后就是这一态），而控件自己从没被
    // 碰过的那一态答的是量程下限（实测 10..100 的杆答 GETSELSTART=10）⇒ 两种空态都锚下限。
    // 不这么处理的话，ClearSel 之后写 SelLength 会以折叠出来的 0 为起点、被量程夹成 (10,10)
    // —— 那是一条"看着写进去了、什么都不发生"的静默 no-op，本项目的头号忌讳。
    a = (LONG)SendMessageW((HWND)hwnd, TBM_GETSELSTART, 0, 0);
    if (a == -1) a = (LONG)SendMessageW((HWND)hwnd, TBM_GETRANGEMIN, 0, 0);
    b = a + (len > 0 ? (LONG)len : 0);   // 负长度按 0 处理（原生压根收不到"反向区段"，实测会被顶平）
    SendMessageW((HWND)hwnd, TBM_SETSEL, TRUE, MAKELONG((WORD)(SHORT)a, (WORD)(SHORT)b));
}

// VB6: Slider.ClearSel（方法，0x000e）。运行期调用形：`sld1.ClearSel()`。
void vb6_Slider_ClearSel(void* hwnd) {
    if (!hwnd) return;
    SendMessageW((HWND)hwnd, TBM_CLEARSEL, TRUE, 0);
}

/* ======================= C29-SL-k: Text 气泡 + TextPosition ======================= *
 * 类型库（.build/slprobe/sltlb.cpp）：ISlider 的 **Text = 0x0010，VT_BSTR**，文档原话
 * "the string displayed in the ToolTip as the slider's position changes"；
 * **TextPosition = 0x0011**，枚举 TextPositionConstants { sldAboveLeft = 0, sldBelowRight = 1 }。
 *
 * 走哪条道是量出来的（.build/slprobe/slmeasure17.c，同一份源码编两份、一份用 mt.exe 挂
 * Common-Controls 6.0 清单 —— v5/v6 的差别这一格是决定性的）：
 *   · **轨道条自己那枚 TBS_TOOLTIPS 气泡服务不了自定义串**：它自带 1 条工具，
 *     `TTM_POP` 在无头里什么都不触发（可见性 0、一条 notify 都不发），文本也是它自己画的数字。
 *     而且 TBS_TOOLTIPS **只有创建时给才建得出那枚气泡**（事后写样式位留着但 TBM_GETTOOLTIPS 恒 0
 *     —— 与 DTPicker 的 DTS_SHOWNONE 同族），运行期 RTL 只拿得到 HWND，补不回来。
 *   · **我们自己持一枚 tooltip 宿主 + TRACK 工具是通的**：v6 下 `TTM_ADDTOOLW` 返回 1、
 *     `TTM_GETTOOLCOUNT` 跟着涨、`TTM_TRACKACTIVATE(TRUE)` 之后 **`IsWindowVisible` 就是 1**
 *     （无头也摆得出来），`TTM_GETTEXTW` 能把宿主里的文本原样读回来。
 *     ⇒ 这三条正好是判据要的三把尺：可见性、宿主文本、气泡顶点。
 *   · **v5 下 `TTM_ADDTOOLW` 直接返回 0**（同一份代码）—— 上一格账 #148 差点被这条探针骗了，
 *     所以这格的判据一律写在产物里，不写在探针里。
 *
 * 两条口径要记：① **不复用 P13.8 那枚共享宿主** —— 那枚按 (父窗, 控件句柄) 存的是 `ToolTipText`
 * 的悬停文本，气泡写同一格会互相覆盖，而 VB6 这两条本来就是分开的两样东西；
 * ② `Text` 空着的时候气泡显示**当前的值**（原生那颗就是画数字），写了就用写的串 ——
 * VB6 文档那句只说"显示这个串"，没说两者怎么共存，本机 OCX 跑不起来拿不到真值，这里按①②实现并把
 * 这一条标成"我们的口径"。
 */

static const wchar_t kSdBubbleText[] = L"VB6_SD_Text";        // BSTR 堆拷贝（同 ToolTipText 那族）
static const wchar_t kSdTextPos[]    = L"VB6_SD_TextPosition";  // 0=上/左，1=下/右

static HWND vb6_Slider_BubbleHost(HWND owner) {
    static HWND s_hwnd = NULL;
    if (!s_hwnd) {
        s_hwnd = CreateWindowExW(0, TOOLTIPS_CLASSW, NULL,
                                 WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
                                 CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
                                 owner, NULL, GetModuleHandleW(NULL), NULL);
    }
    return s_hwnd;
}

// 这一次要显示的文本：Text 非空用它，否则用当前的值（原生那颗画的就是数字）。
#define VB6_SLD_BUBBLE_GETPOS (-0x7FFFFFFF)   // 0=现问控件，正数=拖动里带过来的新值
static void vb6_Slider_BubbleString(void* hwnd, wchar_t* buf, int cch, int valueIfUnknown) {
    HANDLE h = GetPropW((HWND)hwnd, kSdBubbleText);
    buf[0] = 0;
    if (h && (BSTR)h && SysStringLen((BSTR)h)) {
        lstrcpynW(buf, (LPCWSTR)h, cch);
        return;
    }
    if (valueIfUnknown == VB6_SLD_BUBBLE_GETPOS) valueIfUnknown = (int)SendMessageW((HWND)hwnd, TBM_GETPOS, 0, 0);
    wsprintfW(buf, L"%d", valueIfUnknown);
}

// 把工具挂上去 / 换文本。TTM_ADDTOOLW 只发一次，之后一律 UPDATE。
static void vb6_Slider_BubbleSetTool(void* hwnd, const wchar_t* text) {
    HWND owner = GetParent((HWND)hwnd);
    HWND host = vb6_Slider_BubbleHost(owner);
    TOOLINFOW ti;
    if (!host) return;
    memset(&ti, 0, sizeof(ti));
    ti.cbSize = sizeof(ti);
    ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS | TTF_TRACK | TTF_ABSOLUTE;
    ti.hwnd = owner;
    ti.uId = (UINT_PTR)hwnd;
    ti.lpszText = (LPWSTR)text;
    if (!SendMessageW(host, TTM_UPDATETIPTEXTW, 0, (LPARAM)&ti)) {
        SendMessageW(host, TTM_ADDTOOLW, 0, (LPARAM)&ti);
    }
}

// VB6: Slider.Text（读写，BSTR）。
wchar_t* vb6_Slider_GetText(void* hwnd) {
    if (!hwnd) return SysAllocString(L"");
    HANDLE h = GetPropW((HWND)hwnd, kSdBubbleText);
    return SysAllocString(h ? (LPCWSTR)h : L"");
}

void vb6_Slider_SetText(void* hwnd, void* bstrText) {
    BSTR old;
    BSTR copy = NULL;
    wchar_t disp[128];
    if (!hwnd) return;
    old = (BSTR)GetPropW((HWND)hwnd, kSdBubbleText);
    if (old) { SysFreeString(old); RemovePropW((HWND)hwnd, kSdBubbleText); }
    if (bstrText) copy = SysAllocString((BSTR)bstrText);
    if (copy) SetPropW((HWND)hwnd, kSdBubbleText, (HANDLE)copy);
    // 立刻把宿主里的文本也换成新的：读回来才算"到控件了"而不是"到我们口袋里了"。
    vb6_Slider_BubbleString(hwnd, disp, (int)(sizeof(disp) / sizeof(disp[0])), VB6_SLD_BUBBLE_GETPOS);
    vb6_Slider_BubbleSetTool(hwnd, disp);
}

// VB6: Slider.TextPosition（读写，0 = sldAboveLeft / 1 = sldBelowRight）。
// 原生没有这条消息（那颗气泡摆哪儿是控件自己定的），所以这一档是自存 + 我们摆放时用。
int vb6_Slider_GetTextPosition(void* hwnd) {
    return (int)vb6_SdGetProp(hwnd, kSdTextPos, 0);
}

void vb6_Slider_SetTextPosition(void* hwnd, int pos) {
    if (!hwnd) return;
    vb6_SdSetProp(hwnd, kSdTextPos, (LONG)((pos == 1) ? 1 : 0));
}

// 把工具填成"这枚滑杆的那一条"，供 ACTIVATE / 读回用。
static void vb6_Slider_BubbleFillTool(void* hwnd, TOOLINFOW* ti) {
    wchar_t disp[128];
    memset(ti, 0, sizeof(*ti));
    ti->cbSize = sizeof(*ti);
    ti->uFlags = TTF_IDISHWND | TTF_SUBCLASS | TTF_TRACK | TTF_ABSOLUTE;
    ti->hwnd = GetParent((HWND)hwnd);
    ti->uId = (UINT_PTR)hwnd;
    vb6_Slider_BubbleString(hwnd, disp, (int)(sizeof(disp) / sizeof(disp[0])), VB6_SLD_BUBBLE_GETPOS);
    ti->lpszText = disp;      // 只在下面这一次调用里用，ACTIVATE 不读它
}

// 派发那边每收到一条滚动通知就问一次：4/5（滑块那两档）把气泡摆出来，8（ENDTRACK）收回去。
// 认不出是轨道条的（同一扇门 ScrollBar 也走这里）直接不管 —— 用类名判，不靠调用方记类型。
void vb6_Slider_BubbleNotify(void* hwnd, int code, int value) {
    HWND host;
    RECT th;
    POINT p;
    wchar_t disp[128];
    TOOLINFOW ti;
    LONG step;
    if (!hwnd) return;
    {
        wchar_t cls[32];
        cls[0] = 0;
        GetClassNameW((HWND)hwnd, cls, 30);
        if (lstrcmpW(cls, L"msctls_trackbar32") != 0) return;
    }
    host = vb6_Slider_BubbleHost(GetParent((HWND)hwnd));
    if (!host) return;
    vb6_Slider_BubbleString(hwnd, disp, (int)(sizeof(disp) / sizeof(disp[0])), value > 0 ? value : VB6_SLD_BUBBLE_GETPOS);
    if (code == 8) {                       // TB_ENDTRACK：收尾，气泡收回去
        vb6_Slider_BubbleFillTool(hwnd, &ti);
        SendMessageW(host, TTM_TRACKACTIVATE, FALSE, (LPARAM)&ti);
        return;
    }
    if (code != 4 && code != 5) return;    // 只有滑块那两档摆气泡（与 Scroll 的分法同一档）
    vb6_Slider_BubbleSetTool(hwnd, disp);
    SendMessageW(hwnd, TBM_GETTHUMBRECT, 0, (LPARAM)&th);
    p.x = th.left;
    p.y = th.top;
    ClientToScreen((HWND)hwnd, &p);
    step = vb6_Slider_GetTextPosition(hwnd) == 1 ? (LONG)(th.bottom - th.top) + 6 : -24;
    SendMessageW(host, TTM_TRACKPOSITION, 0, MAKELPARAM(p.x, p.y + step));
    vb6_Slider_BubbleFillTool(hwnd, &ti);
    SendMessageW(host, TTM_TRACKACTIVATE, TRUE, (LPARAM)&ti);
}

// 判据证人（C3 扩展，不是 VB6 属性）：气泡此刻摆没摆出来（实测无头也答 1）。
int vb6_Slider_BubbleVisible(void* hwnd) {
    HWND host;
    if (!hwnd) return 0;
    host = vb6_Slider_BubbleHost(GetParent((HWND)hwnd));
    return (host && IsWindowVisible(host)) ? -1 : 0;   // VB6: True = -1
}

// 判据证人：气泡的**上边**（TextPosition 两档的差别就落在这里；夹具用相对高低判，
// 不钉绝对像素 —— 那随 DPI/主题变，与 TickStyle 那一格同一条教训）。
int vb6_Slider_BubbleTop(void* hwnd) {
    HWND host;
    RECT r;
    if (!hwnd) return -1;
    host = vb6_Slider_BubbleHost(GetParent((HWND)hwnd));
    if (!host) return -1;
    r.left = r.top = r.right = r.bottom = 0;
    GetWindowRect(host, &r);
    return (int)r.top;
}

// 判据证人：宿主里此刻挂着的那句文本（问的是 tooltip 控件，不是我们的窗口属性）。
wchar_t* vb6_Slider_BubbleText(void* hwnd) {
    static wchar_t buf[128];
    HWND host;
    TOOLINFOW ti;
    buf[0] = 0;
    if (!hwnd) return SysAllocString(L"");
    host = vb6_Slider_BubbleHost(GetParent((HWND)hwnd));
    if (!host) return SysAllocString(L"");
    vb6_Slider_BubbleFillTool(hwnd, &ti);
    ti.lpszText = buf;
    buf[0] = 0;
    /* TTM_GETTEXT 的返回值不能当成败判据：实测 v6 下它**回 0 而文本照样复制进缓冲**
       （探针 .build/slprobe/slmeasure18.c：GETTEXTW rc=0 text=<VOL>）⇒ 只看缓冲。
       与账 #148 那枚 ToolTipRegistered 同一处坑、同一个修法。 */
    SendMessageW(host, TTM_GETTEXTW, 0, (LPARAM)&ti);
    return SysAllocString(buf);
}

#endif /* _WIN32 */
