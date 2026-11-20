// vb6forms_updown.c — VB6 UpDown 控件 (ai/034 §二-5, ai/029 C30-UD)
//
// 复刻口径: 用 comctl32 的 **msctls_updown32** 复刻 VB6 UpDown 的等价行为,
// 不加载任何 OCX (mscomctl.ocx 只有 32 位, 在本机未注册, x64 进程里装不进来)。
//
// 为什么要有这个文件: 登记之前 UpDown 落在 "未识别类型 ⇒ Fix 143 真 OCX 宿主" 那一支,
// x64 下 CoCreateInstance 直接失败 —— 句柄恒 NULL、属性读写全打空、退出码照旧 0。
// 这正是 ai/029 §二-2 说的那条静默失败, 也是 ai/034 §六 把它列进 P0 的原因。
//
// 本格 (C30-UD-a) 的范围:
//   · 窗口 + 创建样式 (UDS_WRAP / UDS_ALIGNLEFT|RIGHT / UDS_HORZ / UDS_SETBUDDYINT);
//   · 值面 Min / Max / Value / Increment;
//   · 行为面 Wrap / Alignment / Orientation / SyncBuddy / BuddyControl;
//   · 事件面 UpClick / DownClick / Change 的原生源 (UDN_DELTAPOS, 见 vb6_UpDown_OnNotify)。
//
// 证据口径 (诚实标注, 别把设计决定写成实测):
//   · 消息号与样式位取自 commctrl.h (UDM_SETRANGE32 = WM_USER+111 一族, UDS_* 见头),
//     不是猜的; 但**本格没有单独写 .build 探针去量原生默认值** —— 所以下面凡写
//     "原生默认" 的地方都是 Win32 文档口径, 判据一律以夹具读回为准 (tests/ctrlupdown),
//     不拿注释当证人。
//   · 与 Slider 那一批不同, 这里没有 channel/thumb 那种几何假证人可踩: UpDown 的
//     可见结果就是 buddy 里的数与 UDN_DELTAPOS 的 iDelta, 两条都能直接读回。
//
// 消息号: 直接写字面量, 不依赖 commctrl.h 的版本宏 (mstcp 老 SDK 里没有 SETPOS32)。
//   UDM_SETRANGE    = WM_USER+101 (0x0465)
//   UDM_GETRANGE32  = WM_USER+112 (0x0470)
//   UDM_SETRANGE32  = WM_USER+111 (0x046F)
//   UDM_SETPOS32    = WM_USER+113 (0x0471)
//   UDM_GETPOS32    = WM_USER+114 (0x0472)
//   UDM_SETBUDDY    = WM_USER+105 (0x0469)
//   UDM_GETBUDDY    = WM_USER+106 (0x046A)
//   UDM_SETACCEL    = WM_USER+107 (0x046B)
//   UDM_GETACCEL    = WM_USER+108 (0x046C)

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#endif

#include <stdint.h>

#include "vb6forms.h"
#include "vb6forms_internal.h"

#ifdef _WIN32

#define VU_UDM_SETRANGE    0x0465u
#define VU_UDM_GETRANGE32  0x0470u
#define VU_UDM_SETRANGE32  0x046Fu
#define VU_UDM_SETPOS32    0x0471u
#define VU_UDM_GETPOS32    0x0472u
#define VU_UDM_SETBUDDY    0x0469u
#define VU_UDM_GETBUDDY    0x046Au
#define VU_UDM_SETACCEL    0x046Bu
#define VU_UDM_GETACCEL    0x046Cu

#define VU_UDS_WRAP        0x0001L
#define VU_UDS_SETBUDDYINT 0x0002L
#define VU_UDS_ALIGNRIGHT  0x0004L
#define VU_UDS_ALIGNLEFT   0x0008L
#define VU_UDS_AUTOBUDDY   0x0010L
#define VU_UDS_ARROWKEYS   0x0020L
#define VU_UDS_HORZ        0x0040L
#define VU_UDS_NOTHOUSANDS 0x0080L
#define VU_UDS_HOTTRACK    0x0100L

// UDN_FIRST = 0U-721U = -721；UDN_DELTAPOS = UDN_FIRST - 1 = **-722**（SDK 头
// D:\Windows Kits\10\include\10.0.19041.0\um\commctrl.h 逐行核对过）。
// ⚠ 本批第一版把这里写成 -721（把 UDN_FIRST 当成了 DELTAPOS），合成判据两边
// 用同一个错常量所以全绿；**真鼠标手势一次就把它抓出来了**：值照变、buddy 照跟，
// 而 UpClick/DownClick/Change 一条都没有（读数见 tools/updown_manual/UpManual.log，
// seq=1、A v=10、up/dn/ch 全 0）。常量是对的、表也是全的，只是比对的号错了。
#define VU_UDN_DELTAPOS    (-722)

// UDACCEL = { UINT nSec; UINT nInc; } —— 两个 UINT，8 字节（SDK commctrl.h）。
// ⚠ 本批第一版把它写成 `{UINT fVirt; WORD wKey; WORD wInc;}` 并填 wInc=3，落进内存变成
// nSec=0、nInc=(0 | 3<<16)=196608 ⇒ **每次点击直接冲到量程边界**。合成判据又两边自洽
// （写读同一个错结构体，“Increment=3”读得回来），所以全绿；**真鼠标一测就穿**：
// 读数 tools/updown_manual/UpManual.log 第一轮（A 5→10、B 10→0、D 0→20，每个控件
// 都压在极值上，而 up=2/dn=1）。这是“自洽并不等于对”的第二个样本。
typedef struct tagVU_UDACCEL {
    UINT nSec;
    UINT nInc;
} VU_UDACCEL;

// -999 是本项目的 "没写过" 哨兵 (与 vb6_Slider_Init 同一口径): 没写就**不覆盖**
// 原生默认值, 免得把"什么都没写"的控件顶成 0..0 / 0 那一档。
#define VU_UNSET (-999)

static void vu_setStyle(HWND h, long add, long remove) {
    LONG_PTR st = GetWindowLongPtrW(h, GWL_STYLE);
    st = (st & ~(LONG_PTR)remove) | (LONG_PTR)add;
    SetWindowLongPtrW(h, GWL_STYLE, st);
    // 样式位换了还得让控件重排一次, 否则外观停在旧档 (与 Slider 的 Orientation 同一处手法)。
    SetWindowPos(h, NULL, 0, 0, 0, 0,
                 SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
}

static HWND vu_hwnd(void* hwnd) { return (HWND)hwnd; }

// C30-UD-a: 创建/设计期那一趟。**顺序**是 range → increment → value → 行为面 → buddy:
// 先立量程再立值, 否则越界值会被旧量程钳一次; buddy 必须最后设, 因为 UDM_SETBUDDY
// 之后控件会按 buddy 里的现有文本重算一次位置。
void vb6_UpDown_Init(void* hwnd, long min, long max, long value, long increment,
                     long wrap, long alignment, long orientation, long syncBuddy,
                     void* buddy) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;

    long lo = min, hi = max;
    if (lo == VU_UNSET) lo = 0;       // VB6 的 UpDown 默认量程 0..100
    if (hi == VU_UNSET) hi = 100;
    if (lo > hi) { long t = lo; lo = hi; hi = t; }
    SendMessageW(h, VU_UDM_SETRANGE32, (WPARAM)lo, (LPARAM)hi);

    if (increment != VU_UNSET) {
        VU_UDACCEL acc;
        acc.nSec = 0;
        acc.nInc = (UINT)(increment > 0 ? increment : 1);
        SendMessageW(h, VU_UDM_SETACCEL, 1, (LPARAM)&acc);
    }

    if (value != VU_UNSET) SendMessageW(h, VU_UDM_SETPOS32, 0, (LPARAM)value);

    // 行为面: 两条都既写样式位又真下发一次, 因为容器那条创建路可能根本没挂样式。
    if (wrap != VU_UNSET) {
        if (wrap) vu_setStyle(h, VU_UDS_WRAP, 0);
        else      vu_setStyle(h, 0, VU_UDS_WRAP);
    }
    if (alignment != VU_UNSET) {
        if (alignment != 0) vu_setStyle(h, VU_UDS_ALIGNRIGHT, VU_UDS_ALIGNLEFT);
        else                vu_setStyle(h, VU_UDS_ALIGNLEFT, VU_UDS_ALIGNRIGHT);
    }
    if (orientation != VU_UNSET) {
        if (orientation != 0) vu_setStyle(h, VU_UDS_HORZ, 0);
        else                  vu_setStyle(h, 0, VU_UDS_HORZ);
    }
    if (syncBuddy != VU_UNSET) {
        if (syncBuddy) vu_setStyle(h, VU_UDS_SETBUDDYINT, 0);
        else           vu_setStyle(h, 0, VU_UDS_SETBUDDYINT);
    }
    if (buddy) SendMessageW(h, VU_UDM_SETBUDDY, (WPARAM)buddy, 0);
}

// ---- 值面 ---------------------------------------------------------------
// 全部直问直发控件, 不自存 (与 Slider 的 TickFrequency 相反: UpDown 的 range/pos/accel
// 都有原生回读消息, 自存一份只会多一个可能对不上的答案)。

int32_t vb6_UpDown_GetMin(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 0;
    int32_t lo = 0, hi = 0;
    SendMessageW(h, VU_UDM_GETRANGE32, (WPARAM)&lo, (LPARAM)&hi);
    return lo;
}

int32_t vb6_UpDown_GetMax(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 0;
    int32_t lo = 0, hi = 0;
    SendMessageW(h, VU_UDM_GETRANGE32, (WPARAM)&lo, (LPARAM)&hi);
    return hi;
}

void vb6_UpDown_SetMin(void* hwnd, int32_t v) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    int32_t lo = 0, hi = 0;
    SendMessageW(h, VU_UDM_GETRANGE32, (WPARAM)&lo, (LPARAM)&hi);
    SendMessageW(h, VU_UDM_SETRANGE32, (WPARAM)v, (LPARAM)hi);
}

void vb6_UpDown_SetMax(void* hwnd, int32_t v) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    int32_t lo = 0, hi = 0;
    SendMessageW(h, VU_UDM_GETRANGE32, (WPARAM)&lo, (LPARAM)&hi);
    SendMessageW(h, VU_UDM_SETRANGE32, (WPARAM)lo, (LPARAM)v);
}

// 越界由控件自己钳 (与 Slider 的 SETPOS 同一档语义), 我们不再钳第二遍 ——
// 钳两遍就会出现"我们以为的值"和"控件真在显示的值"两个答案。
int32_t vb6_UpDown_GetValue(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 0;
    return (int32_t)SendMessageW(h, VU_UDM_GETPOS32, 0, 0);
}

void vb6_UpDown_SetValue(void* hwnd, int32_t v) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    SendMessageW(h, VU_UDM_SETPOS32, 0, (LPARAM)v);
}

int32_t vb6_UpDown_GetIncrement(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 1;
    VU_UDACCEL acc;
    acc.nSec = 0; acc.nInc = 0;
    // nAccel=1: 只问第 0 档（按键加速那一格我们不下发, VB6 的 Increment 就是步长）。
    if (SendMessageW(h, VU_UDM_GETACCEL, 1, (LPARAM)&acc) <= 0) return 1;
    return (int32_t)acc.nInc;
}

void vb6_UpDown_SetIncrement(void* hwnd, int32_t v) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    VU_UDACCEL acc;
    acc.nSec = 0;
    acc.nInc = (UINT)(v > 0 ? v : 1);
    SendMessageW(h, VU_UDM_SETACCEL, 1, (LPARAM)&acc);
}

// ---- 行为面 -------------------------------------------------------------
// 四条都是样式位, 读回就问窗口自己 (GetWindowLongPtrW), 不另存一份。

int32_t vb6_UpDown_GetWrap(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 0;
    return (GetWindowLongPtrW(h, GWL_STYLE) & VU_UDS_WRAP) ? -1 : 0;
}

void vb6_UpDown_SetWrap(void* hwnd, int32_t on) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    if (on) vu_setStyle(h, VU_UDS_WRAP, 0);
    else    vu_setStyle(h, 0, VU_UDS_WRAP);
}

// VB6: Alignment 0 = Left, 1 = Right (默认 1)。
int32_t vb6_UpDown_GetAlignment(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 1;
    return (GetWindowLongPtrW(h, GWL_STYLE) & VU_UDS_ALIGNLEFT) ? 0 : 1;
}

void vb6_UpDown_SetAlignment(void* hwnd, int32_t v) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    if (v != 0) vu_setStyle(h, VU_UDS_ALIGNRIGHT, VU_UDS_ALIGNLEFT);
    else        vu_setStyle(h, VU_UDS_ALIGNLEFT, VU_UDS_ALIGNRIGHT);
}

// VB6: Orientation 0 = Vertical (默认), 1 = Horizontal。
int32_t vb6_UpDown_GetOrientation(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 0;
    return (GetWindowLongPtrW(h, GWL_STYLE) & VU_UDS_HORZ) ? 1 : 0;
}

void vb6_UpDown_SetOrientation(void* hwnd, int32_t v) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    if (v != 0) vu_setStyle(h, VU_UDS_HORZ, 0);
    else        vu_setStyle(h, 0, VU_UDS_HORZ);
}

int32_t vb6_UpDown_GetSyncBuddy(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return 0;
    return (GetWindowLongPtrW(h, GWL_STYLE) & VU_UDS_SETBUDDYINT) ? -1 : 0;
}

void vb6_UpDown_SetSyncBuddy(void* hwnd, int32_t on) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    if (on) vu_setStyle(h, VU_UDS_SETBUDDYINT, 0);
    else    vu_setStyle(h, 0, VU_UDS_SETBUDDYINT);
}

void* vb6_UpDown_GetBuddy(void* hwnd) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return NULL;
    return (void*)SendMessageW(h, VU_UDM_GETBUDDY, 0, 0);
}

void vb6_UpDown_SetBuddy(void* hwnd, void* buddy) {
    HWND h = vu_hwnd(hwnd);
    if (!h) return;
    SendMessageW(h, VU_UDM_SETBUDDY, (WPARAM)buddy, 0);
}

// ---- 事件面 -------------------------------------------------------------
// 原生 UpDown 把 UDN_DELTAPOS 发到**父窗**, 不是发给自己 ⇒ 拍在窗体的 WM_NOTIFY 里
// (cgen 那条 arm)。返回 iDelta: >0 是"上/加", <0 是"下/减", 0 表示没有位移。
// 与 Slider 的 Scroll 同一分工: 这里只回答"往哪边动了几格", 点哪条处理器由发码那侧定。
// 通知到时控件**还没改值**（实测：udA 在 5 上点▲，处理器读到 v=5，松手才变 6；
// udB 在 10 上点▼，处理器读到 v=10，之后才变 9）。所以“这次会不会真动”只能按
// 老位置 + iDelta×Increment 往前推一格（wrap 按环算，否则钳到量程）。
// VB6 语义：UpClick/DownClick 是“点了箭头”，边界上也该响；Change 是“值变了”，边界上
// 不该响 —— 所以返回值带上“动没动”这一位：
//   0 = 不是我们的通知；±1 = 动了（正=上、负=下）；±2 = 点了但值没动。
static int32_t vu_nextPos(HWND h, int32_t pos, int32_t delta) {
    int32_t lo = 0, hi = 0;
    int32_t inc = vb6_UpDown_GetIncrement((void*)h);
    long long t;
    SendMessageW(h, VU_UDM_GETRANGE32, (WPARAM)&lo, (LPARAM)&hi);
    if (hi < lo) { int32_t tmp = lo; lo = hi; hi = tmp; }
    t = (long long)pos + (long long)delta * (long long)inc;
    if (GetWindowLongPtrW(h, GWL_STYLE) & VU_UDS_WRAP) {
        long long span = (long long)hi - (long long)lo + 1;
        long long off = (t - (long long)lo) % span;
        if (off < 0) off += span;
        return (int32_t)((long long)lo + off);
    }
    if (t < (long long)lo) t = (long long)lo;
    if (t > (long long)hi) t = (long long)hi;
    return (int32_t)t;
}

int32_t vb6_UpDown_OnNotify(void* hwnd, int32_t code, void* lParam, int32_t* outDelta) {
    HWND h = vu_hwnd(hwnd);
    const unsigned char* p;
    HWND from;
    int32_t delta, pos, next;
    if (code != VU_UDN_DELTAPOS) return 0;
    if (!lParam || !h) return 0;
    // NMUPDOWN { NMHDR hdr; int iPos; int iDelta; }
    p = (const unsigned char*)lParam;
    from = (HWND)(((const NMHDR*)p)->hwndFrom);
    if (from != h) return 0;
    // iDelta 在 64 位下仍在偏移 16 处: NMHDR 是三枚指针宽字段 (hwndFrom/idFrom UINT_PTR,
    // code UINT), 之后的 iPos/iDelta 各 4 字节 ⇒ 用结构布局取, 别手算对齐。
    memcpy(&delta, p + sizeof(NMHDR) + sizeof(int), sizeof(delta));
    if (outDelta) *outDelta = delta;
    if (delta == 0) return 0;
    pos = (int32_t)SendMessageW(h, VU_UDM_GETPOS32, 0, 0);
    next = vu_nextPos(h, pos, delta);
    return (delta > 0) ? ((next != pos) ? 1 : 2) : ((next != pos) ? -1 : -2);
}

// C30-UD-c: 把这次位移真落下去，然后才允许发 Change —— 这样 Change 处理器里读到的
// 是**新值**（真 VB6 的 Change 是“值已变”，实测我们以前给的是旧值）。
// 因为 cgen 那一侧随后会 `return 1` 把原生那次实施取消掉，所以这里必须自己设到位，
// 否则值根本不会动；buddy 文本也跟着 UDS_SETBUDDYINT 在这一次 SETPOS 中同步。
int32_t vb6_UpDown_Commit(void* hwnd, int32_t delta) {
    HWND h = vu_hwnd(hwnd);
    int32_t pos, next;
    if (!h || delta == 0) return 0;
    pos = (int32_t)SendMessageW(h, VU_UDM_GETPOS32, 0, 0);
    next = vu_nextPos(h, pos, delta);
    SendMessageW(h, VU_UDM_SETPOS32, 0, (LPARAM)next);
    return (int32_t)SendMessageW(h, VU_UDM_GETPOS32, 0, 0);
}

// ---- 判据专用助手 (C30-UD-b) -------------------------------------------
// 与 Slider.SimNotify / DTPicker.SimChange / MonthView.SimDateClick 同先例:
// 不对应 VB6 语义, 不进手册的方法表。存在理由只有一条 —— 无头环境点不了箭头,
// 而**直接调处理器**会绕开「认来源 → 分流 → 调用」那三段, 等于没验派发链。
//
// 次序按真手势: 控件先把值移到位, 再把 UDN_DELTAPOS 递到父窗 (父窗在处理器里
// 还能改 iDelta 把这次位移取消)。所以这里也先动值、再发通知 —— 只发通知不动值,
// 判据就退化成「在考我们自己的表」。
//
// 两处**判断而非实测**的口径, 如实标在这里 (本机无法在无头环境点箭头,
// 所以「原生在量程边界上到底发不发 UDN_DELTAPOS」没有读数):
//   1) delta 的单位是**点击格数**(+1 = 按一次上箭头), Increment 的乘法在这里做; 原生
//      iDelta 的绝对值就是 wInc 的倍数, 所以两者只在 Increment != 1 时分得开;
//   2) 值没动(顶到量程且没开 UDS_WRAP)时**发 iDelta=0**, 让 Change/UpClick 都不响 ——
//      对齐 VB6 的「没变就不触发 Change」, 而不是照抄原生可能仍发一次通知的行为。
//      夹具刻意不落在边界上, 判据不依赖这条判断。
typedef struct tagVU_NMUPDOWN {
    NMHDR hdr;    // hwndFrom / idFrom / code
    int   iPos;   // 位移后的位置
    int   iDelta; // 位移格数 (正=上/加, 负=下/减)
} VU_NMUPDOWN;

int32_t vb6_UpDown_SimDelta(void* hwnd, int32_t delta) {
    HWND h = vu_hwnd(hwnd);
    HWND parent;
    int32_t pos, next;
    VU_NMUPDOWN nmu;
    if (!h) return 0;
    pos = (int32_t)SendMessageW(h, VU_UDM_GETPOS32, 0, 0);
    if (delta == 0) return pos;
    next = vu_nextPos(h, pos, delta);

    // 先发通知、后落值 —— 与真手势同序（实测见上面 vu_nextPos 那段的读数）。
    // 第一版是“先改后发”，处理器里读到的自然是新值，看起来“很对”，但那是另一回事：
    // 真 VB6 的 UpClick/DownClick 在值变更**前**发，依赖“处理器里读 Value”的代码只有
    // 分清这两条时间线才对得上。C30-UD-c 之后：Up/Down 仍读**旧值**，Change 由 arm
    // 先 Commit 再发，读**新值**（与真 VB6 一致）。这里发完通知后自己再 SETPOS 一次是
    // 冗余但无害的（arm 已经 Commit 到同一个位置）。
    parent = GetParent(h);
    if (parent) {
        nmu.hdr.hwndFrom = h;
        nmu.hdr.idFrom = (UINT_PTR)GetDlgCtrlID(h);
        nmu.hdr.code = (UINT)VU_UDN_DELTAPOS;
        nmu.iPos = (int)next;
        nmu.iDelta = (int)delta;      // 边界上也照原号发（与原生一致），会不会响由 OnNotify 判
        // SendMessage 同步: 负载在栈上, 返回时接收方已经用完, 不存在悬空。
        SendMessageW(parent, WM_NOTIFY, nmu.hdr.idFrom, (LPARAM)&nmu);
    }
    SendMessageW(h, VU_UDM_SETPOS32, 0, (LPARAM)next);
    return (int32_t)SendMessageW(h, VU_UDM_GETPOS32, 0, 0);
}

#endif // _WIN32
