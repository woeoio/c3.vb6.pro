// vb6forms_mscomm.c — VB6 MSComm 控件 (MSCommLib.MSComm / MSCOMM32.OCX) 复刻
//
// 用 **Win32 串口 API 直连**（CreateFileW("\\\\.\\COMx") + DCB + ReadFile 读线程），
// **不加载 MSCOMM32.OCX**。这一枚在 VB6 里没有外观（无边框、无鼠标面），但**要真收事件** ——
// 与 Winsock 同一形状：控件自己有一枚身份窗（类名 VB6_MSCOMM，注册在本文件），读线程把
// "有数据了"投成一条私有消息，窗口过程里再判 RThreshold 去调用户注册的 OnComm。
//
// 为什么是"每端口一条读线程"而不是 WaitCommEvent 异步：
//   · 只有 DCB 档（波特率/校验/数据位/停止位）与 RThreshold 这一档要判，WaitCommEvent 的
//     EV_RXCHAR 又**只报"有字符到了"一次**，凑 RThreshold 还得自己数；读线程 ReadFile 阻塞
//     等数据，天然就是"到多少给多少"；
//   · 关端口要能立刻停 —— 读线程是**同步** ReadFile（间隔超时 50ms），关端口时主线程
//     `CancelSynchronousIo(读线程句柄)` 一下就从 ReadFile 里醒（实测：overlapped 那版在 NUL 上
//     读写都不认，已在文件里记了读数）。
//     不会卡在 ReadFile 里等到下一个字节。
//
// 与 VB6 的口径差异（诚实口径，见 ai/内置控件/MSComm 控件（串口通讯）.md §6.5）：
//   · `InputMode` / `InputLen` 收值与回读，但 **Input 一律按文本（ACP）交付**；二进制
//     （Byte 数组）那一形未做 —— 不上报"每字节一个字符"的假数组。
//   · `Output` 收 BSTR（文本）；Byte 数组写那一形同上未做。
//   · CTS/DSR/CD 的**变化事件**（comEvCTS/DSR/CD=3/4/5）未做：要 WaitCommEvent 的 EV_CTS 档，
//     与上面那条"不用 WaitCommEvent"的选择冲突；三条 Holding **属性**照问不误。
//
// 测试钩子（**只给夹具用，不是 VB6 API**）：环境变量 `C3_MSCOMM_DEVICE` 是设备名（如 `\\.\NUL`），
// 设了就顶替 `\\.\COM<CommPort>`。本机没有串口硬件，判据要在 NUL 上跑"开-写-计数-关闭"这一路；
// NUL 不是通讯设备，DCB 那几步会失败，所以打开流程把"设备不认 DCB"当**降级**而不是失败。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6forms_prop_ctrl.h"

#include <oleauto.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <wchar.h>

#ifdef _WIN32

#pragma comment(lib, "kernel32.lib")

// 读线程投的私有消息。仓库里已用的号：0x7FF0（延迟 Form_Load）、0x7FFF（恢复焦点）、
// WM_APP+0x15（Winsock）。这里取 WM_APP+0x16，理由与 Winsock 那枚相同 —— 通知只投给
// 控件自己的身份窗，一枚号就够，不需要 per-control。
#define WM_MS_NOTIFY   (WM_APP + 0x16)

#define VB6_MS_MAX          32      // 一个工程里的 MSComm 上限（含控件数组展开后的下标）
#define VB6_MS_RX_MAX       (1u << 20)   // 接收缓冲天花板 1 MB，超了丢新的（与 Winsock 同档）
#define VB6_MS_READ_CHUNK   4096    // 一次 ReadFile 交多少字节

typedef void (*vb6_MsCbVoid)(void);

// CommEvent 号：VB6 那两族的数（MSComm 文档 §3），cgen 侧发的是同一个字面量。
#define VB6_MS_EV_SEND      1
#define VB6_MS_EV_RECEIVE   2
#define VB6_MS_EV_CTS       3
#define VB6_MS_EV_DSR       4
#define VB6_MS_EV_CD        5
#define VB6_MS_EV_RING      6
#define VB6_MS_EV_EOF       7
#define VB6_MS_ERR_FRAME    1004
#define VB6_MS_ERR_OVERRUN  1006
#define VB6_MS_ERR_RXOVER   1008
#define VB6_MS_ERR_PARITY   1009
#define VB6_MS_ERR_TXFULL   1010

struct vb6_MsComm {
    HWND      hwnd;                 // 身份窗；NULL = 空槽
    HANDLE    h;                    // 串口句柄（INVALID_HANDLE_VALUE = 关）
    HANDLE    th;                   // 读线程（同步 ReadFile，卡在里面等字节）
    HANDLE    thModem;              // 信号线/线路错误线程（WaitCommEvent）
    LONG      closing;              // 两条线程共用的退出旗（Interlocked）
    CRITICAL_SECTION cs;            // 保护 rx 缓冲（读线程 append / 主线程 Input 取走）
    int       csReady;

    int32_t   commPort;             // 1..16（VB6 口径），0 = 没写；**不钳 16**，COM100 也行
    wchar_t   device[256];          // C3 扩展 CommDevice：任意设备名（\\\\.\\CNCA0 之类），空 = 按 CommPort 拼
    wchar_t   settings[64];         // 原样回读（VB6 的 Settings 就是"你写进来的那串"）
    int32_t   inputMode;            // 0 文本 / 1 二进制（本批只影响回读，见文件头）
    int32_t   inputLen;             // 0 = 一次读光
    int32_t   rThreshold;           // 收到多少字节触发 comEvReceive；0 = 不触发
    int32_t   sThreshold;           // 发送缓冲降到多少以下触发 comEvSend；0 = 不触发
    int32_t   inBufSize, outBufSize;
    int32_t   handshaking;
    int32_t   dtr, rts;
    int32_t   eofEnable;
    // C3 扩展（VB6 没有超时属性）：三项超时 / 一个 BREAK 开关的存值。
    // readInterval = -1 表示"没写过" ⇒ 打开时用默认 50ms。
    int32_t   readInterval, readTotalConst, writeTimeout;
    int32_t   commEvent;            // 最近一次事件号（CommEvent 属性）
    int32_t   openOk;               // PortOpen
    int32_t   degraded;             // 打开成功但设备不认 DCB（NUL 那一类）
    int32_t   rxSignaled;           // 本轮 comEvReceive 已报过（Input 取空后复位）
    int32_t   txTotal;              // 累计写出去的字节数（C3 证人；VB6 没有这条属性）
    int32_t   rxTotal;              // 累计收到的字节数（同上）

    char*     rx;
    uint32_t  rxHead, rxLen, rxCap;

    vb6_MsCbVoid cb;                // OnComm（VB6 里只有这一个事件）
};

static struct vb6_MsComm g_ms[VB6_MS_MAX];
static LONG g_msCsGuard = 0;        // 0 未起 / 1 正在起 / 2 已起（每槽 cs 懒初始化）

static DWORD WINAPI vb6_MsModemThread(LPVOID param);   // 定义在读线程之后（open 要先看见）

static void vb6_MsEnsureCs(void) {
    int i;
    if (InterlockedCompareExchange(&g_msCsGuard, 1, 0) == 0) {
        for (i = 0; i < VB6_MS_MAX; i++) {
            InitializeCriticalSection(&g_ms[i].cs);
            g_ms[i].csReady = 1;
        }
        InterlockedExchange(&g_msCsGuard, 2);
        return;
    }
    while (InterlockedCompareExchange(&g_msCsGuard, 2, 2) != 2) Sleep(1);
}

static struct vb6_MsComm* vb6_MsOf(HWND h) {
    int i;
    if (!h) return NULL;
    for (i = 0; i < VB6_MS_MAX; i++) if (g_ms[i].hwnd == h) return &g_ms[i];
    return NULL;
}

static struct vb6_MsComm* vb6_MsSlot(HWND h) {
    int i;
    struct vb6_MsComm* e = vb6_MsOf(h);
    if (e) return e;
    vb6_MsEnsureCs();
    for (i = 0; i < VB6_MS_MAX; i++) {
        if (!g_ms[i].hwnd) {
            struct vb6_MsComm* s = &g_ms[i];
            // ⚠ 逐个字段复位，**不能** ZeroMemory 一大块 —— 那会把 cs / csReady
            // （懒初始化过的那把锁）一起抹掉，下一个 EnterCriticalSection 就死在未初始化上。
            s->h = INVALID_HANDLE_VALUE;
            s->th = NULL;
            s->thModem = NULL;
            s->closing = 0;
            s->commPort = 0;
            s->inputMode = 0;
            s->inputLen = 0;
            s->rThreshold = 0;
            s->sThreshold = 0;
            s->inBufSize = 1024;
            s->outBufSize = 512;
            s->handshaking = 0;
            s->dtr = 0;
            s->rts = 0;
            s->eofEnable = 0;
            s->readInterval = -1;
            s->readTotalConst = -1;
            s->writeTimeout = -1;
            s->device[0] = 0;
            s->commEvent = 0;
            s->openOk = 0;
            s->degraded = 0;
            s->rxSignaled = 0;
            s->txTotal = 0;
            s->rxTotal = 0;
            s->rx = NULL;
            s->rxHead = s->rxLen = s->rxCap = 0;
            s->cb = NULL;
            wcscpy(s->settings, L"9600,N,8,1");
            s->hwnd = h;                 // hwnd 最后写：它是"这个槽被占了"的旗
            return s;
        }
    }
    return NULL;
}

// 事件：写完 CommEvent 再调回调 —— 用户在处理器里读 .CommEvent 必须已经是新的那个号
// （Winsock 那边同一条纪律，见 vb6forms_winsock.c 的 vb6_WsFire）。
static void vb6_MsFire(struct vb6_MsComm* e, int32_t code) {
    if (!e) return;
    e->commEvent = code;
    if (e->cb) ((vb6_MsCbVoid)e->cb)();
}

static int vb6_MsRxAppend(struct vb6_MsComm* e, const char* src, uint32_t n) {
    uint32_t need;
    if (!e->csReady) return 0;
    EnterCriticalSection(&e->cs);
    if (e->rxHead && e->rxLen + n > e->rxCap) {
        memmove(e->rx, e->rx + e->rxHead, e->rxLen);   // 压实：队头以后再挪
        e->rxHead = 0;
    }
    need = e->rxLen + n;
    if (need > VB6_MS_RX_MAX) { LeaveCriticalSection(&e->cs); return 0; }
    if (need > e->rxCap) {
        uint32_t cap = e->rxCap ? e->rxCap : 4096;
        char* p;
        while (cap < need) cap <<= 1;
        p = (char*)realloc(e->rx, cap);
        if (!p) { LeaveCriticalSection(&e->cs); return 0; }
        e->rx = p;
        e->rxCap = cap;
    }
    memcpy(e->rx + e->rxHead + e->rxLen, src, n);
    e->rxLen += n;
    e->rxTotal += (int32_t)n;
    LeaveCriticalSection(&e->cs);
    return 1;
}

// 取走（VB6 的 Input：读后清空；InputLen 决定一次取多少）
static BSTR vb6_MsTake(struct vb6_MsComm* e, int32_t maxLen) {
    uint32_t take;
    int wn;
    BSTR res;
    if (!e || !e->csReady) return SysAllocString(L"");
    EnterCriticalSection(&e->cs);
    take = e->rxLen;
    if (maxLen > 0 && (uint32_t)maxLen < take) take = (uint32_t)maxLen;
    wn = MultiByteToWideChar(CP_ACP, 0, e->rx + e->rxHead, (int)take, NULL, 0);
    res = SysAllocStringLen(NULL, wn > 0 ? (UINT)wn : 0);
    if (res && wn > 0) MultiByteToWideChar(CP_ACP, 0, e->rx + e->rxHead, (int)take, res, wn);
    e->rxHead += take;
    e->rxLen -= take;
    if (!e->rxLen) e->rxHead = 0;
    if (e->rxLen < (uint32_t)(e->rThreshold > 0 ? e->rThreshold : 1)) e->rxSignaled = 0;
    LeaveCriticalSection(&e->cs);
    return res ? res : SysAllocString(L"");
}

static int32_t vb6_MsRxLen(struct vb6_MsComm* e) {
    int32_t n;
    if (!e || !e->csReady) return 0;
    EnterCriticalSection(&e->cs);
    n = (int32_t)e->rxLen;
    LeaveCriticalSection(&e->cs);
    return n;
}

// 通知到窗口：判 RThreshold 报 comEvReceive；EOF 那一档由读线程单独投一个负数哨兵。
static void vb6_MsOnNotify(struct vb6_MsComm* e, LPARAM what) {
    if (!e) return;
    if (what < 0) {
        // 负数 = “某个事件号到了”（取反存）：-7 EOF / -3 CTS / -4 DSR / -5 CD /
        // -1004 帧错 / -1006 超速 / -1008 收缓冲溢出 / -1009 校验错。
        // 两张来源统一在这一格落：读线程（EOF / 读错误）与信号线线程（线路变化 / CE 错）。
        int32_t code = (int32_t)(-what);
        if (code == VB6_MS_EV_EOF && !e->eofEnable) return;   // EOFEnable=False 时不要这号
        vb6_MsFire(e, code);
        return;
    }
    if (e->rThreshold > 0 && !e->rxSignaled && vb6_MsRxLen(e) >= e->rThreshold) {
        e->rxSignaled = 1;
        vb6_MsFire(e, VB6_MS_EV_RECEIVE);
    }
}

static DWORD WINAPI vb6_MsReaderThread(LPVOID param) {
    struct vb6_MsComm* e = (struct vb6_MsComm*)param;
    char buf[VB6_MS_READ_CHUNK];
    // **同步** ReadFile 阻塞等字节（COMMTIMEOUTS 的间隔超时让它 50ms 醒一次），关端口时
    // 主线程用 CancelSynchronousIo 把它叫醒。这一版是把 overlapped 那版换掉的：实测
    // `\\.\NUL` 拿 FILE_FLAG_OVERLAPPED 打开后，ReadFile/WriteFile 不带 OVERLAPPED 一律
    // ERROR_INVALID_PARAMETER（第一版夹具读数就是 MC3-tx=0、MC5-ev=1008）。
    for (;;) {
        DWORD n = 0;
        BOOL ok;
        if (InterlockedCompareExchange(&e->closing, 0, 0)) break;
        ok = ReadFile(e->h, buf, sizeof(buf), &n, NULL);
        if (!ok) {
            DWORD err = GetLastError();
            if (err == ERROR_OPERATION_ABORTED || err == ERROR_INVALID_HANDLE) break;  // 关端口时正常路径
            PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(-(intptr_t)VB6_MS_ERR_RXOVER));
            break;
        }
        if (n == 0) {
            // 0 字节 = 文件尾。真串口不会这样（同步读没数据是先超时再返回 0）；`\\.\NUL`
            // （夹具的降级设备）每次读到"文件尾"就是 0 字节 —— 报一次 EOF 就退出线程，
            // 不退会在 NUL 上空转成 100% CPU。
            PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(-(intptr_t)VB6_MS_EV_EOF));
            break;
        }
        if (!vb6_MsRxAppend(e, buf, n)) {
            PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(-(intptr_t)VB6_MS_ERR_RXOVER));
            break;
        }
        PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(intptr_t)n);
    }
    return 0;
}

// 信号线变化 / 线路错误：WaitCommEvent 是阻塞的，所以单开一条线程（degraded 档不起）。
// 关端口时主线程同样用 CancelSynchronousIo 叫醒它。NUL 这一类设备 SetCommMask 会失败
// 或 WaitCommEvent 直接报错 ⇒ 歇 20ms 再来，不空转（夹具的降级路因此不会烧 CPU）。
// CE_* → VB6 那几个错误号的映射：帧错 1004 / 超速 1006 / 收缓冲溢出 1008 / 校验错 1009。
static DWORD WINAPI vb6_MsModemThread(LPVOID param) {
    struct vb6_MsComm* e = (struct vb6_MsComm*)param;
    if (!e) return 0;
    SetCommMask(e->h, EV_CTS | EV_DSR | EV_RLSD | EV_ERR);
    for (;;) {
        DWORD got = 0;
        if (InterlockedCompareExchange(&e->closing, 0, 0)) break;
        if (!WaitCommEvent(e->h, &got, NULL)) {
            DWORD err = GetLastError();
            if (err == ERROR_OPERATION_ABORTED || err == ERROR_INVALID_HANDLE) break;
            Sleep(20);
            continue;
        }
        if (got & EV_CTS)  PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(-(intptr_t)VB6_MS_EV_CTS));
        if (got & EV_DSR)  PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(-(intptr_t)VB6_MS_EV_DSR));
        if (got & EV_RLSD) PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(-(intptr_t)VB6_MS_EV_CD));
        if (got & EV_ERR) {
            DWORD errs = 0;
            COMSTAT st;
            int32_t code = 0;
            if (ClearCommError(e->h, &errs, &st)) {
                if (errs & CE_FRAME)         code = VB6_MS_ERR_FRAME;
                else if (errs & CE_OVERRUN)  code = VB6_MS_ERR_OVERRUN;
                else if (errs & CE_RXPARITY) code = VB6_MS_ERR_PARITY;
                else if (errs & CE_RXOVER)   code = VB6_MS_ERR_RXOVER;
            }
            if (code) PostMessageW(e->hwnd, WM_MS_NOTIFY, 0, (LPARAM)(-(intptr_t)code));
        }
    }
    return 0;
}

static void vb6_MsClosePort(struct vb6_MsComm* e) {
    if (!e) return;
    InterlockedExchange(&e->closing, 1);
    if (e->th) {
        // 同步 ReadFile 卡在读线程里：CancelSynchronousIo 是把它叫醒的唯一一手
        //（CloseHandle 不会取消正在进行的同步 I/O）。叫不醒就等它按间隔超时自己返回。
        CancelSynchronousIo(e->th);
        WaitForSingleObject(e->th, 2000);
        CloseHandle(e->th);
        e->th = NULL;
    }
    if (e->thModem) {
        CancelSynchronousIo(e->thModem);          // WaitCommEvent 同样要叫
        WaitForSingleObject(e->thModem, 2000);
        CloseHandle(e->thModem);
        e->thModem = NULL;
    }
    if (e->h != INVALID_HANDLE_VALUE && e->h) { CloseHandle(e->h); }
    e->h = INVALID_HANDLE_VALUE;
    e->openOk = 0;
    e->degraded = 0;
    e->closing = 0;
}

// Settings 解析："波特率,校验,数据位,停止位"（VB6 那一串）。返回 1 成功。
// 为什么要自己拆而不用 BuildCommDCBW：那一条把**解析失败**吞成「返回 FALSE」，而它的
// 高波特率（921600/1500000/3000000）在不同 SDK 版本上的接受度并不一致 —— 精度不够时
// 就会让 DCB 停在 GetCommState 拿到的旧值（9600），形成一个**看着打开成功、实际跑错速度**
// 的静默缺陷。这里自己拆、自己判，失败就硬失败（PortOpen=False）。
static int vb6_MsParseSettings(const wchar_t* s, DCB* dcb) {
    wchar_t buf[128];
    wchar_t* ctx = NULL;
    wchar_t* tok;
    long baud;
    int field = 0, databits = 8, stop = 0;
    wchar_t parity = L'N';
    if (!s || !*s || !dcb) return 0;
    wcsncpy(buf, s, 127);
    buf[127] = 0;
    tok = wcstok(buf, L",", &ctx);
    while (tok) {
        // 去掉两侧空白（.frm 里常见 "9600, N, 8, 1"）
        while (*tok == L' ' || *tok == L'\t') tok++;
        { wchar_t* end = tok + wcslen(tok);
          while (end > tok && (end[-1] == L' ' || end[-1] == L'\t')) *--end = 0; }
        if (field == 0) {
            wchar_t* endp = NULL;
            baud = wcstol(tok, &endp, 10);
            if (!endp || *endp || baud < 50 || baud > 12000000L) return 0;
            dcb->BaudRate = (DWORD)baud;
        } else if (field == 1) {
            parity = (wchar_t)towupper(tok[0]);
            if (!parity) return 0;
        } else if (field == 2) {
            databits = _wtoi(tok);
            if (databits < 5 || databits > 8) return 0;
        } else if (field == 3) {
            if (wcscmp(tok, L"1") == 0) stop = 1;
            else if (wcscmp(tok, L"1.5") == 0) stop = 2;
            else if (wcscmp(tok, L"2") == 0) stop = 3;
            else return 0;
        }
        field++;
        tok = wcstok(NULL, L",", &ctx);
    }
    if (field < 4) return 0;
    switch (parity) {
    case L'N': dcb->Parity = NOPARITY;   dcb->fParity = FALSE; break;
    case L'E': dcb->Parity = EVENPARITY; dcb->fParity = TRUE;  break;
    case L'O': dcb->Parity = ODDPARITY;  dcb->fParity = TRUE;  break;
    case L'M': dcb->Parity = MARKPARITY; dcb->fParity = TRUE;  break;
    case L'S': dcb->Parity = SPACEPARITY; dcb->fParity = TRUE; break;
    default: return 0;
    }
    dcb->ByteSize = (BYTE)databits;
    dcb->StopBits = (stop == 1) ? ONESTOPBIT : (stop == 2 ? ONE5STOPBITS : TWOSTOPBITS);
    dcb->fBinary = TRUE;
    return 1;
}

// 打开：设备名 = CommDevice（C3 扩展）> C3_MSCOMM_DEVICE（夹具钩子）> \\.\COM<CommPort>。
// 返回 1 成功。两种“不认 DCB”的情形分开处理：
//   · GetCommState 失败 ⇒ 设备不是通讯设备（`\\.\NUL`、虚拟对的一半）⇒ **降级档**，读写照走；
//   · GetCommState 成功但 Settings 解析/SetCommState 失败 ⇒ **硬失败**（关句柄、PortOpen=False），
//     不再“PortOpen=True 但按旧 DCB 跑”那个静默错。
static int vb6_MsOpenPort(struct vb6_MsComm* e) {
    wchar_t dev[256];
    wchar_t hook[256];
    HANDLE h;
    DCB dcb;
    COMMTIMEOUTS to;
    DWORD hookLen;
    if (!e) return 0;
    if (e->openOk) return 1;
    if (e->device[0]) {
        wcsncpy(dev, e->device, 255);
        dev[255] = 0;
    } else {
        hookLen = GetEnvironmentVariableW(L"C3_MSCOMM_DEVICE", hook, 256);
        if (hookLen > 0 && hookLen < 256) {
            wcsncpy(dev, hook, 255);
            dev[255] = 0;
        } else {
            int port = e->commPort > 0 ? e->commPort : 1;
            swprintf(dev, 256, L"\\\\.\\COM%d", port);
        }
    }
    h = CreateFileW(dev, GENERIC_READ | GENERIC_WRITE, 0, NULL, OPEN_EXISTING,
                    FILE_ATTRIBUTE_NORMAL, NULL);
    if (h == INVALID_HANDLE_VALUE) return 0;      // 端口被占 / 不存在：PortOpen 保持 False
    e->h = h;
    ZeroMemory(&dcb, sizeof(dcb));
    dcb.DCBlength = sizeof(dcb);
    e->degraded = 0;
    if (GetCommState(h, &dcb)) {
        if (!vb6_MsParseSettings(e->settings, &dcb)) {
            CloseHandle(h); e->h = INVALID_HANDLE_VALUE;
            return 0;                              // Settings 写错：硬失败，不再静默按旧 DCB 跑
        }
        // Handshaking：0 无 / 1 XOnXOff / 2 RTS / 3 两者（VB6 那四个数）
        dcb.fOutxCtsFlow = (e->handshaking == 2 || e->handshaking == 3) ? TRUE : FALSE;
        dcb.fRtsControl = (e->handshaking == 2 || e->handshaking == 3)
                              ? RTS_CONTROL_HANDSHAKE : (e->rts ? RTS_CONTROL_ENABLE : RTS_CONTROL_DISABLE);
        dcb.fOutX = (e->handshaking == 1 || e->handshaking == 3) ? TRUE : FALSE;
        dcb.fInX  = (e->handshaking == 1 || e->handshaking == 3) ? TRUE : FALSE;
        dcb.fDtrControl = e->dtr ? DTR_CONTROL_ENABLE : DTR_CONTROL_DISABLE;
        if (!SetCommState(h, &dcb)) {
            CloseHandle(h); e->h = INVALID_HANDLE_VALUE;
            return 0;                              // 设备不接这一档波特率：硬失败
        }
        SetupComm(h, (DWORD)(e->inBufSize > 0 ? e->inBufSize : 1024),
                      (DWORD)(e->outBufSize > 0 ? e->outBufSize : 512));
    } else {
        e->degraded = 1;               // NUL 这一类：不是通讯设备，DCB 问不出来
    }
    ZeroMemory(&to, sizeof(to));
    to.ReadIntervalTimeout = (DWORD)(e->readInterval >= 0 ? e->readInterval : 50);
    to.ReadTotalTimeoutMultiplier = 0;
    to.ReadTotalTimeoutConstant = (DWORD)(e->readTotalConst >= 0 ? e->readTotalConst : 0);
    to.WriteTotalTimeoutConstant = (DWORD)(e->writeTimeout >= 0 ? e->writeTimeout : 2000);
    SetCommTimeouts(h, &to);            // 对 NUL 无效，无害

    InterlockedExchange(&e->closing, 0);
    e->th = CreateThread(NULL, 0, vb6_MsReaderThread, e, 0, NULL);
    if (!e->th) {
        CloseHandle(h); e->h = INVALID_HANDLE_VALUE;
        return 0;
    }
    // 信号线变化事件那条线程（只在真串口上有意义；NUL 上 WaitCommEvent 会返回 FALSE
    // 并歇下来，不空转 —— 见 vb6_MsModemThread 里的 Sleep）。
    if (!e->degraded) {
        e->thModem = CreateThread(NULL, 0, vb6_MsModemThread, e, 0, NULL);
    }
    e->openOk = 1;
    return 1;
}

// ============================ 窗口类 / 生命周期 ============================

static LRESULT CALLBACK vb6_MsWndProc(HWND h, UINT m, WPARAM wp, LPARAM lp) {
    struct vb6_MsComm* e = vb6_MsOf(h);
    switch (m) {
    case WM_MS_NOTIFY:
        vb6_MsOnNotify(e, (LPARAM)lp);
        return 0;
    case WM_DESTROY:
        if (e) vb6_MsClosePort(e);
        return 0;
    }
    return DefWindowProcW(h, m, wp, lp);
}

void vb6_RegisterMsCommClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return;
    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = vb6_MsWndProc;
    wc.hInstance = (HINSTANCE)hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = L"VB6_MSCOMM";
    if (RegisterClassW(&wc)) done = 1;
}

void* vb6_MsComm_Create(void* hwnd) {
    if (hwnd) vb6_MsSlot((HWND)hwnd);
    return hwnd;
}

void vb6_MsComm_Destroy(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) return;
    vb6_MsClosePort(e);
    free(e->rx);
    e->rx = NULL; e->rxCap = e->rxLen = e->rxHead = 0;
    e->rxSignaled = 0;
    e->hwnd = NULL;
    e->cb = NULL;
}

void vb6_MsComm_SetEventHandler(void* hwnd, int32_t kind, void* fn) {
    struct vb6_MsComm* e;
    if (kind != VB6_MS_EV_ONCOMM) return;
    e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (e) e->cb = (vb6_MsCbVoid)fn;
}

// ============================ 属性面 ============================

int32_t vb6_MsComm_GetCommPort(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->commPort : 0;
}

void vb6_MsComm_SetCommPort(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e || e->openOk) return;          // 打开着改端口号：VB6 也是无效（要关掉重开）
    e->commPort = v;
}

const void* vb6_MsComm_GetSettings(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return SysAllocString(e ? e->settings : L"");
}

void vb6_MsComm_SetSettings(void* hwnd, const void* bstr) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    const wchar_t* s = (const wchar_t*)bstr;
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e || !s || e->openOk) return;
    wcsncpy(e->settings, s, 63);
    e->settings[63] = 0;
}

int32_t vb6_MsComm_GetPortOpen(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->openOk : 0;
}

void vb6_MsComm_SetPortOpen(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e) return;
    if (v) {
        if (!vb6_MsOpenPort(e)) { e->openOk = 0; e->commEvent = 0; }
    } else {
        vb6_MsClosePort(e);
    }
}

int32_t vb6_MsComm_GetInputMode(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->inputMode : 0;
}

void vb6_MsComm_SetInputMode(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (e) e->inputMode = v;
}

int32_t vb6_MsComm_GetInputLen(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->inputLen : 0;
}

void vb6_MsComm_SetInputLen(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (e) e->inputLen = v;
}

// Input：读后清空（VB6 口径）。InputLen=0 一次读光；>0 只取那么多个。
const void* vb6_MsComm_GetInput(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return vb6_MsTake(e, e ? e->inputLen : 0);
}

// Output：文本按 ACP 落成字节写出去；写不出去（端口没开）静默不写，CommEvent 不动。
void vb6_MsComm_SetOutput(void* hwnd, const void* bstr) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    const wchar_t* s = (const wchar_t*)bstr;
    char* bytes;
    int need, n;
    DWORD written = 0;
    if (!e || !s || !e->openOk || e->h == INVALID_HANDLE_VALUE) return;
    need = WideCharToMultiByte(CP_ACP, 0, s, -1, NULL, 0, NULL, NULL);
    if (need <= 0) return;
    bytes = (char*)malloc((size_t)need);
    if (!bytes) return;
    n = WideCharToMultiByte(CP_ACP, 0, s, -1, bytes, need, NULL, NULL);
    if (n > 0) n -= 1;                     // 去掉结尾 NUL：线上不带它
    if (n > 0) {
        if (WriteFile(e->h, bytes, (DWORD)n, &written, NULL) && written > 0) {
            e->txTotal += (int32_t)written;
            // SThreshold：VB6 是"发送缓冲降到该值以下"报 comEvSend；同步写完之后缓冲恒空，
            // 所以设了它就报一次 —— 这正是"写完了"那一档的语义（差异见文件头）。
            if (e->sThreshold > 0) vb6_MsFire(e, VB6_MS_EV_SEND);
        }
    }
    free(bytes);
}

int32_t vb6_MsComm_GetInBufferCount(void* hwnd) {
    return vb6_MsRxLen(vb6_MsOf((HWND)hwnd));
}

void vb6_MsComm_SetInBufferCount(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e || v != 0 || !e->csReady) return;   // VB6 只有"赋 0 = 清空"这一档有意义
    EnterCriticalSection(&e->cs);
    e->rxHead = e->rxLen = 0;
    e->rxSignaled = 0;
    LeaveCriticalSection(&e->cs);
    // 驱动里那半也得倒：只清我们的缓冲，下一次 ReadFile 会把旧字节又读上来。
    if (e->openOk && e->h != INVALID_HANDLE_VALUE) PurgeComm(e->h, PURGE_RXCLEAR);
}

int32_t vb6_MsComm_GetOutBufferCount(void* hwnd) {
    // 同步写 ⇒ 我们这层没有队列，但**驱动里**可能还有没发完的（大块写 + 慢波特率）。
    // 所以真去问 ClearCommError 的 cbOutQue，而不是恒答 0（那是“看不见”不是“没有”）。
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    COMSTAT st;
    DWORD errs = 0;
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE) return 0;
    ZeroMemory(&st, sizeof(st));
    if (!ClearCommError(e->h, &errs, &st)) return 0;
    return (int32_t)st.cbOutQue;
}

void vb6_MsComm_SetOutBufferCount(void* hwnd, int32_t v) {
    (void)hwnd; (void)v;                   // 同上：没有待发残留可清
}

int32_t vb6_MsComm_GetInBufferSize(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->inBufSize : 1024;
}

void vb6_MsComm_SetInBufferSize(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e || e->openOk) return;           // VB6 口径：打开着改缓冲大小无效
    if (v > 0) e->inBufSize = v;
}

int32_t vb6_MsComm_GetOutBufferSize(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->outBufSize : 512;
}

void vb6_MsComm_SetOutBufferSize(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e || e->openOk) return;
    if (v > 0) e->outBufSize = v;
}

int32_t vb6_MsComm_GetRThreshold(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->rThreshold : 0;
}

void vb6_MsComm_SetRThreshold(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (e) { e->rThreshold = v; e->rxSignaled = 0; }
}

int32_t vb6_MsComm_GetSThreshold(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->sThreshold : 0;
}

void vb6_MsComm_SetSThreshold(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (e) e->sThreshold = v;
}

int32_t vb6_MsComm_GetHandshaking(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->handshaking : 0;
}

void vb6_MsComm_SetHandshaking(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e || e->openOk) return;
    e->handshaking = v;
}

int32_t vb6_MsComm_GetDTREnable(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->dtr : 0;
}

void vb6_MsComm_SetDTREnable(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e) return;
    // 运行期生效（VB6 允许打开后改）：真串口直接拉线，降级设备只存值。
    if (e->openOk && e->h != INVALID_HANDLE_VALUE && !e->degraded)
        EscapeCommFunction(e->h, v ? SETDTR : CLRDTR);
    e->dtr = v ? 1 : 0;
}

int32_t vb6_MsComm_GetRTSEnable(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->rts : 0;
}

void vb6_MsComm_SetRTSEnable(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e) return;
    if (e->openOk && e->h != INVALID_HANDLE_VALUE && !e->degraded)
        EscapeCommFunction(e->h, v ? SETRTS : CLRRTS);
    e->rts = v ? 1 : 0;
}

int32_t vb6_MsComm_GetEOFEnable(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->eofEnable : 0;
}

void vb6_MsComm_SetEOFEnable(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (e) e->eofEnable = v;
}

int32_t vb6_MsComm_GetCommEvent(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->commEvent : 0;
}

// CTS/DSR/CD：VB6 只读。真串口问 GetCommModemStatus；NUL 这类设备问不出来 ⇒ 答 0
// （**不**拿"上一次的"假装，宁可答一个诚实的 0）。
int32_t vb6_MsComm_GetCTSHolding(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    DWORD st = 0;
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE) return 0;
    if (!GetCommModemStatus(e->h, &st)) return 0;
    return (st & MS_CTS_ON) ? 1 : 0;
}

int32_t vb6_MsComm_GetDSRHolding(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    DWORD st = 0;
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE) return 0;
    if (!GetCommModemStatus(e->h, &st)) return 0;
    return (st & MS_DSR_ON) ? 1 : 0;
}

int32_t vb6_MsComm_GetCDHolding(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    DWORD st = 0;
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE) return 0;
    if (!GetCommModemStatus(e->h, &st)) return 0;
    return (st & MS_RLSD_ON) ? 1 : 0;
}

// C3 扩展（不是 VB6 属性）：夹具用它把"写了个数"升级成"真写进设备了"的证人。
int32_t vb6_MsComm_GetTxTotal(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->txTotal : 0;
}

int32_t vb6_MsComm_GetRxTotal(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? e->rxTotal : 0;
}

// ============================ C3 扩展：超时 / 设备名 / 信号线 / 枚举 / 二进制 ============================// 这五组都不是 VB6 的属性（VB6 的 MSComm 没有超时、没有枚举、没有 Break、没有二进制属性面），
// 但它们恰好是“现代设备”最常要的几件：高波特率、USB CDC 的自定义设备名、帧间隙超时、
// 用字节数组收发。口径一律：**能真下发就真下发，下不了就首次上升为失败，不静默吞**。

static void vb6_MsApplyTimeouts(struct vb6_MsComm* e) {
    COMMTIMEOUTS to;
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE) return;
    ZeroMemory(&to, sizeof(to));
    to.ReadIntervalTimeout = (DWORD)(e->readInterval >= 0 ? e->readInterval : 50);
    to.ReadTotalTimeoutMultiplier = 0;
    to.ReadTotalTimeoutConstant = (DWORD)(e->readTotalConst >= 0 ? e->readTotalConst : 0);
    to.WriteTotalTimeoutConstant = (DWORD)(e->writeTimeout >= 0 ? e->writeTimeout : 2000);
    SetCommTimeouts(e->h, &to);
}

int32_t vb6_MsComm_GetReadIntervalTimeout(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? (e->readInterval >= 0 ? e->readInterval : 50) : 50;
}

void vb6_MsComm_SetReadIntervalTimeout(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e) return;
    e->readInterval = v < 0 ? 0 : v;      // 帧间隙：分帧协议把它当“一帧结束”的判据
    vb6_MsApplyTimeouts(e);
}

int32_t vb6_MsComm_GetReadTotalTimeout(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? (e->readTotalConst >= 0 ? e->readTotalConst : 0) : 0;
}

void vb6_MsComm_SetReadTotalTimeout(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e) return;
    e->readTotalConst = v < 0 ? 0 : v;
    vb6_MsApplyTimeouts(e);
}

int32_t vb6_MsComm_GetWriteTimeout(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return e ? (e->writeTimeout >= 0 ? e->writeTimeout : 2000) : 2000;
}

void vb6_MsComm_SetWriteTimeout(void* hwnd, int32_t v) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e) return;
    e->writeTimeout = v < 0 ? 0 : v;
    vb6_MsApplyTimeouts(e);
}

// CommDevice：任意设备名（`\\.\CNCA0`、com0com 的 `\\.\CNCA0`、USB CDC 的 `\\.\COM120` 都行）。
// 这就是“突破 MSComm 16 个端口限制”的正式口——CommPort 本身其实也没钳 16，但它只能表达“COM<n>”。
const void* vb6_MsComm_GetCommDevice(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    return SysAllocString(e ? e->device : L"");
}

void vb6_MsComm_SetCommDevice(void* hwnd, const void* bstr) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    const wchar_t* s = (const wchar_t*)bstr;
    if (!e) e = vb6_MsSlot((HWND)hwnd);
    if (!e || e->openOk) return;           // 开着改设备名与改 CommPort 同口径：先关
    if (!s) { e->device[0] = 0; return; }
    wcsncpy(e->device, s, 255);
    e->device[255] = 0;
}

// BREAK 与两根信号线的**运行期**拉放（VB6 的 DTREnable/RTSEnable 只管打开那一刻）。
void vb6_MsComm_SetBreak(void* hwnd, int32_t on) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE || e->degraded) return;
    EscapeCommFunction(e->h, on ? SETBREAK : CLRBREAK);
}

int32_t vb6_MsComm_SetDTRLine(void* hwnd, int32_t on) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE || e->degraded) return 0;
    e->dtr = on ? 1 : 0;
    return EscapeCommFunction(e->h, on ? SETDTR : CLRDTR) ? 1 : 0;
}

int32_t vb6_MsComm_SetRTSLine(void* hwnd, int32_t on) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    if (!e || !e->openOk || e->h == INVALID_HANDLE_VALUE || e->degraded) return 0;
    e->rts = on ? 1 : 0;
    return EscapeCommFunction(e->h, on ? SETRTS : CLRRTS) ? 1 : 0;
}

// 端口枚举：QueryDosDeviceW(L"COM", …) 拿到的是双 NUL 结尾的多字符串
// （"COM1\0COM3\0COM7\0\0"）—— 这台机器上**真实存在**的那些，不是注册表里的历史残留。
// 友好名（"USB-SERIAL CH340 (COM7)"）需要 SetupDi 遍历，不在这一批（见文档 §5.5）。
static DWORD vb6_MsEnumComPorts(wchar_t* buf, DWORD cch) {
    DWORD n;
    if (!buf || cch < 8) return 0;
    n = QueryDosDeviceW(L"COM", buf, cch);
    return n;                              // 0 = 失败（枚举不到任何 COM 设备也是 0）
}

int32_t vb6_MsComm_PortCount(void* hwnd) {
    static wchar_t buf[16384];
    DWORD n;
    DWORD i = 0;
    int32_t cnt = 0;
    (void)hwnd;                            // 全局信息：接口形状统一成“控件成员读”才带这一参
    n = vb6_MsEnumComPorts(buf, 16384);
    while (i < n && buf[i]) {
        cnt++;
        i += (DWORD)wcslen(buf + i) + 1;
    }
    return cnt;
}

// PortList：`COM1;COM3;COM7`。VB6 侧没有枚举，所以这一形与 PortCount 都是 C3 扩展；
// 用一条 BSTR 而不是 PortName(i) 是因为属性读写只能带一个 hwnd 实参（不支持下标）。
const void* vb6_MsComm_GetPortList(void* hwnd) {
    static wchar_t buf[16384];
    wchar_t joined[16384];
    DWORD n, i = 0;
    int first = 1;
    joined[0] = 0;
    (void)hwnd;
    n = vb6_MsEnumComPorts(buf, 16384);
    while (i < n && buf[i]) {
        if (!first) wcscat(joined, L";");
        wcsncat(joined, buf + i, 16000);
        first = 0;
        i += (DWORD)wcslen(buf + i) + 1;
    }
    return SysAllocString(joined);
}

// 二进制收：**每字节一个字符**的 BSTR（码位 0..255，BSTR 长度前缀所以内嵌 NUL 也能装）。
// 为什么不用 Byte() Variant：cgen 的属性面只支持标量/String 这一档，数组那一形要按目标类型
// 分派（VB6 的 `v = .Input: b = v` 属于那条），本批先把“真字节不出错”落地 —— 取回后用
// AscW(Mid$(s, i + 1, 1)) 就能还原每个字节（夹具里就是这么验的）。
const void* vb6_MsComm_GetInputBytes(void* hwnd) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    BSTR res;
    uint32_t take;
    wchar_t* p;
    uint32_t k;
    if (!e || !e->csReady) return SysAllocString(L"");
    EnterCriticalSection(&e->cs);
    take = e->rxLen;
    if (e->inputLen > 0 && (uint32_t)e->inputLen < take) take = (uint32_t)e->inputLen;
    res = SysAllocStringLen(NULL, take);
    p = res;
    for (k = 0; k < take && p; k++) p[k] = (wchar_t)(unsigned char)e->rx[e->rxHead + k];
    e->rxHead += take;
    e->rxLen -= take;
    if (!e->rxLen) e->rxHead = 0;
    if (e->rxLen < (uint32_t)(e->rThreshold > 0 ? e->rThreshold : 1)) e->rxSignaled = 0;
    LeaveCriticalSection(&e->cs);
    return res ? res : SysAllocString(L"");
}

// 二进制发：交进来的 BSTR 每位取低 8 位写一个字节（不过码页）。
void vb6_MsComm_SetOutputBytes(void* hwnd, const void* bstr) {
    struct vb6_MsComm* e = vb6_MsOf((HWND)hwnd);
    const wchar_t* s = (const wchar_t*)bstr;
    char buf[1024];
    int n, i;
    DWORD written = 0;
    if (!e || !s || !e->openOk || e->h == INVALID_HANDLE_VALUE) return;
    n = (int)SysStringLen((BSTR)s);
    if (n <= 0) return;
    n = n > (int)sizeof(buf) ? (int)sizeof(buf) : n;
    for (i = 0; i < n; i++) buf[i] = (char)(s[i] & 0xFF);
    if (WriteFile(e->h, buf, (DWORD)n, &written, NULL) && written > 0) {
        e->txTotal += (int32_t)written;
        if (e->sThreshold > 0) vb6_MsFire(e, VB6_MS_EV_SEND);
    }
}

#else  /* !_WIN32 */
void vb6_RegisterMsCommClass(void* hInstance) { (void)hInstance; }

#endif /* _WIN32 */
