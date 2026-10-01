// vb6forms_winsock.c — ai/029 C29-WS: VB6 Winsock 控件的原生落点（Winsock2）
//
// 本格范围 = 控件身份（不可见窗 + **真在收通知的 WndProc**）+ 实例表 + 事件槽 + 状态与属性面
//   + **UDP 一整轮**（Bind / SendData / GetData / PeekData / Close + DataArrival / StateChanged）。
// TCP 的 Listen / ConnectionRequest / Accept / Connect（服务端与客户端互为前提，拆不开就
// 一起做）由 WS-b 接上；Error / SendProgress / SendComplete / Byte 数组那一形 GetData 留 WS-c。
//
// 三条决定这里形状的测量（理由与读数在 vb6forms_prop_ctrl.h 头上与 029 §九；
// WS-b 量的第四条 —— 事件属于哪条 socket 只能问 wParam —— 记在下面 TCP 三面那一节的头上）：
//   ① 事件种类读消息 lParam 低字，不读 WSAEnumNetworkEvents；
//   ② FD_CLOSE 会带着未读的尾巴到 ⇒ 先 drain 再发 Close；
//   ③ 还剩数据没读时 Winsock 自己会重投 FD_READ ⇒ RTL 一次读到干净，用户不调 GetData
//      也不会被反复打扰（老控件那条"DataArrival 里必须 GetData"的坑就填在这里）。

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
// winsock2.h 必须在 windows.h 之前（否则 windows.h 先把 v1 的 winsock.h 拽进来，两套
// fd_set / TIMEVAL 打架）。同仓库的先例是 src/rtl/core/di/vb6_di_net_stubs.c。
#include <winsock2.h>
#include <ws2tcpip.h>
#include <windows.h>
#endif

#include "vb6forms.h"
#include "vb6forms_internal.h"
#include "vb6rtl_variant.h"   /* vb6rtl_array.h 里要 vb6_VARIANT，先后顺序不能反 */
#include "vb6rtl_array.h"   /* vb6_SafeArray1D / Create1D / Destroy1D（ GetData 的 Byte 那一形要） */
#include <oleauto.h>   /* SysAllocString / SysFreeString / SysReAllocStringLen */
#include <string.h>
#include <wchar.h>
#include <stdlib.h>

#ifdef _WIN32

#pragma comment(lib, "ws2_32.lib")

// WSAAsyncSelect 投的私有消息。仓库里另外两条私有号是 0x7FF0（延迟 Form_Load）与
// 0x7FFF（恢复焦点），都排在 WM_APP 之外；这里用 WM_APP 段，控件自己的窗口只收自己那份通知，
// 所以**一枚号就够**（不需要 per-control 的消息号 —— 实测两枚控件各配各的窗口互不串）。
#define WM_WS_NOTIFY (WM_APP + 0x15)

#define VB6_WS_MAX           64     // 一个工程里的 Winsock 控件上限（含数组展开后的下标）
#define VB6_WS_RX_MAX        (1u << 20)   // 单枚控件接收缓冲的天花板（1 MB），超了就丢新的
#define VB6_WS_TX_MAX        (8u << 20)   // 发送队列的上限（8 MB）：装不下就报 WSAENOBUFS，不静默丢
#define VB6_WS_TX_CHUNK      (1u << 16)   // 一次 send 最多交 64 KB —— SendProgress 因此有"进度"的意义
#define VB6_WS_T_BYTEARRAY 8209  // vbByteArray = VT_ARRAY|VT_UI1(8192+17)：GetData/PeekData 的 Byte 数组那一形

typedef void (*vb6_WsCbVoid)(void);
typedef void (*vb6_WsCbLong)(int32_t);          // DataArrival / ConnectionRequest（ByVal As Long）
typedef void (*vb6_WsCbLongLong)(int32_t);   // SendProgress（ByVal bytesSent As Long）
typedef void (*vb6_WsCbInt)(int16_t);           // StateChanged（ByVal As Integer）

// 已 accept、还等用户 Accept(requestID) 来领的那条连接。监听面一次 FD_ACCEPT 挂一节。
struct vb6_WsPending {
    SOCKET                 sock;
    int32_t                reqId;
    struct vb6_WsPending*  next;
};

struct vb6_WsInstance {
    HWND     hwnd;                 // 身份窗（也是通知的目标窗）；NULL = 空槽
    SOCKET   sock;                 // 数据面（客户端连出去的、或 Accept 受理进来的那一条）
    SOCKET   listenSock;           // 服务端监听面。与 sock **分开**存 —— 这是老坑的填法之一：
                                   // `Accept(requestID)` 把待受理socket 挪进 sock，监听面照旧活着，
                                   // 所以"同一个控件自己 Accept 之后就不再收新连接"这件事不会发生。
    struct vb6_WsPending* pending; // 已 accept、等用户 Accept() 来领的队列（FIFO）
    int32_t  nextReqId;            // 交给 ConnectionRequest(requestID) 的号（只增不减，非 0）
    int32_t  protocol;             // VB6_WS_TCP / VB6_WS_UDP
    int32_t  state;                // VB6_WS_* 那一族
    int32_t  lastErr;              // 最近一次 WSA 错误码（State=sckError 时的编号）
    int32_t  localPort;
    wchar_t  localIP[64];
    wchar_t  remoteHost[256];
    int32_t  remotePort;
    int32_t  family;               // AF_INET / AF_INET6（Bind 或第一次发送时定）
    // 接收缓冲：FD_READ 里一次读干净放这儿，GetData 从这里取（②③ 两条测量的落点）
    char*    rx;
    uint32_t rxHead, rxLen;
    uint32_t rxCap;
    // 发送缓冲：SendData 交进来而内核一时装不下的那一截。有它挂着 = 监听/数据面的掩码里
    // 多一位 FD_WRITE（发完就摘掉，否则 Windows 会不停投 FD_WRITE）。
    char*    tx;
    uint32_t txHead, txLen;
    uint32_t txCap;
    int32_t  txTotal;              // 本轮 SendData 已交出去的累计（SendProgress 的增量按它算）
    int      txActive;             // 有一笔 SendData 还没发完（决定什么时候发 SendComplete）
    int32_t  bytesReceived;        // 本次事件新到的字节数（VB6 的 bytesTotal 与 BytesReceived）
    int32_t  byteTransferred;
    void*    cb[VB6_WS_EV_COUNT];
};

static struct vb6_WsInstance g_ws[VB6_WS_MAX];
static int g_wsSockStarted = 0;

static void vb6_WsEnsureWsa(void) {
    if (!g_wsSockStarted) {
        WSADATA d;
        if (WSAStartup(MAKEWORD(2, 2), &d) == 0) g_wsSockStarted = 1;
    }
}

static struct vb6_WsInstance* vb6_WsOf(HWND h) {
    int i;
    if (!h) return NULL;
    for (i = 0; i < VB6_WS_MAX; i++) if (g_ws[i].hwnd == h) return &g_ws[i];
    return NULL;
}

// 一条 socket 属于哪枚控件。WS-b 之后同一枚控件可能**同时**挂着两条 socket（监听面 +
// 数据面），二者都往同一个窗、同一个私有号投通知 ⇒ 只能问 wParam（实测 = 触发的那个
// SOCKET）。已 closesocket 的旧句柄落不进任何一条 ⇒ 回 NULL，那条通知就地丢掉。
static struct vb6_WsInstance* vb6_WsOfSocket(HWND h, SOCKET s) {
    struct vb6_WsInstance* e = vb6_WsOf(h);
    if (!e) return NULL;
    if (s != INVALID_SOCKET && (s == e->sock || s == e->listenSock)) return e;
    return NULL;
}

static void vb6_WsFreePending(struct vb6_WsInstance* e) {
    struct vb6_WsPending* p = e->pending;
    while (p) { struct vb6_WsPending* n = p->next; closesocket(p->sock); free(p); p = n; }
    e->pending = NULL;
}

static struct vb6_WsPending* vb6_WsTakePending(struct vb6_WsInstance* e, int32_t id) {
    struct vb6_WsPending** pp = &e->pending;
    while (*pp) {
        if ((*pp)->reqId == id) {
            struct vb6_WsPending* hit = *pp;
            *pp = hit->next;
            hit->next = NULL;
            return hit;
        }
        pp = &(*pp)->next;
    }
    return NULL;
}

static struct vb6_WsInstance* vb6_WsSlot(HWND h) {
    int i;
    for (i = 0; i < VB6_WS_MAX; i++) if (!g_ws[i].hwnd) {
        ZeroMemory(&g_ws[i], sizeof(g_ws[i]));
        g_ws[i].hwnd = h;
        g_ws[i].sock = INVALID_SOCKET;
        g_ws[i].listenSock = INVALID_SOCKET;
        g_ws[i].protocol = VB6_WS_TCP;      // VB6 的默认
        g_ws[i].state = VB6_WS_CLOSED;
        return &g_ws[i];
    }
    return NULL;
}

static void vb6_WsSetState(struct vb6_WsInstance* e, int32_t st) {
    if (!e || e->state == st) return;
    e->state = st;
    // 与 Timer 的 winmm 线程同一条纪律：事件只在消息循环里发，且**发完再调**，
    // 处理器里再动这枚控件（Close / SendData）不会踩到半成品状态。
    if (e->cb[VB6_WS_EV_STATECHANGED]) ((vb6_WsCbInt)e->cb[VB6_WS_EV_STATECHANGED])((int16_t)st);
}

// BSTR -> 定长 UTF-16 缓冲（NULL / 空串都当空）
static void vb6_WsCopyBstr(wchar_t* dst, size_t cap, const void* bstr) {
    const wchar_t* s = (const wchar_t*)bstr;
    size_t n;
    dst[0] = 0;
    if (!s) return;
    n = (size_t)SysStringLen((BSTR)s);
    if (n > cap - 1) n = cap - 1;
    if (n) memcpy(dst, s, n * sizeof(wchar_t));
    dst[n] = 0;
}

// 线格式 = ANSI(CP_ACP)，与 VB6 那颗 OCX 一致（String 那一形）。要传非 ASCII 的字节流
// 得用 Byte 数组那一形（WS-c）。
static void vb6_WsRxAppend(struct vb6_WsInstance* e, const char* p, int n) {
    uint32_t need;
    char* grown;
    if (n <= 0) return;
    need = e->rxLen + (uint32_t)n;
    if (need > VB6_WS_RX_MAX) return;                  // 超天花板：丢掉新的，保住已排队的那些
    if (need > e->rxCap) {
        uint32_t cap = e->rxCap ? e->rxCap : 4096u;
        while (cap < need) cap *= 2u;
        grown = (char*)realloc(e->rx, cap);
        if (!grown) return;
        e->rx = grown;
        e->rxCap = cap;
    }
    if (e->rxHead + e->rxLen > e->rxCap) {             // 环形留白不够就整体前移一次
        memmove(e->rx, e->rx + e->rxHead, e->rxLen);
        e->rxHead = 0;
    }
    memcpy(e->rx + e->rxHead + e->rxLen, p, (size_t)n);
    e->rxLen += (uint32_t)n;
}

// UDP 收到东西时把**来包地址**写回 RemoteHost / RemotePort —— 这是 VB6 那颗 OCX 的既有语义
// （"收到之后 RemoteHost 就是发件人"），回信那条路靠它才成立，不是 C3 新发明的口径。
static void vb6_WsRememberPeer(struct vb6_WsInstance* e, struct sockaddr* from, int flen) {
    wchar_t txt[80];
    char  ansi[80];
    ULONG n = sizeof(ansi);
    u_short port = 0;
    int i;
    if (WSAAddressToStringA(from, flen, NULL, ansi, &n) != 0) return;
    for (i = 0; ansi[i] && i < 79; i++) txt[i] = (wchar_t)(unsigned char)ansi[i];
    txt[i] = 0;
    port = (from->sa_family == AF_INET6) ? ntohs(((struct sockaddr_in6*)from)->sin6_port)
                                         : ntohs(((struct sockaddr_in*)from)->sin_port);
    // "127.0.0.1:54241" / "[::1]:54241" → 地址与端口分开存（LocalIP 那处同一截法）
    {
        wchar_t* colon = wcsrchr(txt, L':');
        if (colon) {
            wchar_t* lb = wcsrchr(txt, L'[');
            if (lb && lb < colon) {   // IPv6：[addr]:port —— 取括号里那截，且端口是括号之后的冒号
                wchar_t* rb = wcschr(lb, L']');
                if (rb) { *rb = 0; memmove(txt, lb + 1, (wcslen(lb + 1) + 1) * sizeof(wchar_t)); }
            } else {
                *colon = 0;
            }
        }
    }
    for (i = 0; txt[i] && i < 255; i++) e->remoteHost[i] = txt[i];
    e->remoteHost[i] = 0;
    e->remotePort = (int32_t)port;
}

// 连上之后把本地的两只读数问回来：LocalPort + LocalIP（getsockname）。没连上/没绑过时
// 保持 0 / 空串 —— 从没建过的 socket 问 getsockname 本来就是失败的（实测）。
static void vb6_WsFillLocal(struct vb6_WsInstance* e) {
    struct sockaddr_storage ss;
    int l = sizeof(ss), i;
    char txt[64];
    ULONG n = (ULONG)sizeof(txt);
    e->localPort = 0; e->localIP[0] = 0;
    if (e->sock == INVALID_SOCKET) return;
    if (getsockname(e->sock, (struct sockaddr*)&ss, &l) != 0) return;
    e->localPort = ntohs(ss.ss_family == AF_INET6 ? ((struct sockaddr_in6*)&ss)->sin6_port
                                                  : ((struct sockaddr_in*)&ss)->sin_port);
    txt[0] = 0;
    if (WSAAddressToStringA((struct sockaddr*)&ss, l, NULL, txt, &n) != 0) return;
    for (i = 0; txt[i] && i < 63; i++) e->localIP[i] = (wchar_t)(unsigned char)txt[i];
    e->localIP[i] = 0;
    // "127.0.0.1:54241" / "[::1]:54241" ⇒ 只要地址那一半，从最后一个冒号截（IPv6 的组冒号在括号里）
    while (i > 0 && e->localIP[i - 1] != L':') i--;
    if (i > 0) { e->localIP[i - 1] = 0; if (i - 1 > 0 && e->localIP[i - 2] == L']') e->localIP[i - 2] = 0; }
}

// 把 FD_READ 之后"还能读的"全部读进缓冲：一次读干净 = ③ 那条坑不再复发。
static void vb6_WsDrain(struct vb6_WsInstance* e) {
    char tmp[8192];
    int got, total = 0;
    for (;;) {
        if (e->protocol == VB6_WS_UDP) {
            struct sockaddr_storage fs;
            int fl = sizeof(fs);
            got = recvfrom(e->sock, tmp, (int)sizeof(tmp), 0, (struct sockaddr*)&fs, &fl);
            if (got > 0) vb6_WsRememberPeer(e, (struct sockaddr*)&fs, fl);
        } else {
            got = recv(e->sock, tmp, (int)sizeof(tmp), 0);
        }
        if (got == 0) { total = -1; break; }
        if (got < 0) {
            int err = WSAGetLastError();
            if (err != WSAEWOULDBLOCK) e->lastErr = err;
            break;
        }
        vb6_WsRxAppend(e, tmp, got);
        total += got;
    }
    // 每次通知只报**这一份**新到的字节。归零不是收尾好看：TCP 的 FD_CLOSE 那一格会先 drain
    // 一次（②），若上一笔的计数还挂着，就会凭空多发一次 DataArrival。
    e->bytesReceived = (total > 0) ? total : 0;
}

static void vb6_WsRaiseData(struct vb6_WsInstance* e) {
    if (!e->cb[VB6_WS_EV_DATAARRIVAL]) return;
    ((vb6_WsCbLong)e->cb[VB6_WS_EV_DATAARRIVAL])(e->bytesReceived);
}

static void vb6_WsRaiseClose(struct vb6_WsInstance* e) {
    vb6_WsSetState(e, VB6_WS_CLOSED);
    if (e->cb[VB6_WS_EV_CLOSE]) ((vb6_WsCbVoid)e->cb[VB6_WS_EV_CLOSE])();
}

static void vb6_WsRaiseConnect(struct vb6_WsInstance* e) {
    if (e->cb[VB6_WS_EV_CONNECT]) ((vb6_WsCbVoid)e->cb[VB6_WS_EV_CONNECT])();
}

static void vb6_WsRaiseConnRequest(struct vb6_WsInstance* e, int32_t id) {
    if (e->cb[VB6_WS_EV_CONNREQUEST]) ((vb6_WsCbLong)e->cb[VB6_WS_EV_CONNREQUEST])(id);
}

// 错误文本问系统要（FormatMessage 的 socket 那一段）。拿不到也要给一条非空串：
// Description 是 ByRef，处理器读到空串会以为「没有描述」。
static BSTR vb6_WsErrMsg(int err) {
    wchar_t* buf = NULL;
    DWORD n;
    BSTR out;
    n = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM
                       | FORMAT_MESSAGE_IGNORE_INSERTS,
                       NULL, (DWORD)err, 0, (wchar_t*)&buf, 0, NULL);
    if (n && buf) {
        while (n && (buf[n - 1] == 13 || buf[n - 1] == 10)) buf[--n] = 0;   /* 去掉结尾 CR LF */
        out = SysAllocStringLen(buf, n);
        LocalFree(buf);
        if (out) return out;
    }
    {
        wchar_t tmp[40];
        int i = wsprintfW(tmp, L"socket error %d", err);
        return SysAllocStringLen(tmp, i > 0 ? i : 0);
    }
}

// 出错那一格的统一出口：记下编号、翻成 sckError(9)、再发 Error 事件。
// **没有处理器也要翻 State** —— 否则 State 的读数就与「有没有写那条 Sub」耦合上了。
// 七参数的形状照发码侧生成的原型（--emit-c 实测）：
//   void f(int16_t Number, BSTR* Description, int32_t Scode, BSTR Source,
//          BSTR HelpFile, int32_t HelpContext, int16_t* CancelDisplay)
// ByRef 那两格是指针（Description / CancelDisplay），其余按值 —— 这条只有 x86 验得出来。
static void vb6_WsFail(struct vb6_WsInstance* e, int err) {
    BSTR desc, src, hf;
    int16_t cancel = 0;
    if (!e) return;
    e->lastErr = err;
    vb6_WsSetState(e, VB6_WS_ERROR);
    if (!e->cb[VB6_WS_EV_ERROR]) return;
    desc = vb6_WsErrMsg(err);
    src = SysAllocString(L"Winsock");
    hf = (BSTR)SysAllocStringLen(NULL, 0);
    if (!desc || !src || !hf) { SysFreeString(desc); SysFreeString(src); SysFreeString(hf); return; }
    ((void (*)(int16_t, void*, int32_t, void*, void*, int32_t, int16_t*))
        e->cb[VB6_WS_EV_ERROR])((int16_t)err, (void*)&desc, 0, (void*)src, (void*)hf, 0, &cancel);
    // Description 是 ByRef ⇒ 释放「回来那一枚」，不是进去那一枚（处理器可以换掉它）。
    SysFreeString(desc); SysFreeString(src); SysFreeString(hf);
}

// UDP 的"对端"：没 Connect 这一步，发往 RemoteHost:RemotePort；收到谁的算谁的（VB6 同）
static int vb6_WsUdpTarget(struct vb6_WsInstance* e, struct sockaddr_storage* ss, int* slen) {
    ZeroMemory(ss, sizeof(*ss));
    if (e->family == AF_INET6) {
        struct sockaddr_in6* a = (struct sockaddr_in6*)ss;
        a->sin6_family = AF_INET6;
        a->sin6_port = htons((u_short)e->remotePort);          /* 网络字节序 */
        if (e->remoteHost[0] == 0 || wcscmp(e->remoteHost, L"*") == 0) {
            a->sin6_addr = in6addr_any;
        } else if (!InetPtonW(AF_INET6, e->remoteHost, &a->sin6_addr)) {
            return 0;
        }
        *slen = sizeof(struct sockaddr_in6);
        return 1;
    } else {
        struct sockaddr_in* a = (struct sockaddr_in*)ss;
        a->sin_family = AF_INET;
        a->sin_port = htons((u_short)e->remotePort);          /* 网络字节序 */
        if (e->remoteHost[0] == 0 || wcscmp(e->remoteHost, L"*") == 0) {
            a->sin_addr.s_addr = INADDR_ANY;
        } else if (!InetPtonW(AF_INET, e->remoteHost, &a->sin_addr)) {
            return 0;
        }
        *slen = sizeof(struct sockaddr_in);
        return 1;
    }
}

static void vb6_WsFreeTx(struct vb6_WsInstance* e) {
    free(e->tx);
    e->tx = NULL; e->txCap = e->txHead = e->txLen = 0;
}

// 存进发送队列；容量与接收缓冲同一把尺（超了就返回 0，让调用方自己兜着）
static int vb6_WsTxAppend(struct vb6_WsInstance* e, const char* p, uint32_t n) {
    uint32_t cap;
    char* grown;
    if (!n) return 1;
    if (e->txLen + n > VB6_WS_TX_MAX) { e->lastErr = WSAENOBUFS; return 0; }
    if (e->txHead + e->txLen > e->txCap) {            // 尾部留白不够就整体前移
        memmove(e->tx, e->tx + e->txHead, e->txLen);
        e->txHead = 0;
    }
    if (e->txHead + e->txLen + n > e->txCap) {
        cap = e->txCap ? e->txCap : 4096u;
        while (cap < e->txLen + n) cap *= 2u;
        grown = (char*)realloc(e->tx, cap);
        if (!grown) { e->lastErr = WSAENOBUFS; return 0; }
        e->tx = grown; e->txCap = cap;
    }
    memcpy(e->tx + e->txHead + e->txLen, p, n);
    e->txLen += n;
    return 1;
}

static void vb6_WsTxDone(struct vb6_WsInstance* e);

// 把队列里的尽量交出去：一次最多 VB6_WS_TX_CHUNK，每交成一截就报一次 SendProgress。
// 交完 => 摘 FD_WRITE + 发 SendComplete；交不动（WSAEWOULDBLOCK）=> 留着等下一次 FD_WRITE。
// 分块不是为了快，是为了让 SendProgress 真有"进度"可报：一次 send 把 1 MB 全吞下去的机器上，
// 不分块就永远只有一条 SendProgress，"发出去多少"这件事对用户就只剩一个总数了。
static void vb6_WsTxDrain(struct vb6_WsInstance* e) {
    for (;;) {
        int want, n;
        if (!e->txLen) { vb6_WsTxDone(e); return; }
        if (e->sock == INVALID_SOCKET) { vb6_WsFreeTx(e); e->txActive = 0; return; }
        want = (int)(e->txLen > VB6_WS_TX_CHUNK ? VB6_WS_TX_CHUNK : e->txLen);
        n = send(e->sock, e->tx + e->txHead, want, 0);
        if (n > 0) {
            e->txHead += (uint32_t)n;
            e->txLen  -= (uint32_t)n;
            e->byteTransferred += n;
            if (e->cb[VB6_WS_EV_SENDPROGRESS]) ((vb6_WsCbLongLong)e->cb[VB6_WS_EV_SENDPROGRESS])(n);
            continue;
        }
        if (n == SOCKET_ERROR && WSAGetLastError() == WSAEWOULDBLOCK) return;   // 等下一次 FD_WRITE
        if (n == SOCKET_ERROR) {                                                // 真错了：报出来并丢掉队列
            int err = WSAGetLastError();          // 同一条纪律：先抓号（closesocket 会清）
            vb6_WsFreeTx(e);
            e->txActive = 0;
            vb6_WsFail(e, err);
            return;
        }
        return;                                  // n == 0：TCP 上不该有，防御性收尾
    }
}

static void vb6_WsTxDone(struct vb6_WsInstance* e) {
    if (!e->txActive) return;
    e->txActive = 0;
    vb6_WsFreeTx(e);
    if (e->sock != INVALID_SOCKET)
        WSAAsyncSelect(e->sock, e->hwnd, WM_WS_NOTIFY, FD_READ | FD_CLOSE);   // 摘掉 FD_WRITE
    if (e->cb[VB6_WS_EV_SENDCOMPLETE]) ((vb6_WsCbVoid)e->cb[VB6_WS_EV_SENDCOMPLETE])();
}

static LRESULT CALLBACK vb6_WsWndProc(HWND h, UINT m, WPARAM w, LPARAM l) {
    struct vb6_WsInstance* e;
    if (m == WM_WS_NOTIFY) {
        SOCKET s = (SOCKET)(UINT_PTR)w;      // 实测：wParam = 触发这条事件的那条 socket
        e = vb6_WsOfSocket(h, s);
        if (!e) return 0;
        switch (LOWORD(l)) {           // ① 事件种类只认这里
            case FD_READ:
                vb6_WsDrain(e);
                vb6_WsRaiseData(e);
                break;
            case FD_CLOSE:
                if (s == e->listenSock) break;   // 监听面没有"对端关了"这一说
                // ② FD_CLOSE 允许带着没读完的尾巴到（实测还剩 7 字节）⇒ 先 drain、
                //   把这一批作为一次 DataArrival 交出去，再发 Close。VB6 那颗 OCX 在这里丢数据。
                vb6_WsDrain(e);
                if (e->bytesReceived > 0) vb6_WsRaiseData(e);
                if (e->sock != INVALID_SOCKET) { closesocket(e->sock); e->sock = INVALID_SOCKET; }
                vb6_WsRaiseClose(e);
                break;
            case FD_WRITE:
                // 只有挂着发送队列时才会收到（掩码平时不带 FD_WRITE）
                if (s == e->sock) vb6_WsTxDrain(e);
                break;
            case FD_ACCEPT: {
                // 监听面来客：先把这条连接收进兜里（accept 是 Winsock 要求的，不做连接就丢了），
                // 再发 ConnectionRequest(requestID)。此刻**不动** sock / state —— 用户可能压根
                // 不 Accept，或者要 Load 一枚新的控件数组元素来 Accept（两条路都留着）。
                struct vb6_WsPending* node;
                struct vb6_WsPending** pp;
                SOCKET p = accept(e->listenSock, NULL, NULL);
                if (p == INVALID_SOCKET) { e->lastErr = WSAGetLastError(); break; }
                WSAAsyncSelect(p, e->hwnd, WM_WS_NOTIFY, 0);
                // 清完子 socket 那份**再**重登记监听面：Winsock 的口径是"FD_ACCEPT 报一次就
                // 缴械，直到 accept() 重新武装"，而 accept() 与上面那条清选择都在同一趟里，
                // 谁先把这条补回来没有保证 ⇒ 显式补一次，不赌。
                WSAAsyncSelect(e->listenSock, e->hwnd, WM_WS_NOTIFY, FD_ACCEPT);   // 清掉可能继承来的事件选择
                node = (struct vb6_WsPending*)malloc(sizeof(struct vb6_WsPending));
                if (!node) { closesocket(p); break; }
                node->sock = p; node->next = NULL;
                node->reqId = ++e->nextReqId;      // 只增不减、恒非 0（VB6 的 requestID 同形状）
                for (pp = &e->pending; *pp; pp = &(*pp)->next) { }
                *pp = node;
                vb6_WsRaiseConnRequest(e, node->reqId);
                break;
            }
            default:
                break;                          // FD_WRITE 那一档（SendComplete / SendProgress）归 WS-c
        }
        return 0;
    }
    if (m == WM_DESTROY) { vb6_Ws_Destroy((void*)h); return 0; }
    return DefWindowProcW(h, m, w, l);
}

void vb6_RegisterWinsockClass(void* hInstance) {
    static int done = 0;
    WNDCLASSW wc;
    if (done) return;
    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = vb6_WsWndProc;
    wc.hInstance = (HINSTANCE)hInstance;
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.lpszClassName = L"VB6_WINSOCK";
    if (RegisterClassW(&wc)) done = 1;
}

void* vb6_Ws_Create(void* hwnd) {
    vb6_WsEnsureWsa();
    if (hwnd) vb6_WsSlot((HWND)hwnd);
    return hwnd;
}

void vb6_Ws_Destroy(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    if (!e) return;
    vb6_WsFreePending(e);
    if (e->listenSock != INVALID_SOCKET) { closesocket(e->listenSock); e->listenSock = INVALID_SOCKET; }
    if (e->sock != INVALID_SOCKET) { closesocket(e->sock); e->sock = INVALID_SOCKET; }
    vb6_WsFreeTx(e);
    free(e->rx);
    e->rx = NULL; e->rxCap = e->rxLen = e->rxHead = 0;
    e->hwnd = NULL;                       // 槽位回收
}

void vb6_Ws_SetEventHandler(void* hwnd, int32_t kind, void* fn) {
    struct vb6_WsInstance* e;
    if (kind < 0 || kind >= VB6_WS_EV_COUNT) return;
    e = vb6_WsOf((HWND)hwnd);
    if (!e) { e = vb6_WsSlot((HWND)hwnd); }
    if (e) e->cb[kind] = fn;
}

// ---------------- 属性面 ----------------
int32_t vb6_Ws_GetProtocol(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return e ? e->protocol : VB6_WS_TCP;
}

void vb6_Ws_SetProtocol(void* hwnd, int32_t v) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    if (!e) return;
    if (e->sock != INVALID_SOCKET || e->listenSock != INVALID_SOCKET) return;   // VB6：已建套接字后改 Protocol 会报错，这里保持不动
    e->protocol = (v == VB6_WS_UDP) ? VB6_WS_UDP : VB6_WS_TCP;
}

int32_t vb6_Ws_GetState(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return e ? e->state : VB6_WS_CLOSED;
}

int32_t vb6_Ws_GetLocalPort(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return e ? e->localPort : 0;
}

const void* vb6_Ws_GetLocalIP(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return (const void*)SysAllocString(e ? e->localIP : L"");
}

const void* vb6_Ws_GetRemoteHost(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return (const void*)SysAllocString(e ? e->remoteHost : L"");
}

void vb6_Ws_SetRemoteHost(void* hwnd, const void* bstr) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    if (!e) return;
    vb6_WsCopyBstr(e->remoteHost, sizeof(e->remoteHost) / sizeof(wchar_t), bstr);
}

int32_t vb6_Ws_GetRemotePort(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return e ? e->remotePort : 0;
}

void vb6_Ws_SetRemotePort(void* hwnd, int32_t v) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    if (e) e->remotePort = v;
}

int32_t vb6_Ws_GetBytesReceived(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return e ? (int32_t)e->rxLen : 0;
}

int32_t vb6_Ws_GetByteTransferred(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    return e ? e->byteTransferred : 0;
}

// ---------------- 方法面 ----------------
// Bind(port[, ip])：UDP 的入口（TCP 也走这条定 LocalPort，服务端那一档在 WS-b 里接）。
// 实测：port=0 ⇒ 系统自动分配，从 getsockname 读回（VB6 的"LocalPort 留 0 让它自己挑"）；
// 从没绑过的 socket 问 getsockname 是**失败**的 ⇒ localPort/localIP 保持 0 / 空串。
void vb6_Ws_Bind(void* hwnd, int32_t port, const void* ip) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    SOCKET s;
    int slen, family = AF_INET, on = 1;
    wchar_t ipw[64];
    struct sockaddr_storage ss;

    if (!e) return;
    vb6_WsEnsureWsa();
    // 重新 Bind = 换一条全新的 socket：两条面都先撤（否则老的监听口还占着那个端口）
    vb6_WsFreePending(e);
    if (e->listenSock != INVALID_SOCKET) { closesocket(e->listenSock); e->listenSock = INVALID_SOCKET; }
    if (e->sock != INVALID_SOCKET) { closesocket(e->sock); e->sock = INVALID_SOCKET; }

    ZeroMemory(&ss, sizeof(ss));
    vb6_WsCopyBstr(ipw, 64, ip);
    if (ipw[0] && wcschr(ipw, L':')) family = AF_INET6;
    e->family = family;

    s = socket(family, (e->protocol == VB6_WS_UDP) ? SOCK_DGRAM : SOCK_STREAM,
               (e->protocol == VB6_WS_UDP) ? IPPROTO_UDP : IPPROTO_TCP);
    if (s == INVALID_SOCKET) { vb6_WsFail(e, WSAGetLastError()); return; }
    /* 刻意**不设** SO_REUSEADDR：设了之后两枚控件绑同一个 UDP 口，第二枚静默成功
       （Windows 的 SO_REUSEADDR 允许抢口，而包只有一份落点）—— 那是老控件没有的坑。
       不设 = 撞口当场 10048，Error 事件才报得出来；而实测「关掉再重绑同一个口」
       两种设法都成功（wsprobe14 Q3/Q4），所以这条没有快速重启的代价。 */
    (void)on;

    if (family == AF_INET6) {
        struct sockaddr_in6* a = (struct sockaddr_in6*)&ss;
        a->sin6_family = AF_INET6;
        a->sin6_port = htons((u_short)port);
        if (ipw[0] && wcscmp(ipw, L"*") && !InetPtonW(AF_INET6, ipw, &a->sin6_addr)) {
            closesocket(s); vb6_WsFail(e, WSAEADDRNOTAVAIL); return;
        }
        slen = sizeof(struct sockaddr_in6);
    } else {
        struct sockaddr_in* a = (struct sockaddr_in*)&ss;
        a->sin_family = AF_INET;
        a->sin_port = htons((u_short)port);
        if (ipw[0] && wcscmp(ipw, L"*") && !InetPtonW(AF_INET, ipw, &a->sin_addr)) {
            closesocket(s); vb6_WsFail(e, WSAEADDRNOTAVAIL); return;
        }
        slen = sizeof(struct sockaddr_in);
    }
    if (bind(s, (struct sockaddr*)&ss, slen) == SOCKET_ERROR) {
        int err = WSAGetLastError();                   // 撞已用端口 = 10048（wsprobe14 Q1）
        closesocket(s);
        vb6_WsFail(e, err);
        return;
    }
    {
        int l = slen;
        if (getsockname(s, (struct sockaddr*)&ss, &l) == 0) {
            e->localPort = ntohs(family == AF_INET6 ? ((struct sockaddr_in6*)&ss)->sin6_port
                                                    : ((struct sockaddr_in*)&ss)->sin_port);
            {
                char txt[64];
                ULONG n = (ULONG)sizeof(txt);
                txt[0] = 0;
                if (WSAAddressToStringA((struct sockaddr*)&ss, l, NULL, txt, &n) == 0) {
                    int i;
                    // WSAAddressToString 给的是 "127.0.0.1:54241" / "[::1]:54241" 这种形状，
                    // VB6 的 LocalIP 只要地址那一半 ⇒ 从最后一个冒号截断（IPv6 的组冒号都在括号里）。
                    for (i = 0; txt[i] && i < 63; i++) e->localIP[i] = (wchar_t)(unsigned char)txt[i];
                    while (i > 0 && e->localIP[i-1] != L':') i--;
                    if (i > 0) { e->localIP[i-1] = 0; if (i-1 > 0 && e->localIP[i-2] == L']') e->localIP[i-2] = 0; }
                }
            }
        }
    }
    e->sock = s;
    WSAAsyncSelect(s, e->hwnd, WM_WS_NOTIFY,
                   FD_READ | FD_CLOSE /* WS-b 再加 FD_ACCEPT / FD_CONNECT / FD_WRITE */);
    // UDP 没有"连接"一说，绑完就是 sckOpen（实测：这也是 VB6 对 UDP 的观感）
    vb6_WsSetState(e, (e->protocol == VB6_WS_UDP) ? VB6_WS_OPEN : VB6_WS_CLOSED);
}

// ---------------- TCP 三面：Listen / Accept / Connect（C29-WS-b）----------------
//
// 一条 socket 只有一张"面"：监听面（listenSock）负责收新连接，数据面（sock）负责说话。
// 老控件把这两件事塞进同一条 socket ⇒ 服务端 Accept 一条之后就再也听不了，必须开控件数组
// 绕。这里分开存，所以**同一个控件可以一边听一边跟已连上的客户往来**（多客户端仍建议
// Load 数组元素，每条连接一份状态才清楚，但不是必须）。
//
// Connect 这一侧量到的两条决定了它的形状（探针 wsprobe1/5/8）：
//   · getaddrinfo("localhost") 回**两条**：先 ::1 再 127.0.0.1 ⇒ 只试第一条会把纯 IPv4 的
//     监听端配错地址（本机那个监听端就是 IPv4 的，第一条必然被拒）⇒ 这里逐条试到通为止。
//   · 挂了 WSAAsyncSelect 的 socket 上 select() 问不出东西，所以**先试通再挂通知**：
//     非阻塞 connect + select(可写) + getsockopt(SO_ERROR) 三条合起来判成败（V4 那格实测），
//     每条候选最多等 VB6_WS_CONNECT_MS。回环上"被拒"是立刻返回的，等满的只有防火墙丢包那种。

#define VB6_WS_CONNECT_MS  750      // 单个候选地址的握手上限
#define VB6_WS_CONNECT_MAX 8        // 一条不中就换下一条，最多试这么多条

static int vb6_WsSockErrOf(SOCKET s) {
    int err = 0, len = sizeof(err);
    getsockopt(s, SOL_SOCKET, SO_ERROR, (char*)&err, &len);
    return err;
}

static void vb6_WsItow(int v, wchar_t* dst) {
    wchar_t tmp[12];
    int i = 0, j = 0;
    if (v < 0) v = 0;
    do { tmp[i++] = (wchar_t)(L'0' + v % 10); v /= 10; } while (v);
    while (i) dst[j++] = tmp[--i];
    dst[j] = 0;
}

void vb6_Ws_Listen(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    if (!e || e->protocol != VB6_WS_TCP) return;   // UDP 没有"监听"这一格
    if (e->sock == INVALID_SOCKET) return;         // 得先 Bind 定下本地端口（VB6 同：否则报错）
    if (listen(e->sock, SOMAXCONN) != 0) {
        SOCKET ls = e->sock;
        int lerr = WSAGetLastError();                  // 同上：先抓号，再关
        closesocket(ls); e->sock = INVALID_SOCKET;
        vb6_WsFail(e, lerr);
        return;
    }
    e->listenSock = e->sock;
    e->sock = INVALID_SOCKET;
    WSAAsyncSelect(e->listenSock, e->hwnd, WM_WS_NOTIFY, FD_ACCEPT);
    vb6_WsSetState(e, VB6_WS_LISTENING);
}

void vb6_Ws_Accept(void* hwnd, int32_t id) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    struct vb6_WsPending* node;
    u_long one = 1;
    if (!e) return;
    if (e->listenSock == INVALID_SOCKET) return;   // 没在听 ⇒ 无从受理
    if (e->sock != INVALID_SOCKET) return;         // 数据面已被占（VB6 在这里报错）
    node = vb6_WsTakePending(e, id);
    if (!node) { e->lastErr = WSAEINVAL; return; } // 号对不上：不是本控件发出去的那批
    e->sock = node->sock;
    free(node);
    // 握手到受理之间对手可能已经发了东西。这条 socket 默认是阻塞的，直接 recv 会**挂死**，
    // 所以先转非阻塞读干净（③ 那条：一次读到没数据为止），再把通知挂上。
    ioctlsocket(e->sock, FIONBIO, &one);
    vb6_WsDrain(e);
    one = 0;
    WSAAsyncSelect(e->sock, e->hwnd, WM_WS_NOTIFY, FD_READ | FD_CLOSE);
    ioctlsocket(e->sock, FIONBIO, &one);     // 挂在通知之后：send 走阻塞口径，一次发全
    vb6_WsSetState(e, VB6_WS_CONNECTED);
    if (e->bytesReceived > 0) vb6_WsRaiseData(e);
    // 监听面**不动**：这条控件接完客户还继续听，新的 ConnectionRequest 照发。
    // State 只有一个格子，按 VB6 的口径给数据面那半（sckConnected）⇒ "还在不在听"问 State 问不出。
    vb6_WsRaiseConnect(e);   // VB6 只在客户端那一侧明确发 Connect；这里受理后同样发一次，C3 口径
}

void vb6_Ws_Connect(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    ADDRINFOW hints, *res = NULL, *ai;
    wchar_t svc[8];
    SOCKET bound = INVALID_SOCKET;
    int keepPort = 0, tried = 0, ok = 0;
    int lastErr = 0;

    if (!e) return;
    vb6_WsEnsureWsa();
    if (e->protocol != VB6_WS_TCP) return;                 // UDP 不需要握手（直接 SendData）
    if (e->remoteHost[0] == 0) { e->lastErr = WSAEINVAL; return; }
    if (e->listenSock != INVALID_SOCKET) return;           // 正在听：要当客户端得先 Close

    // "先 Bind 再 Connect"是 VB6 定本地端口的写法 ⇒ 那条已经绑好的 socket 直接拿来握手
    // （家族对得上时）。所有权先交给局部变量，循环里统一关；它身上的通知要先摘掉 ——
    // 挂着 WSAAsyncSelect 的 socket 问不出 select()（这节头上第②条）。
    bound = e->sock;
    keepPort = (bound != INVALID_SOCKET) ? e->localPort : 0;
    if (bound != INVALID_SOCKET) {
        e->sock = INVALID_SOCKET;
        WSAAsyncSelect(bound, e->hwnd, WM_WS_NOTIFY, 0);
    }
    vb6_WsSetState(e, VB6_WS_RESOLVING_HOST);
    ZeroMemory(&hints, sizeof(hints));
    hints.ai_family = AF_UNSPEC;
    hints.ai_socktype = SOCK_STREAM;
    hints.ai_protocol = IPPROTO_TCP;
    vb6_WsItow(e->remotePort, svc);
    if (GetAddrInfoW(e->remoteHost, svc, &hints, &res) != 0) {
        vb6_WsFail(e, WSAHOST_NOT_FOUND);   // 查不到名字 = 11001（wsprobe14 Q6：rc 与 WSA 同值）
        if (bound != INVALID_SOCKET) closesocket(bound);
        return;
    }
    vb6_WsSetState(e, VB6_WS_HOST_RESOLVED);
    vb6_WsSetState(e, VB6_WS_CONNECTING);

    for (ai = res; ai && tried < VB6_WS_CONNECT_MAX; ai = ai->ai_next) {
        SOCKET s;
        struct timeval tv;
        fd_set wf;
        u_long one = 1;

        if (ai->ai_family != AF_INET && ai->ai_family != AF_INET6) continue;
        tried++;
        if (bound != INVALID_SOCKET) {
            if ((int32_t)ai->ai_family != e->family) continue;   // 绑好的那条不属于这个家族
            s = bound; bound = INVALID_SOCKET;                  // 复用：本地端口已经是它要的那个
        } else {
            s = socket(ai->ai_family, ai->ai_socktype, ai->ai_protocol);
            if (s == INVALID_SOCKET) { lastErr = WSAGetLastError(); continue; }
            if (keepPort) {
                struct sockaddr_storage any;
                ZeroMemory(&any, sizeof(any));
                if (ai->ai_family == AF_INET) {
                    struct sockaddr_in* a = (struct sockaddr_in*)&any;
                    a->sin_family = AF_INET; a->sin_port = htons((u_short)keepPort);
                } else {
                    struct sockaddr_in6* a = (struct sockaddr_in6*)&any;
                    a->sin6_family = AF_INET6; a->sin6_port = htons((u_short)keepPort);
                }
                if (bind(s, (struct sockaddr*)&any,
                         ai->ai_family == AF_INET ? sizeof(struct sockaddr_in)
                                                  : sizeof(struct sockaddr_in6)) != 0) {
                    lastErr = WSAGetLastError(); closesocket(s); continue;
                }
            }
        }
        ioctlsocket(s, FIONBIO, &one);                  // 非阻塞：connect 立刻回，成败问 SO_ERROR
        if (connect(s, ai->ai_addr, (int)ai->ai_addrlen) != 0) {
            if (WSAGetLastError() != WSAEWOULDBLOCK) {
                lastErr = WSAGetLastError(); closesocket(s); continue;
            }
        }
        FD_ZERO(&wf); FD_SET(s, &wf);
        tv.tv_sec = 0; tv.tv_usec = VB6_WS_CONNECT_MS * 1000;
        if (select(0, NULL, &wf, NULL, &tv) <= 0) { lastErr = WSAETIMEDOUT; closesocket(s); continue; }
        if ((lastErr = vb6_WsSockErrOf(s)) != 0) { closesocket(s); continue; }
        // 通了才挂通知（挂早了 select 与 async select 会打架）；数据在队列里 Winsock 自己会投
        // FD_READ（③ 那条实测），所以握手期间对手已经发来的东西不会漏。
        e->sock = s;
        e->family = (int32_t)ai->ai_family;
        WSAAsyncSelect(s, e->hwnd, WM_WS_NOTIFY, FD_READ | FD_CLOSE);
        vb6_WsFillLocal(e);
        ok = 1;
        break;
    }
    FreeAddrInfoW(res);
    if (bound != INVALID_SOCKET) closesocket(bound);   // 那条绑好的没派上用场（家族全对不上）
    if (!ok) {
        vb6_WsFail(e, lastErr);      // 连不上 = sckError + Error（不是静默回 sckClosed）
        return;
    }
    e->bytesReceived = 0; e->rxLen = e->rxHead = 0;   // 上一轮（若有）的读数不带进这一轮
    vb6_WsSetState(e, VB6_WS_CONNECTED);
    vb6_WsRaiseConnect(e);
}

// 把 tx 里剩下的尽量交出去。全交完了才摘 FD_WRITE 并发 SendComplete。
// Windows 的口径：FD_WRITE 只在"发不动了、缓冲又空出来"的时刻投，且刚挂上时会先投一次
// （所以 SendData 里那一步"挂 FD_WRITE"不会漏掉一个永远不来的通知）。

void vb6_Ws_SendData(void* hwnd, const void* data) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    const wchar_t* s = (const wchar_t*)data;
    char* bytes;
    int need, n, slen = 0;
    struct sockaddr_storage ss;

    if (!e || e->sock == INVALID_SOCKET || !s) return;
    need = WideCharToMultiByte(CP_ACP, 0, s, -1, NULL, 0, NULL, NULL);
    if (need <= 0) return;
    bytes = (char*)malloc((size_t)need);
    if (!bytes) return;
    n = WideCharToMultiByte(CP_ACP, 0, s, -1, bytes, need, NULL, NULL);
    if (n > 0) n -= 1;                       // 去掉结尾 NUL：线格式里不带它
    if (n <= 0) { free(bytes); return; }        // n 此刻是转换后的字节数（含结尾 NUL，已减掉）
    if (e->protocol == VB6_WS_UDP) {
        // UDP 不分块：报文边界本来就在这一侧，拆头发出去就是拆成好几封
        if (!vb6_WsUdpTarget(e, &ss, &slen)) { free(bytes); return; }
        n = sendto(e->sock, bytes, n, 0, (struct sockaddr*)&ss, slen);
        if (n > 0) {
            e->byteTransferred = n;
            if (e->cb[VB6_WS_EV_SENDPROGRESS]) ((vb6_WsCbLongLong)e->cb[VB6_WS_EV_SENDPROGRESS])(n);
            if (e->cb[VB6_WS_EV_SENDCOMPLETE]) ((vb6_WsCbVoid)e->cb[VB6_WS_EV_SENDCOMPLETE])();
        } else if (n == SOCKET_ERROR) {
            int err = WSAGetLastError();
            if (err == WSAENOBUFS) err = WSAENOBUFS;      // 装不下/发不动：一律报出来，不静默丢
            vb6_WsFail(e, err);
        }
        free(bytes);
        return;
    }
    /* TCP：先挂上 FD_WRITE 再入队，然后就地尽量交出去 —— 一截都交不出去时才真的留在队列里。
       老控件在这一格的两条形态（卡在阻塞 send 上 / 剩下的静默丢）就是这么填掉的：SendData
       从不等内核，交不下的部分由 FD_WRITE 继续，全交完才 SendComplete。 */
    e->byteTransferred = 0;
    e->txActive = 1;
    if (!vb6_WsTxAppend(e, bytes, (uint32_t)n)) { free(bytes); return; }   // 队列满：已在里面报过
    free(bytes);
    WSAAsyncSelect(e->sock, e->hwnd, WM_WS_NOTIFY, FD_READ | FD_CLOSE | FD_WRITE);
    vb6_WsTxDrain(e);
}

// 从缓冲里取：maxLen<=0 = VB6 的"全给我"（默认那条）
static BSTR vb6_WsTake(struct vb6_WsInstance* e, int32_t maxLen) {
    uint32_t take = e->rxLen;
    if (maxLen > 0 && (uint32_t)maxLen < take) take = (uint32_t)maxLen;
    if (take) {
        int wn = MultiByteToWideChar(CP_ACP, 0, e->rx + e->rxHead, (int)take, NULL, 0);
        BSTR out = SysAllocStringLen(NULL, wn > 0 ? wn : 0);
        if (out) {
            if (wn > 0) MultiByteToWideChar(CP_ACP, 0, e->rx + e->rxHead, (int)take, out, wn);
            e->rxHead += take;
            e->rxLen -= take;
            if (!e->rxLen) e->rxHead = 0;
        }
        return out;
    }
    return (BSTR)SysAllocStringLen(NULL, 0);   /* 空串：与 vb6_BSTR_Empty 同物，这枚 TU 没引那个头 */
}

// Byte 那一形的取数：把 [rxHead, rxHead+take) 这段**原始字节**装成一枚新的 vb6_sa_byte 数组交给调用方。
// VB6 的形态就是"控件把那个变量的数组描述符换掉"，所以旧的那枚由我们销毁（不销就是每取一次漏一块）；
// 只销 signature 对得上的 1D 把手（Fix 082g 那枚魔数就是为这种场合准备的）。
// take = 0 给的是 count=0 的空数组（UBound=-1、LBound=0 ⇒ `UBound-LBound+1` 读出 0，
// 与 VB6"没数据就是空数组"同形）。consume = 0 是 PeekData 那一半：形状照做，但缓冲的头不动。
static int32_t vb6_WsTakeBytes(struct vb6_WsInstance* e, void* out, int32_t maxLen, int consume) {
    vb6_SafeArray1D** dst = (vb6_SafeArray1D**)out;
    vb6_SafeArray1D* arr;
    uint32_t take = e->rxLen;
    if (maxLen > 0 && (uint32_t)maxLen < take) take = (uint32_t)maxLen;
    arr = vb6_SafeArrayCreate1D(vb6_sa_byte, 0, (int32_t)take - 1);
    if (!arr) return 0;
    if (take) memcpy(arr->data, e->rx + e->rxHead, take);
    if (consume) {
        e->rxHead += take;
        e->rxLen -= take;
        if (!e->rxLen) e->rxHead = 0;
    }
    if (*dst && (*dst)->signature == 0x5A1D) vb6_SafeArrayDestroy1D(*dst);
    *dst = arr;
    return (int32_t)take;
}

int32_t vb6_Ws_GetData(void* hwnd, void* out, int32_t type, int32_t maxLen) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    BSTR got;
    if (!e || !out) return 0;
    // vbByteArray(8209 = VT_ARRAY|VT_UI1) 那一形：交出去的是**线上那串字节本身**，
    // 不过码页翻译 —— 这正是它与 String 那一形的全部区别（String 那形在 vb6_WsTake 里过 ACP）。
    if (type == VB6_WS_T_BYTEARRAY) return vb6_WsTakeBytes(e, out, maxLen, 1);
    got = vb6_WsTake(e, maxLen);
    *(BSTR*)out = got;
    return (int32_t)SysStringLen(got);
}

int32_t vb6_Ws_PeekData(void* hwnd, void* out, int32_t type, int32_t maxLen) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    uint32_t head, len;
    BSTR res;
    if (!e || !out) return 0;
    if (type == VB6_WS_T_BYTEARRAY) return vb6_WsTakeBytes(e, out, maxLen, 0);
    head = e->rxHead; len = e->rxLen;
    if (maxLen > 0 && (uint32_t)maxLen < len) len = (uint32_t)maxLen;
    {
        int wn = MultiByteToWideChar(CP_ACP, 0, e->rx + head, (int)len, NULL, 0);
        res = SysAllocStringLen(NULL, wn > 0 ? wn : 0);
        if (res && wn > 0) MultiByteToWideChar(CP_ACP, 0, e->rx + head, (int)len, res, wn);
    }
    *(BSTR*)out = res;                       // **不消费**：这是 PeekData 与 GetData 唯一的差别
    return (int32_t)len;
}

void vb6_Ws_Close(void* hwnd) {
    struct vb6_WsInstance* e = vb6_WsOf((HWND)hwnd);
    if (!e) return;
    vb6_WsFreePending(e);
    vb6_WsFreeTx(e); e->txActive = 0;
    if (e->listenSock != INVALID_SOCKET) { closesocket(e->listenSock); e->listenSock = INVALID_SOCKET; }
    if (e->sock != INVALID_SOCKET) { closesocket(e->sock); e->sock = INVALID_SOCKET; }
    e->rxLen = e->rxHead = 0;
    e->bytesReceived = 0;
    vb6_WsSetState(e, VB6_WS_CLOSED);
    // VB6 的 Close **不**再触发 Close 事件（那是"对端关了"才有的）⇒ 这里只翻状态。
}

#else  /* !_WIN32 */

void vb6_RegisterWinsockClass(void* hInstance) { (void)hInstance; }

#endif /* _WIN32 */
