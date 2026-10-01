# Winsock 控件

# Winsock 控件

               

**Winsock** 控件对用户来说是不可见的，它提供了访问 TCP 和 UDP 网络服务的方便途径。Microsoft Access、Visual Basic、Visual C++ 或 Visual FoxPro 的开发人员都可使用它。为编写客户或服务器应用程序，不必了解 TCP 的细节或调用低级的 Winsock APIs。通过设置控件的属性并调用其方法就可轻易连接到一台远程机器上去，并且还可双向交换数据。

**TCP 基础**

数据传输协议允许创建和维护与远程计算机的连接。连接两台计算机就可彼此进行数据传输。

如果创建客户应用程序，就必须知道服务器计算机名或者 IP 地址（**RemoteHost** 属性），还要知道进行“侦听”的端口（**RemotePort** 属性），然后调用 **Connect** 方法。

如果创建服务器应用程序，就应设置一个收听端口（**LocalPort** 属性）并调用 **Listen** 方法。当客户计算机需要连接时就会发生 ConnectionRequest 事件。为了完成连接，可调用 ConnectionRequest 事件内的 **Accept** 方法。

建立连接后，任何一方计算机都可以收发数据。为了发送数据，可调用 **SendData** 方法。当接收数据时会发生 DataArrival 事件。调用 DataArrival 事件内的 **GetData** 方法就可获取数据。

**UDP 基础**

用户数据文报协议 (UDP) 是一个无连接协议。跟 TCP 的操作不同，计算机并不建立连接。另外 UDP 应用程序可以是客户机，也可以是服务器。

为了传输数据，首先要设置客户计算机的 **LocalPort** 属性。然后，服务器计算机只需将 **RemoteHost** 设置为客户计算机的 Internet 地址，并将 **RemotePort** 属性设置为跟客户计算机的 **LocalPort** 属性相同的端口，并调用 **SendData** 方法来着手发送信息。于是，客户计算机使用 DataArrival 事件内的 **GetData** 方法来获取已发送的信息。

---

## C3 的实现面（C29-WS-a 到 WS-f 已做：身份窗 + 状态机 + 属性面 + UDP 一整轮 + TCP 一整轮 + Error + 发送面两条事件 + 取数两个形 + 多连接的排队与串行受理）

C3 里 **Winsock 不走 `MSWINSCK.OCX`**（本页没写、但 VB6 工程发布清单里那条"把 ocx 装进 System32"
的要求在 C3 不存在）：那颗 OCX 是 32 位 inproc 服务器，64 位进程里 `CoCreateInstance` 直接失败。
C3 把它复刻成**原生 Winsock2**（`ws2_32.dll`，`winsock2.h` / `ws2tcpip.h`），语法、属性名、事件名
照 VB6，不链任何 OCX。控件本体是一枚不可见的身份窗（`VB6_WINSOCK` 类，与 `VB.Timer` 同一形状），
网络事件由 `WSAAsyncSelect` 投进它自己的窗口过程。

| 写法 | C3 里实际发生的事 |
| --- | --- |
| `W.Protocol` | 0 = sckTCPProtocol（VB6 默认）、1 = sckUDPProtocol；**已建套接字之后写不进去**（照 VB6 的报错口径，这里静默保持原值） |
| `W.LocalPort` / `W.LocalIP` | 只读。端口归 `Bind` 管、地址归系统定 —— 发一条"写得动但什么都不改"的 setter 比不发更难查 |
| `W.RemoteHost` / `W.RemotePort` | 可读可写（UDP 的目标就靠这两位） |
| `W.State` | 那十个 `sck*` 值里，UDP 绑完 = 1（sckOpen）、TCP 监听 = 2、连上/受理完 = 7、关掉 = 0 |
| `W.BytesReceived` | 缓冲里**还没取走**的字节数（`GetData` 取完就归零） |
| `W.ByteTransferred` | 本次 `SendData` **已经交给内核**的累计字节数：一笔大数据边发边长，发完就停在总长度上（不是 VB6 那句文档里模糊的"最后一次"，取数口径见下面第 8 条） |
| `W.Bind port[, ip]` | `socket` + `bind` + `getsockname` + `WSAAsyncSelect`；**`port` 写 0 就是让系统挑**，挑中了从 `LocalPort` 读回 |
| `W.Listen` | `listen` + 把这条 socket 搬进"监听面"，只挂 `FD_ACCEPT`。没 `Bind` 过时**什么都不做**（VB6 靠 `LocalPort` 属性定端口，这里那位是只读的） |
| `W.Accept requestID` | 从兜里（`FD_ACCEPT` 时已 `accept` 好的队列）领出那条连接，挂上读写通知 |
| `W.Connect` | 解析 `RemoteHost` → **逐条候选地址试到通为止**（见下面第 3 条） |
| `W.SendData s` | UDP：一次 `sendto` 发完即走；TCP：**进队列就走**，内核装不下的那些留在控件自己的发送队列里，随 `FD_WRITE` 分块泵出去，边交边报 `SendProgress`、交完报 `SendComplete`（见下面第 8 条）。线格式 = 本机 ANSI 码页 |
| `W.GetData v[, type][, maxlen]` | 从缓冲取；缺省 = 全取走。`v` 是字符串时过一道本机码页；`v()` 是 Byte 数组、`type` 写 `vbByteArray`(8209) 时给的是**线上那串字节本身**，不过码页，而且控件**重建那枚数组**（LBound 回 0，没数据就是空数组） |
| `W.PeekData v` | 同一把尺，但**不消费** —— 这是它与 `GetData` 唯一的差别 |
| `W.Close` | 关掉这条控件的两张面（监听 + 数据），兜里排队的连接一起撤 |
| `W_Bind` 撞口 / `W_Connect` 连不上 | `Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)` + `State` 落 `sckError`(9) |

四条量出来的口径，写代码时要按它们想：

1. **`DataArrival` 里不强制 `GetData`**（老控件那条"必须取，否则事件反复触发"的坑在这里填掉了）。
   RTL 在每次通知里把接收队列**一次读干净**存进自己的缓冲；而 Winsock 本身在"还剩数据没读"时
   会重投一条 `FD_READ` —— 老控件那句判据说的就是它。`bytesTotal` 给的是**本次这一份**新到的字节数，
   每次通知重新起算（上一笔的数不会带过来）。
2. **UDP 的报文边界不保证**。两笔连发（`"AB"` + `"CD"`）实测会被同一次通知里一并读走、
   `bytesTotal` 给合计（`"ABCD"`）。要保边界就得改成"读一封、发一次事件"，那是另一档事。
   ⇒ 判据一律问**内容**，别问事件条数：同一段代码在 x64 与 x86 上合并与否就不一样。
3. **`Connect` 是同步的**，与 VB6 的"发出去、等 `Connect` 事件"不同：C3 先解析
   （`getaddrinfo`，本机 `localhost` 会回**两条**：先 `::1` 再 `127.0.0.1`），再逐条候选试到通为止
   （每条最多等 750 ms），所以只对着 IPv4 监听的机器用主机名也连得上。返回时 `State` 已落定
   （7 = 通、0 = 没通），`Connect` 事件与 `StateChanged` 照发。
4. **监听面与数据面是分开存的**：`Accept` 一条连接不会把监听拆掉（老控件把两件事塞进同一条
   socket，`Accept` 之后就听不见了，只能开控件数组绕）。受理时顺手把握手期间已到的数据读干净
   ⇒ 可能立刻多发一次 `DataArrival`。而 `State` 只有一格，按 VB6 的口径给数据面那一半
   ⇒ **"还在不在听"问 `State` 问不出**。
   **而这条设计已经被实测确认**（C29-WS-f，判据 WS55..WS65）：受理过一条连接之后监听面
   照旧投递新连接的 `FD_ACCEPT`，`ConnectionRequest` 带着新的号到达；当前这条收完之后受理
   同一个新号就成功 ⇒ 一枚控件上是**排队 + 串行受理**，要**同时**服务多条才需要控件数组。

5. **`Error` 事件的 `Number` 就是 Winsock 错误码本身**（`10048` 撞口、`11001` 查不到名字、
   `10061` 被拒…），`Description` 是系统给的那句本地化文本（`FormatMessage`），`Source` 恒为
   `"Winsock"`、`Scode` / `HelpFile` / `HelpContext` 是占位。出错时 `State` 一律落到 `sckError`(9)，
   **要回 `sckClosed` 得自己 `Close`** —— 与 VB6 那侧"错误之后自己爬回 0"的传闻不同，这里不猜。
6. **绑同一个口会被当场拒**（不设 `SO_REUSEADDR`）：两枚控件 `Bind` 同一个端口，第二枚直接
   `Error 10048`。这是刻意选的：Windows 的 `SO_REUSEADDR` 语义是"允许抢口"，留着它就会
   "两枚都自称绑上了同一个口、而包只有一份收得到"那种查不出来的坑（实测见 `ai/029` §九 WS-c）。
7. **`Error` 处理器里那两枚 String 参数别直接存起来**：`gDesc = Description` 这种赋值目前发的是
   **裸指针拷贝**（不复制、不加引用），而 RTL 在处理器返回后就释放原串 ⇒ 存下来的那一格是悬垂指针。
   要留就用 `gDesc = Description & ""`（拼接才是真拷贝）。这条已另立缺陷，见 `ai/029`。

8. **发送不阻塞界面，而且不接受"发一半"这件事**：TCP 上 `SendData` 把字节丢进控件自己的
   发送队列就返回，之后每次内核腾出空位（`FD_WRITE`）交一块，**每交一块报一次 `SendProgress`**（参数 =
   这一块的长度，不是累计），**队列排空时报一次 `SendComplete`**，然后把通知恢复成只听读写。
   这里有两条工程事实：一枚已经挂了异步通知的 socket 上 `send` **允许只发一部分**
   （VB6 那个 8000 字节的内部分片所以也得有 `SendProgress`）；而 `FD_WRITE` **只在"写阻塞转
   为可写"那一下投一次**，不是池化的"现在能写"，所以队列非空期间不得摸那一位，否则剩下的内容永远敲在队里。
   队列上限 8 MiB，超了报 `Error 10055`（`WSAENOBUFS`），**不静默丢字节**。
9. **多客户端 = 一枚控件"排队 + 串行受理"**（C29-WS-f 实测）：`Accept` 一条连接不会拆掉监听面，
   新连接的 `ConnectionRequest` 照样到达（号是新的、不复用）；而**数据面只有一格** ⇒ 占着的时候
   `Accept(新号)` **被拒且不搅动状态**（`State` 照旧 7，正在服务的那条不会被玩坏），当前这条收完之后
   受理**同一个号**就成功。要同时服务多条仍按 VB6 的办法：控件数组、`Index 0` 专职 `Listen`。
   （这一条之前写的是"第二条 `FD_ACCEPT` 送不到、多客户端不可用"，那是**夹具错**不是产品错，
   三个坑叠在一起制造了假阴性：判据钉在固定拍号上、监听面被上一轮 `Close` 关掉、以及拿了撞口
   那一轮的端口去连新监听。复盘见 `ai/029` §九 C29-WS-f —— 看似"产品收不到消息"的结论，先怀疑自己的夹具。）

### 还没做的一格

`GetData` 的 Byte 数组那一形**已做**（C29-WS-e）；多客户端的排队与串行受理**已做**（C29-WS-f）。
剩：UDP 广播（`sckBroadcast`），以及**同一时刻**服务多条连接（那要控件数组，见第 9 条）。
