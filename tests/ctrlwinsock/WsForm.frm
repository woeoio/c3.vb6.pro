VERSION 5.00
Begin VB.Form WsForm 
   Caption         =   "WsForm"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "WsForm"
   ScaleHeight     =   4800
   ScaleWidth      =   9000
   Begin MSWinsockLib.Winsock wsA 
      Left            =   120
      Protocol        =   1
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsB 
      Left            =   120
      Protocol        =   1
      Top             =   720
   End
   Begin MSWinsockLib.Winsock wsC 
      Left            =   120
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock wsS 
      Left            =   120
      Top             =   1920
   End
   Begin MSWinsockLib.Winsock wsT 
      Left            =   120
      Top             =   2520
   End
   Begin MSWinsockLib.Winsock wsT2 
      Left            =   120
      Top             =   3120
   End
   Begin MSWinsockLib.Winsock wsD 
      Left            =   120
      Top             =   3720
   End
   Begin MSWinsockLib.Winsock wsE 
      Left            =   120
      Top             =   4320
   End
   Begin VB.Timer evtTimer 
      Interval        =   120
      Left            =   4680
      Top             =   120
   End
End
Attribute VB_Name = "WsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-WS-a/b：VB6 Winsock 控件走**原生 Winsock2**（不加载 MSWINSCK.OCX —— 它是 32 位
' inproc，x64 里 CoCreateInstance 直接失败，同 029 §三 D6）。WS-a = 控件身份 + 状态机 +
' 属性面 + UDP 一整轮；WS-b = TCP 的服务端与客户端一整轮（Listen / ConnectionRequest /
' Accept / Connect）；Error / SendProgress / SendComplete / Byte 数组那一形 GetData 在 WS-c。
'
' 控件分工：wsA / wsB = UDP 的一来一回（含"回信靠收来包的地址"这条 VB6 语义），
' wsC = .frm 里**什么都没写**的那枚 ⇒ 钉 Protocol / State / LocalPort 的原生默认值，
' wsS = TCP 服务端（Bind+Listen，受理两条先后到来的连接），wsT / wsT2 = 两条 TCP 客户端。
'
' 为什么一步一- tick（不是一路 DoEvents 泵到底）：网络通知本来就是"下一条消息"，
' 泵多少次全看这台机器多快 —— 第一版就是这么写的，结果 60 次 DoEvents 里对端的包还没到，
' 判据当场假红。改成 Timer 每 tick 只做一步、检查放在下一个 tick（间隔 120 ms，比回环
' 往返慢三个数量级），CI 上才是确定的。载入期间派发被 vb6_formLoading 挡着（DT-c / MV-c /
' RT-d 同先例），所以一切都在 Timer 里。
'
' 端口一律交给系统挑（Bind 0 + 读回 LocalPort）：不固定端口 ⇒ 不跟 CI 上别的作业抢号；
' 地址一律 127.0.0.1 ⇒ 不出本机、不看防火墙、不看有没有外网。

Private gArrB As Long           ' wsB 的 DataArrival 次数
Private gArrA As Long           ' wsA 的 DataArrival 次数
Private gTotB As Long           ' 最近一次 wsB 的 bytesTotal
Private gTotA As Long
Private gStB As Long            ' wsB 的 StateChanged 次数
Private gStLastB As Long        ' 最近一次 StateChanged 带回来的状态
Private gClsB As Long           ' wsB 的 Close 事件次数
Private gGotB As String         ' wsB 第一次 GetData 拿到的内容
Private gGot2 As String         ' 紧接着第二次 GetData 拿到的内容
Private gPeek As String         ' PeekData 拿到的内容
Private gGot3 As String         ' 再 GetData 拿到的内容
Private gGot4 As String         ' 两段连发拼回来的内容
Private gGot5 As String         ' 关掉之后"不该再来"的那一次取读
Private gPA As Long             ' 两枚控件各自被系统挑中的端口（跨 tick 要用，不能放局部）
Private gPB As Long
Private gPS As Long             ' TCP 服务端的监听端口（也是自动挑的）
Private gReqS As Long           ' wsS 的 ConnectionRequest 次数
Private gReq1 As Long           ' 第一条连接的 requestID
Private gReq2 As Long           ' 第二条连接的 requestID
Private gConnS As Long          ' wsS 的 Connect 事件次数（受理成功那一下）
Private gConnT As Long          ' wsT 的 Connect 事件次数（握手完成那一下）
Private gArrS As Long           ' wsS 的 DataArrival 次数
Private gTotS As Long           ' wsS 各次 bytesTotal 的合计
Private gClsS As Long           ' wsS 的 Close 事件次数
Private gProgSum As Long        ' wsT 的 SendProgress 累计字节数（**总长度**才是判据，截数不是）
Private gProgN As Long          ' wsT 的 SendProgress 次数（证人用，不当判据）
Private gDoneT As Long          ' wsT 的 SendComplete 次数
Private gProg0 As Long          ' 大 payload 那一笔开始前的 gProgSum（前面的小发送也算在里面）
Private gDone0 As Long          ' 同上，SendComplete 的起点
Private gArrE As Long           ' wsE 的 DataArrival 次数（证人用）
Private gReqE As Long           ' wsE 的 ConnectionRequest 次数
Private gReqEId As Long         ' wsE 最近一次的 requestID
Private gReqEId2 As Long        ' 第二条连接的 requestID（WS-f 那格的主证人）
Private gClsE As Long           ' wsE 的 Close 事件次数（数据面空出来的证人）
Private gReqE3 As Long           ' 第三条连接的 requestID（数据面**还被占着**的时候来的那一条）
Private gGotG As String          ' 第三条连接上取回来的内容
Private gPES As Long           ' wsE 这一轮真正在听的那个端口（gPE 是撞口那一轮的，别混）
Private gGotF As String         ' 第二条连接上拿回来的内容
Private gGotE As String         ' wsE 每次 DataArrival 当场取走的那一段
Private gGotEAll As String      ' 拼起来的全部内容与 payload 直接比（`=`，不吃 Len()）
Private gGotEBytes As Long      ' wsE 累计取到的字节数
Private gBig As String          ' WS-d 那轮的大 payload
Private k As Integer            ' 拼 gBig 的循环变量
Private gWant As Long           ' payload 的字节数（2^20，写死当尺，不用 Len 量自己（#115））
Private gBytes() As Byte        ' WS-e： GetData 的 Byte 那一形原始字节
Private gPart() As Byte         ' 带 maxLen 的那一取
Private gRest() As Byte         ' 剩下那截的再取
Private gNil() As Byte          ' 取完之后再取一次（空数组）
Private gNB As Long             ' 证人：三轮取到的元素个数
Private gWait As Long           ' 等前提成立的原地等拍计数（见 evtTimer_Timer 顶上那道闸）
Private gErrCnt As Long         ' wsD 的 Error 事件次数
Private gErrN As Long           ' 最近一次 Error 带回来的编号
Private gErrD2 As String        ' 用 `Description & ""` 复制走的描述：RTL 放掉原串之后仍要读得到
Private gErrSC As Long          ' Scode（第三参）
Private gErrHC As Long          ' HelpContext（第六参）
Private gErrCD As Integer       ' 进来时 CancelDisplay 的值
Private gErrL1 As Long          ' Len(Description)
Private gErrL2 As Long          ' Len(Source)
Private gErrL3 As Long          ' Len(HelpFile)
Private gPE As Long             ' wsE 占住的端口
Private gGotS As String         ' 服务端取回的内容
Private gGotT As String         ' 客户端取回的内容

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Form_Load()
    ' --- 1..5 默认值：什么都没写的 wsC（Protocol 0 = sckTCPProtocol、State 0 = sckClosed、
    '     没 bind 过 ⇒ LocalPort 0 / LocalIP 空串），与 .frm 里写了 Protocol=1 的两枚对照 ---
    Debug.Print "WS1=" & TF(wsC.Protocol = 0)
    Debug.Print "WS2=" & TF(wsA.Protocol = 1 And wsB.Protocol = 1)
    Debug.Print "WS3=" & TF(wsC.State = 0)
    Debug.Print "WS4=" & TF(wsC.LocalPort = 0 And wsC.LocalIP = "")
    Debug.Print "WS5=" & TF(wsC.RemotePort = 0 And wsC.RemoteHost = "")
End Sub

' ---------------- C29-WS-a: UDP 一整轮（一步一 tick） ----------------
Private Sub evtTimer_Timer()
    Static step As Integer

    ' --- 通用闸：FD_ACCEPT 落在哪一拍上没有承诺、一笔 1 MiB 什么时候发完也没有承诺，把判据钉在**固定拍号**上 = 随机红
    '     （本文件在 x64 上复跑时就红过一次：第二轮的受理晚到了一拍，后面五条一起倒）。前提没到就原地等下一拍，
    '     上限 25 拍（Timer = 120 ms ⇒ 约 3 s，九个闸全触发也不到 60 s 的运行上限）；超上限照样往下走
    Dim gateOk As Boolean
    gateOk = True
    If step = 1 Or step = 3 Then gateOk = (wsB.BytesReceived >= 7)
    If step = 4 Then gateOk = (gTotB >= 18)
    If step = 5 Then gateOk = (gTotA >= 4)
    If step = 8 Then gateOk = (gReqS >= 1)
    If step = 10 Then gateOk = (gTotS >= 5)
    If step = 11 Then gateOk = (wsT.BytesReceived >= 5)
    If step = 14 Then gateOk = (gClsS >= 1)
    If step = 20 Or step = 21 Then gateOk = (gReqE >= 1)
    If step = 25 Then gateOk = (gDoneT > gDone0 And gGotEBytes >= gWant)
    If step = 27 Then gateOk = (wsB.BytesReceived >= 5)
    If step = 28 Then gateOk = (wsB.BytesReceived >= 7)
    If step = 29 Then gateOk = (gClsE >= 1)
    If step = 30 Then gateOk = (gReqE >= 2)
    If step = 32 Then gateOk = (wsT.BytesReceived >= 9)
    If step = 33 Then gateOk = (gReqE >= 3)
    If step = 35 Then gateOk = (gClsE >= 2)
    If step = 36 Then gateOk = (wsT2.BytesReceived >= 9)
    If Not gateOk Then
        gWait = gWait + 1
        If gWait < 25 Then Exit Sub
    End If
    gWait = 0
    Dim s As String

    If step = 0 Then
        ' --- 6..8 Bind(0)：端口交给系统挑（没 bind 前 getsockname 是失败的 ⇒ 那时只能读 0），
        '      UDP 绑完的状态是 sckOpen = 1，而 StateChanged 就是那条转换的证人 ---
        wsB.Bind 0
        wsA.Bind 0
        gPB = wsB.LocalPort: gPA = wsA.LocalPort
        Debug.Print "WS6=" & TF(gPB > 0 And gPB < 65536)
        Debug.Print "WS7=" & TF(wsB.State = 1 And wsA.State = 1)
        Debug.Print "WS8=" & TF(gStB = 1 And gStLastB = 1)
        ' A → B 的第一趟
        wsA.RemoteHost = "127.0.0.1"
        wsA.RemotePort = gPB
        wsA.SendData "hello-B"
    ElseIf step = 1 Then
        ' --- 9..11 对端收到了一条 DataArrival，bytesTotal = 7 且字节真的在缓冲里 ---
        Debug.Print "WS9=" & TF(gArrB = 1 And wsB.BytesReceived = 7)
        Debug.Print "WS10=" & TF(gTotB = 7)
        Debug.Print "WS11=" & TF(wsB.LocalPort = gPB And wsA.LocalPort = gPA)
        wsB.GetData gGotB
        wsB.GetData gGot2
        Debug.Print "W0=" & gArrB & "/" & gTotB & "/" & wsB.BytesReceived
    ElseIf step = 2 Then
        ' --- 12..13 取走就消费掉；再取是空串（VB6 的 GetData 同口径，不是"读不到"而是"没了"）---
        Debug.Print "WS12=" & TF(gGotB = "hello-B" And wsB.BytesReceived = 0)
        Debug.Print "WS13=" & TF(gGot2 = "")
        wsA.SendData "peek-me"
    ElseIf step = 3 Then
        ' --- 14..15 PeekData 只看不取：BytesReceived 还在，第二次 GetData 拿到的还是那份 ---
        wsB.PeekData gPeek
        Debug.Print "WS14=" & TF(gPeek = "peek-me" And wsB.BytesReceived = 7)
        wsB.GetData gGot3
        Debug.Print "WS15=" & TF(gGot3 = "peek-me" And wsB.BytesReceived = 0)
        ' 连发两段：两笔都得拿回来。**几笔到达不算判据** —— 见下面 WS17 那段
        wsA.SendData "AB"
        wsA.SendData "CD"
    ElseIf step = 4 Then
        ' --- 16..17 VB6 那条"收到之后 RemoteHost 就是发件人"的语义：wsB 的 RemotePort 在 .frm
        '      里没写过（=0），收到 wsA 的包之后应当等于 A 自己那枚自动挑的端口 ⇒ 回信不用
        '      谁去告诉 A 它的端口是多少 ---
        Debug.Print "WS16=" & TF(wsB.RemotePort = gPA And gPA > 0)
        ' 两笔连发的判据问**内容**（"ABCD" 都在、取完就空），不问事件条数：同一对报文在
        ' x64 上落在同一次 FD_READ 里（本机 gArrB=3），在 CI 的 x86 作业上分成两次（gArrB=4）
        ' —— 合不合并只取决于那一次 drain 有没有抢到干净，是时序不是语义。第一版把 3 写成了
        ' 判据，门 #159 的 x86 那格就红在这里（证人面仍在 P= 那行报 gArrB）。
        wsB.GetData gGot4
        Debug.Print "WS17=" & TF(gGot4 = "ABCD" And wsB.BytesReceived = 0)
        wsB.RemoteHost = "127.0.0.1"
        wsB.RemotePort = gPA
        wsB.SendData "hi-A"
    ElseIf step = 5 Then
        ' --- 18..19 回信到了 A：A 的 DataArrival + A 那侧也把发件地址记了下来（同一族语义）---
        Debug.Print "WS18=" & TF(gArrA >= 1 And gTotA = 4)
        Debug.Print "WS19=" & TF(wsA.RemoteHost = "127.0.0.1" And wsA.RemotePort = gPB)
        wsA.GetData s
        Debug.Print "WS20=" & TF(s = "hi-A")
        ' --- 20..22 Close：状态回 sckClosed；之后再发不该炸、也不该有到达 ---
        wsA.Close
        Debug.Print "WS21=" & TF(wsA.State = 0)
        wsA.RemoteHost = "127.0.0.1"
        wsA.RemotePort = gPB
        wsA.SendData "after-close"
    ElseIf step = 6 Then
        ' --- 22..24 关掉之后再发：不该有东西到达 wsB（取回来还是空串），也不该有 Close 事件 ---
        wsB.GetData gGot5
        Debug.Print "WS22=" & TF(gGot5 = "" And wsB.BytesReceived = 0)
        Debug.Print "WS23=" & TF(wsA.State = 0 And gClsB = 0)
        ' --- 23..24 两条通道各自独立：wsA 的一整轮里 wsB 的 Close 事件一次都不该来；
        '      而 Close 之后 wsA 的 LocalPort 读数保持它绑过的那个（VB6：Close 不清身份）---
        Debug.Print "WS24=" & TF(wsA.LocalPort = gPA)
    ElseIf step = 7 Then
        ' ============== C29-WS-b: TCP 一整轮（服务端与客户端互为前提，一起做）=============
        ' 服务端 Bind(0) 让系统挑端口 → Listen ⇒ State = sckListening(2)
        gPS = 0
        wsS.Bind 0
        wsS.Listen
        gPS = wsS.LocalPort
        Debug.Print "WS25=" & TF(gPS > 0 And wsS.State = 2)
        ' Connect 是**同步**逐条候选试到通为止（每条最多等 VB6_WS_CONNECT_MS）⇒ 返回时状态已
        ' 落定，判据紧跟在同一步里就行，不用等下一个 tick
        wsT.RemoteHost = "127.0.0.1"
        wsT.RemotePort = gPS
        wsT.Connect
        Debug.Print "WS26=" & TF(wsT.State = 7 And gConnT = 1)
    ElseIf step = 8 Then
        ' --- 27..28 服务端的 ConnectionRequest 带着一个非 0 的号到了；没受理前状态就是监听 ---
        Debug.Print "WS27=" & TF(gReqS = 1 And gReq1 > 0)
        Debug.Print "WS28=" & TF(wsS.State = 2 And gConnS = 0)
        wsS.Accept gReq1
    ElseIf step = 9 Then
        ' --- 29 受理后这条控件是 sckConnected(7)；服务端那一侧也发一次 Connect（C3 的口径，
        '      VB6 只在客户端那侧明说过这条事件）---
        Debug.Print "WS29=" & TF(wsS.State = 7 And gConnS = 1)
        wsT.SendData "tcp-1"
    ElseIf step = 10 Then
        ' --- 30 客户端 → 服务端：内容完整拿回来（"通知里一次读干净"那条纪律在 TCP 同样成立）---
        wsS.GetData gGotS
        Debug.Print "WS30=" & TF(gGotS = "tcp-1" And gArrS = 1 And gTotS = 5)
        wsS.SendData "tcp-2"        ' --- 31 反方向：受理之后的这条控件也能说话 ---
    ElseIf step = 11 Then
        wsT.GetData gGotT
        Debug.Print "WS31=" & TF(gGotT = "tcp-2")
        Debug.Print "WS32=" & TF(wsS.LocalPort = gPS And wsT.LocalPort > 0)
    ElseIf step = 12 Then
        ' --- 33 客户端主动 Close：它自己回 sckClosed。Close 不催自己的 Close 事件（WS-a 同口径）---
        wsT.Close
        Debug.Print "WS33=" & TF(wsT.State = 0)
    ElseIf step = 13 Then
        ' 这一 tick 什么都不判：FD_CLOSE 是"对端关掉之后下一条消息"，给它两个 tick 的余量
        ' （一个 tick = 120 ms，回环上的 FIN 早就到了；CI 上被抢占时这半格就是余量）
        Debug.Print "W=" & gClsS & "/" & wsS.State
    ElseIf step = 14 Then
        ' --- 35 第二条连接，目标故意写成**主机名**：本机 "localhost" 解析出两条（先 ::1 再
        '      127.0.0.1，探针 wsprobe1），而这里的监听端是纯 IPv4 的 ⇒ 只试第一条必然连不上。
        '      这一格连得上，就是"逐条候选试到通"那条的证人 ---
        wsT2.RemoteHost = "localhost"
        wsT2.RemotePort = gPS
        wsT2.Connect
        Debug.Print "WS35=" & TF(wsT2.State = 7)
        ' --- 34 服务端的**数据面**收到过 FD_CLOSE ⇒ Close 事件 + 状态回 0（"关得干净"的证人）---
        Debug.Print "WS34=" & TF(gClsS >= 1 And wsS.State = 0)
    ElseIf step = 15 Then
        ' ======================= C29-WS-c: Error 事件 =======================
        ' 撞口：wsE 先占一个口，wsD 再绑同一个口 ⇒ 当场 10048。
        ' 这条同时也是"不设 SO_REUSEADDR"的证人（探针 wsprobe14 Q1/Q2：设了的话第二枚**静默成功**）
        wsE.Bind 0
        gPE = wsE.LocalPort
        wsD.Bind gPE
        Debug.Print "WS36=" & TF(gErrCnt = 1 And gErrN = 10048)
        Debug.Print "WS37=" & TF(wsD.State = 9)          ' sckError
    ElseIf step = 16 Then
        ' 查不到名字：保留域 .invalid ⇒ 11001（wsprobe14 Q6：rc 与 WSA 同值）
        wsD.RemoteHost = "no-such-host.invalid"
        wsD.RemotePort = 80
        wsD.Connect
        Debug.Print "WS38=" & TF(gErrCnt = 2 And gErrN = 11001)
        ' 七参数里那两格 String：ByVal 的 Source 与 ByRef 的 Description 都要真拿到东西
        ' （Description 空串 = RTL 那侧 FormatMessage 没接上；Source 空 = ByVal BSTR 传坏了）
        Debug.Print "WS39=" & TF(gErrL2 = 7 And gErrL1 > 0 And gErrL3 = 0)
        ' 参数**当场**是好的：Source 的长度就是 "Winsock" 那 7 个字符、Description 非空、HelpFile 空。
        ' 但**别把这两枚 String 参数存进模块变量** —— 形参赋值发的是指针拷贝（不复制、不加引用），
        ' RTL 在调用返回后就放掉原串 ⇒ 存下来的那一格是悬垂指针（本批实测：读出来是别人用过的
        ' 堆块，比如 "WS38="）。要留就得 & "" 复制走，WS40 钉的正是这条。
        Debug.Print "W=" & gErrCnt & "/" & gErrN & "/" & gErrSC & "/" & gErrHC & "/" & gErrCD
    ElseIf step = 17 Then
        Debug.Print "WS40=" & TF(gErrD2 <> "" And wsD.State = 9)
        ' 出错之后用户显式 Close 就回 sckClosed（State 不会自己从 9 爬回去）
        wsD.Close
        wsE.Close
        Debug.Print "WS41=" & TF(wsD.State = 0)
        ' 第二条连接的 ConnectionRequest 这一格**量出来收不到**：探针里同样的顺序
        ' （accept 完第一条、把它读完再关，然后第二条进来）FD_ACCEPT 会再投一次，
        ' 产品里同一枚控件的窗口从此不再收到任何 FD_ACCEPT（裸码 trace 记在 029 §九
        ' C29-WS-b 那一格）⇒ 这一档只把"一条连接 + 名字逐条试连"钉住，多客户端的
        ' 受理路归 WS-c（control array 每元素一条监听 / 或换 WSAEventSelect 轮询）。
        Debug.Print "W=" & gReqS & "/" & gReq1 & "/" & gClsS & "/" & wsS.State
        Debug.Print "P=" & gPA & "/" & gPB & "/" & gPS & "/" & gStB & "/" & gArrA & "/" & gArrB
        Debug.Print "P2=" & gConnT & "/" & gConnS & "/" & gArrS & "/" & gTotS & "/" & gReq2
    ElseIf step = 18 Then
        ' ===================== C29-WS-d: 一笔大过内核发送缓冲的 SendData =====================
        ' 老控件在这一格的形态是"界面卡死在阻塞 send 上"或"剩下的静默丢"。C3 的填法：一次交不完
        ' 就进这条控件自己的发送队列 + 挂 FD_WRITE，边交边报 SendProgress，全交完报 SendComplete。
        wsE.Bind 0
        wsE.Listen
        gPES = wsE.LocalPort
        Debug.Print "WS42=" & TF(wsE.State = 2 And wsE.LocalPort > 0)
    ElseIf step = 19 Then
        gBig = "x"
        For k = 1 To 20
            gBig = gBig & gBig              ' 2^20 个字符，全 ASCII ⇒ 字节数 = 字符数 = 1048576
        Next k
        gWant = 1048576                     ' 刻意大到内核两端缓冲装不下 ⇒ 真走发送队列
        wsT.RemoteHost = "127.0.0.1"
        wsT.RemotePort = wsE.LocalPort
        wsT.Connect
        Debug.Print "WS43=" & TF(wsT.State = 7 And gConnT = 2)
    ElseIf step = 20 Then
        ' 走到这一步时上面那道闸已经保证 FD_ACCEPT 到了（没到就一直等，最多 60 拍）。
        If gReqE >= 1 Then wsE.Accept gReqEId
    ElseIf step = 21 Then
        If gReqE >= 1 And wsE.State <> 7 Then wsE.Accept gReqEId
        Debug.Print "WS44=" & TF(gReqE = 1 And gReqEId > 0)
        Debug.Print "WS45=" & TF(wsE.State = 7 And gArrE = 0)
    ElseIf step = 22 Then
        gProg0 = gProgSum                   ' 前几轮的小发送也算过账，判据一律问增量
        gDone0 = gDoneT
        wsT.SendData gBig                   ' 交出去就走，不在这里等内核
    ElseIf step = 25 Then
        ' 三条判据全部问**总量**：SendProgress 的字节合计 = payload 长度（分了几截是内核缓冲
        ' 大小的函数，不算承诺，与 WS-a 那格"两笔 UDP 合并成一次 DataArrival"同一条教训）；
        ' SendComplete 一笔一次；接收侧按事件参数 bytesTotal 累计，内容与 payload 直接比。
        Debug.Print "WS46=" & TF(gProgSum - gProg0 = gWant And wsT.ByteTransferred = gWant)
        Debug.Print "WS47=" & TF(gDoneT - gDone0 = 1 And gGotEBytes = gWant)
        Debug.Print "WS48=" & TF(gGotEAll = gBig)
        wsT.Close                          ' 只关客户端那头；wsE 的监听面留给下面 WS-f 那一轮
        Debug.Print "W=" & gProgSum & "/" & gProgN & "/" & gDoneT & "/" & gGotEBytes
        Debug.Print "P3=" & gReqE & "/" & gReqEId & "/" & gWant
        ' 本轮只把"发送侧两条事件"钉下；取数的 Byte 那一形在下面的 WS-e 那轮
    ElseIf step = 26 Then
        ' ======================= C29-WS-e: GetData 的 Byte 数组那一形 =======================
        ' 这一形的区别不是“取得动作不同”，而是**取出来的东西不同**：String 那形过一道本机码页，
        ' Byte 那形给的是线上那串字节本身，而且 VB6 是**控件重建那个数组**（LBound 回 0）。
        ' 载荷一律纯 ASCII（任何码页下都是逐字节相等），否则字节数随 CI 的 ACP 变（同账 #79 那族）。
        wsA.Bind 0
        wsA.RemoteHost = "127.0.0.1"
        wsA.RemotePort = gPB
        wsA.SendData "abcde"
    ElseIf step = 27 Then
        wsB.GetData gBytes, vbByteArray
        wsB.GetData gNil, vbByteArray
        Debug.Print "WS49=" & TF(UBound(gBytes) - LBound(gBytes) + 1 = 5 And LBound(gBytes) = 0)
        Debug.Print "WS50=" & TF(gBytes(0) = 97 And gBytes(4) = 101)
        Debug.Print "WS51=" & TF(UBound(gNil) - LBound(gNil) + 1 = 0 And wsB.BytesReceived = 0)
        wsA.SendData "abcdefg"
    ElseIf step = 28 Then
        ' 带 maxLen 的那一取：VB6 只给前三个字节，剩下四个必须还在缓冲里（不丢不重）
        wsB.GetData gPart, vbByteArray, 3
        Debug.Print "WS52=" & TF(UBound(gPart) - LBound(gPart) + 1 = 3 And gPart(2) = 99)
        Debug.Print "WS53=" & TF(wsB.BytesReceived = 4)
        wsB.GetData gRest, vbByteArray
        Debug.Print "WS54=" & TF(UBound(gRest) - LBound(gRest) + 1 = 4 And gRest(0) = 100 And gRest(3) = 103)
        wsA.Close
        wsB.Close
        gNB = (UBound(gBytes) - LBound(gBytes) + 1) * 100 + (UBound(gPart) - LBound(gPart) + 1) * 10              + (UBound(gRest) - LBound(gRest) + 1)
        Debug.Print "W1=" & gNB & "/" & gArrB
        ' 第一条连接已在上面关掉，这里只把 UDP 那一轮收尾
    ElseIf step = 29 Then
        ' ============= C29-WS-f：第二条连接（旧账"受理过一条之后 FD_ACCEPT 不再来"的定性重量）=============
        ' 之前的结论是把判据写在固定拍号上量出来的，而 WS-d 那一跑证明了"到达的拍号也是时序"
        ' （门 #165 红的就是这条）⇒ 现在用闸等到事件真到再问。这一格若是绿的，多客户端就从
        ' "不可用"升级成"同一枚控件串行可用"。
        Debug.Print "WS55=" & TF(gClsE >= 1 And wsT.State = 0)
        wsT.RemoteHost = "127.0.0.1"
        wsT.RemotePort = gPES
        wsT.Connect
        Debug.Print "W9=" & wsT.State & "/" & gConnT & "/" & gPES & "/" & wsE.LocalPort
        Debug.Print "WS56=" & TF(wsT.State = 7 And gConnT = 3)
    ElseIf step = 30 Then
        Debug.Print "WS57=" & TF(gReqE = 2 And gReqEId2 > 0 And gReqEId2 <> gReqEId)
        Debug.Print "W2=" & gReqE & "/" & gReqEId & "/" & gReqEId2 & "/" & gClsE & "/" & wsE.State
        If gReqEId2 > 0 Then wsE.Accept gReqEId2
    ElseIf step = 31 Then
        Debug.Print "WS58=" & TF(wsE.State = 7)
        wsE.SendData "second-ok"
    ElseIf step = 32 Then
        wsT.GetData gGotF
        Debug.Print "WS59=" & TF(gGotF = "second-ok")
        ' 关键的一问：**数据面还被第二条占着**的时候，第三条连接的请求到不到？（VB6 要控件数组
        ' 绕的就是这个，而 C3 的监听面与数据面是分开的 ⇒ 监听根本没停过）
        wsT2.Close                          ' WS-b 那一轮它还连着：先收干净，否则 Connect 被拒是**对的**
        wsT2.RemoteHost = "127.0.0.1"
        wsT2.RemotePort = gPES
        wsT2.Connect
        Debug.Print "W5=" & wsT2.State & "/" & wsT2.LocalPort & "/" & gPES & "/" & wsE.LocalPort & "/" & gErrN & "/" & gErrCnt
        Debug.Print "WS60=" & TF(wsT2.State = 7)
    ElseIf step = 33 Then
        Debug.Print "WS61=" & TF(gReqE = 3 And gReqE3 > 0 And gReqE3 <> gReqEId2)
        Debug.Print "WS62=" & TF(wsE.State = 7)
        wsE.Accept gReqE3
        ' 占着的时候受理：按 VB6 的口径应当被拒、并且不搅动状态
        Debug.Print "WS63=" & TF(wsE.State = 7)
    ElseIf step = 34 Then
        wsT.Close                          ' 收掉第二条，把数据面空出来
    ElseIf step = 35 Then
        wsE.Accept gReqE3                  ' 空出来之后受理同一条：这才是"排队不丢"的证人
        Debug.Print "WS64=" & TF(wsE.State = 7)
        wsE.SendData "third-ok"
    ElseIf step = 36 Then
        wsT2.GetData gGotG
        Debug.Print "WS65=" & TF(gGotG = "third-ok")
        Debug.Print "W3=" & gReqE & "/" & gReqEId2 & "/" & gReqE3 & "/" & gClsE & "/" & wsE.State
        wsE.Close
        wsT.Close
        wsT2.Close
        Debug.Print "W4=" & gArrE & "/" & gDoneT & "/" & gConnT
        Debug.Print "CTRLWINSOCK-DONE"
        Unload Me
    End If
    step = step + 1
End Sub

Private Sub wsA_DataArrival(ByVal bytesTotal As Long)
    gArrA = gArrA + 1
    gTotA = gTotA + bytesTotal
End Sub

Private Sub wsB_DataArrival(ByVal bytesTotal As Long)
    gArrB = gArrB + 1
    gTotB = gTotB + bytesTotal
End Sub

Private Sub wsB_StateChanged(ByVal State As Integer)
    gStB = gStB + 1
    gStLastB = State
End Sub

Private Sub wsB_Close()
    gClsB = gClsB + 1
End Sub

' ---------------- C29-WS-b: TCP 那一轮的事件面 ----------------
Private Sub wsS_ConnectionRequest(ByVal requestID As Long)
    gReqS = gReqS + 1
    If gReq1 = 0 Then
        gReq1 = requestID
    Else
        gReq2 = requestID
    End If
End Sub

Private Sub wsS_DataArrival(ByVal bytesTotal As Long)
    gArrS = gArrS + 1
    gTotS = gTotS + bytesTotal
End Sub

Private Sub wsS_Close()
    gClsS = gClsS + 1
End Sub

Private Sub wsS_Connect()
    gConnS = gConnS + 1
End Sub

Private Sub wsD_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, _
                      ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, _
                      CancelDisplay As Boolean)
    gErrCnt = gErrCnt + 1
    gErrN = Number
    gErrD2 = Description & ""
    gErrL1 = Len(Description)
    gErrL2 = Len(Source)
    gErrL3 = Len(HelpFile)
    gErrSC = Scode
    gErrHC = HelpContext
    gErrCD = CancelDisplay
    CancelDisplay = True
End Sub

' ---------------- C29-WS-d：发送侧两条事件 + 第二对服务端的三面 ----------------
Private Sub wsT_SendProgress(ByVal bytesSent As Long)
    gProgSum = gProgSum + bytesSent
    gProgN = gProgN + 1
End Sub

Private Sub wsT_SendComplete()
    gDoneT = gDoneT + 1
End Sub

Private Sub wsE_ConnectionRequest(ByVal requestID As Long)
    gReqE = gReqE + 1
    If gReqE = 1 Then
        gReqEId = requestID
    ElseIf gReqE = 2 Then
        gReqEId2 = requestID
    Else
        gReqE3 = requestID
    End If
End Sub

Private Sub wsE_Close()
    gClsE = gClsE + 1
End Sub

Private Sub wsE_DataArrival(ByVal bytesTotal As Long)
    gArrE = gArrE + 1
    wsE.GetData gGotE                     ' 当场取走：接收侧也一次读干净，窗口不会堆爆
    gGotEAll = gGotEAll & gGotE
    gGotEBytes = gGotEBytes + bytesTotal  ' 计数吃事件参数，不用 Len(模块级 String)（#115）
End Sub

Private Sub wsT_Connect()
    gConnT = gConnT + 1
End Sub
