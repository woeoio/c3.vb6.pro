# Winsock 控件（网络通讯）

> 部件：Microsoft WinSock Control 6.0 (SP6)，MSWINSCK.OCX

## 1. 概述
TCP / UDP 网络通讯控件：做聊天工具、联机程序、设备联网、自建协议服务端。
TCP 面向连接（服务器 Listen+Accept，客户端 Connect）；UDP 无连接直接收发。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `Protocol` | 0 - sckTCPProtocol（默认）、1 - sckUDPProtocol |
| `LocalPort / LocalIP` | 本地端口 / （只读）本机 IP |
| `RemoteHost / RemotePort` | 目标主机（IP 或域名）/ 目标端口 |
| `State` | （只读）状态机：0 sckClosed、2 sckListening、6 sckConnecting、7 sckConnected、8 sckClosing、9 sckError 等 |

## 3. 方法

| 方法 | 说明 |
|---|---|
| `Bind(localPort, [localIP])` | 绑定本地端口（UDP 必用；TCP 服务器可选） |
| `Listen` | TCP 服务器开始监听 |
| `Accept(requestID)` | **在 ConnectionRequest 事件里调用**，接受连接 |
| `Connect([remoteHost], [remotePort])` | TCP 客户端发起连接 |
| `SendData(data)` | 发送（字符串 / Byte 数组 / 数值） |
| `GetData(data, [type], [maxLen])` | **在 DataArrival 事件里调用**取数据，取后清缓冲 |
| `PeekData` | 窥视数据但不清缓冲 |
| `Close` | 关闭连接 / 停止监听 |

## 4. 事件

| 事件 | 说明 |
|---|---|
| `ConnectionRequest(requestID As Long)` | 服务器收到连接请求 → 调 Accept |
| `Connect` | 客户端连接建立成功 |
| `DataArrival(bytesTotal As Long)` | **核心**：数据到达 → 调 GetData |
| `Close` | 对方关闭连接 |
| `SendComplete / SendProgress` | 发送完成 / 发送进度 |
| `Error(Number, Description, …)` | 网络错误（拒绝连接、复位等） |
| `StateChanged(state As Integer)` | 状态机变化（调试/显示连接状态用） |

## 5. 示例（TCP 一问一答最小系统）

**服务器端：**
```vb
Private Sub Form_Load()
    Winsock1.LocalPort = 8001
    Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    If Winsock1.State <> sckClosed Then Winsock1.Close
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim s As String
    Winsock1.GetData s, vbString
    txtLog.Text = txtLog.Text & "客户端：" & s & vbCrLf
    Winsock1.SendData "服务器已收到：" & s
End Sub
```

**客户端：**
```vb
Winsock1.RemoteHost = "192.168.1.10": Winsock1.RemotePort = 8001
Winsock1.Connect
' 连接成功后（Connect 事件或 State=7）：
Winsock1.SendData "你好，服务器！"
' 接收同样写在 DataArrival 里
```

> 坑：① DataArrival 里**必须** GetData，否则缓冲不清、事件反复触发；② 多客户端服务器要用 Winsock 控件数组：index 0 专职 Listen，ConnectionRequest 时 Load 新下标 Accept 交给它；③ UDP 无 Connect，Bind 本地端口 + 设 RemoteHost/RemotePort 后直接 SendData；④ 大文件/粘包问题同串口：应用层自定义长度头或分隔符组包；⑤ 跨机器测试先关防火墙或加例外。