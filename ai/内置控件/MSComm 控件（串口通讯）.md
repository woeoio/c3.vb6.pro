# MSComm 控件（串口通讯）

> 部件：Microsoft Comm Control 6.0 (SP6)，MSCOMM32.OCX

## 1. 概述
COM 串口通讯控件：PLC、单片机、条码枪、读卡器、电子秤等上位机程序的扛把子。
**一个 MSComm 实例对应一个串口**，多串口就放多个控件（或控件数组）。

## 2. 主要属性

| 属性 | 说明 |
|---|---|
| `CommPort` | 串口号 1~16（COM1、COM2…） |
| `Settings` | "波特率,校验,数据位,停止位"，如 "9600,N,8,1" |
| `PortOpen` | True 打开 / False 关闭串口 |
| `Input` | 读接收缓冲区（Variant；读后**自动清空**） |
| `Output` | 发送（字符串或 Byte 数组） |
| `InputMode` | 0 - comInputModeText 文本、1 - comInputModeBinary 二进制（Hex 协议必用） |
| `InputLen` | 每次 Input 读取的字符数，0 = 一次读全部 |
| `InBufferSize / OutBufferSize` | 收/发缓冲区大小（默认 1024/512） |
| `InBufferCount / OutBufferCount` | 缓冲区当前字节数；**赋 0 = 清空缓冲区** |
| `RThreshold` | 收到多少字符触发一次 OnComm(comEvReceive)；0 = 不触发（只能轮询） |
| `SThreshold` | 发送缓冲降到该值以下触发 comEvSend |
| `DTREnable / RTSEnable` | DTR / RTS 握手线控制 |
| `Handshaking` | 0 - comNone、1 - comXOnXOff 软件、2 - comRTS 硬件、3 - 两者 |
| `CTSHolding / DSRHolding / CDHolding` | （只读）握手线状态 |

## 3. 事件：只有 OnComm 一个
用 `CommEvent` 区分原因：
- 正常：comEvSend=1、**comEvReceive=2**、comEvCTS=3、comEvDSR=4、comEvCD=5、comEvRing=6、comEvEOF=7
- 错误：comEventFrame=1004 帧错误、comEventOverrun=1006 超速、**comEventRxOver=1008 缓冲溢出**、**comEventRxParity=1009 校验错**、comEventTxFull=1010 发送缓冲满

## 4. 示例（二进制收发骨架）

```vb
Private Sub Form_Load()
    With MSComm1
        .CommPort = 1
        .Settings = "9600,N,8,1"
        .InputMode = comInputModeBinary
        .InputLen = 0            ' 一次读光
        .RThreshold = 1          ' 每收到1字节触发OnComm
        .InBufferCount = 0
        .PortOpen = True
    End With
End Sub

Private Sub MSComm1_OnComm()
    Dim v As Variant, b() As Byte, i As Long, s As String
    If MSComm1.CommEvent = comEvReceive Then
        v = MSComm1.Input: b = v
        For i = LBound(b) To UBound(b)
            s = s & Right("0" & Hex(b(i)), 2) & " "     ' 拼成 Hex 串显示
        Next
        txtRecv.Text = txtRecv.Text & s
    ElseIf MSComm1.CommEvent >= 1000 Then
        MsgBox "串口错误：" & MSComm1.CommEvent
    End If
End Sub

' 发送：文本 / 二进制
MSComm1.Output = "AT" & vbCrLf
Dim b(2) As Byte: b(0) = &HAA: b(1) = &H1: b(2) = &HFF
MSComm1.Output = b

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen Then MSComm1.PortOpen = False   ' 务必关端口
End Sub
```

> 坑：① RThreshold=0 永远不收事件；② 二进制协议用 Text 模式必乱码；③ 设备回包可能**粘包/拆包**，要按协议（定长/帧头帧尾/校验和）自己组包；④ 端口被占（已打开或他程序占用）时 PortOpen=True 会报 8005 错误。