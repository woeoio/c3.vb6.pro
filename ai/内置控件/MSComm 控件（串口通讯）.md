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
## 5. C3 的实现与口径（C30-4，2026-11-15）

落点：**Win32 串口 API 直连** —— `CreateFileW(L"\\\\.\\COM<n>")` + `DCB`/`SetCommState` +
`ReadFile` 读线程，RTL 文件 `src/rtl/core/vb6forms/vb6forms_mscomm.c`；不加载 `MSCOMM32.OCX`。

### 5.1 为什么是“每端口一条读线程”

`WaitCommEvent` 的 `EV_RXCHAR` 只报“有字符到了”一次，凑 `RThreshold` 还得自己数；而读线程
`ReadFile` 阻塞等数据，天然就是“到多少给多少”。读线程里把“有数据了”投成一条私有消息
（`WM_APP+0x16`）到控件自己的身份窗（类名 `VB6_MSCOMM`，RTL 自注册），窗口过程里再判
`RThreshold` 去调 `OnComm` —— 与 Winsock 同一个形状。

### 5.2 一处被废掉的设计（实测栽过）

第一版用 `FILE_FLAG_OVERLAPPED` 打开 + `OVERLAPPED` 结构读写。夹具在 `\\.\NUL` 上的读数是
`MC3-tx=0`（写没出去）与 `MC5-ev=1008`（读报错）：**overlapped 句柄上不带 `OVERLAPPED` 结构的
读写一律 `ERROR_INVALID_PARAMETER`**。改成同步读写后：

- 写：直接 `WriteFile`（同步句柄）；
- 读：读线程同步 `ReadFile`，`COMMTIMEOUTS.ReadIntervalTimeout=50` —— “新字节之间空 50ms 就返回”，
  所以线程会自然醒来一轮，不是死等；
- 关：`CancelSynchronousIo(读线程句柄)` 把它从 `ReadFile` 里叫醒（`CloseHandle` 不会取消
  正在进行的同步 I/O）。

### 5.3 属性面

`CommPort` / `Settings` / `PortOpen` / `Input` / `Output` / `InputMode` / `InputLen` /
`InBufferCount`（赋 0 = 清空）/ `OutBufferCount` / `InBufferSize` / `OutBufferSize` /
`RThreshold` / `SThreshold` / `Handshaking` / `DTREnable` / `RTSEnable` / `EOFEnable` /
`CommEvent` / `CTSHolding` / `DSRHolding` / `CDHolding`，事件只有 `OnComm`。
`CommEvent` 的号严格照 §3（正常 1–7、错误 1004/1006/1008/1009/1010）；`PortOpen` 打开失败时
保持 `False`（VB6 那侧会抛 8005，C3 没有可捕获错误面）。

### 5.4 实跑读数（`tests/ctrlmscomm`，x64 与 x86 逐字一致，`exit=0`）

本机没有串口硬件，判据用 C3 的**测试钩子** `C3_MSCOMM_DEVICE=\\.\NUL`（只给夹具用，不是 VB6 API）：

```
MC1-prop=1|9600,N,8,1|0   CommPort / Settings / InputLen 回读
MC2-open=True             PortOpen=True 真开出句柄
MC3-tx=2                  Output="AT" 真写出 2 字节（TxTotal 是 C3 证人）
MC4-in=0|0                InBufferCount=0、Input 空串
MC5-ev=7                  NUL 读到 0 字节 ⇒ comEvEOF(7)
MC6-cnt=True              OnComm 至少到过一次
MC7-close=False           关端口
MC8-reopen=True           关掉再开（句柄/线程没漏）
MC9-tx2=5                 累计计数器：2 + 3
```

### 5.5 未覆盖（诚实口径）

1. **二进制收发未做**：`InputMode` 只存与回读，`Input`/`Output` 一律按文本（ACP）走。
   不上报“每字节一个字符”的假数组 —— 那种数据用户拿去 `Hex()` 会静默错。
2. **握手线变化事件未做**：`comEvCTS/DSR/CD`（3/4/5）要 `WaitCommEvent` 的 `EV_CTS` 档，与
   5.1 的选择冲突；三条 `*Holding` 属性照问不误。
3. **`comEvRing`（6）/ `comEvSend`（1）**：前者要有振铃的设备，后者按“同步写完即缓冲空”报一次，
   与 VB6 “降到 `SThreshold` 以下才报”的时机不是同一刻（差一次写）。
4. **端口被占的 8005 错误**：C3 只表现为 `PortOpen=False`，没有可捕获的错误号。

## 6. C30-4b：把“现代设备”那几件补上（2026-11-15）

§5 那一版能跑，但对着你列的清单还缺五项。这一批全补，并且**把“静默失败”改成“当场失败”**。

### 6.1 Settings 自己解析（不再交给 BuildCommDCBW）

`9600,N,8,1` 这一串现在由 RTL 自己拆：波特率 `50..12000000`、校验 `N/E/O/M/S`、数据位 `5..8`、
停止位 `1/1.5/2`，有一项不合法就**打开失败**（`PortOpen=False`）。
旧实现把它丢给 `BuildCommDCBW`，而它把解析失败吞成 FALSE —— 于是一次写错的高波特率会让 DCB
停在 `GetCommState` 拿到的旧值上，**`PortOpen` 报 True、实际按 9600 跑**。这个静默错现在没有了。

两种“不认 DCB”的情形也分开了：`GetCommState` 失败（`\\.\NUL`、虚拟对的一半）⇒ 降级档，读写照走；
`GetCommState` 成功但 `Settings` 解析或 `SetCommState` 失败 ⇒ 硬失败（关句柄）。

### 6.2 超时（C3 扩展，VB6 没有这三条）

| 属性 | 默认 | 对应 COMMTIMEOUTS |
|---|---|---|
| `ReadIntervalTimeout` | 50 ms | `ReadIntervalTimeout` —— **帧间隙**：两个字节之间空这么久就算一帧结束 |
| `ReadTotalTimeout` | 0 | `ReadTotalTimeoutConstant` |
| `WriteTimeout` | 2000 ms | `WriteTotalTimeoutConstant` |

打开时与运行期赋值都会真下发；读回是存的那一组数。

### 6.3 任意设备名 / 端口枚举（突破 16 端口）

- `CommDevice`（C3 扩展）：任意设备路径，`\\.\CNCA0`、com0com 的虚拟对、USB CDC 的 `\\.\COM120` 都行。
  它优先于 `CommPort`，也优先于夹具钩子 `C3_MSCOMM_DEVICE`。注意 `CommPort` 本身**从没钳过 16**
  （`\\.\COM100` 一直是能开的），`CommDevice` 补的是“不是 `COM<n>` 形状”的那一半。
- `PortCount` / `PortList`（C3 扩展）：`QueryDosDeviceW(L"COM", …)` 枚举**这台机器上真实存在**的
  COM 设备，`PortList` 交回 `"COM1;COM3;COM7"` 这样的串。友好名（“USB-SERIAL CH340 (COM7)”）
  要 SetupDi 遍历，未做。

### 6.4 二进制（按字节，不按文本）

`InputBytes` / `OutputBytes`（C3 扩展）：**每字节一个字符**的 BSTR —— `ChrW(&HAA)` 就是一个 0xAA 字节，
`AscW(Mid$(s, i + 1, 1))` 就能还原。BSTR 是长度前缀的，所以内嵌 `&H00` 也不会被截断。
为什么不做 VB6 那个 `v = .Input: b = v` 的 Variant 数组形：cgen 的属性面只有标量/String 这一档，
数组要**按目标类型分派**，那是另一条 cgen 通道；本批先把“真字节不出错”落地，这一形写在 §6.6。

### 6.5 信号线与事件补全

- 运行期 `DTREnable` / `RTSEnable`：打开后改也生效（`EscapeCommFunction` 拉线），不再是“只在打开那一刻有用”。
- `SetBreak`（C3 扩展）：`SETBREAK` / `CLRBREAK`。
- 新增第二条线程：`SetCommMask(EV_CTS|EV_DSR|EV_RLSD|EV_ERR)` + 阻塞 `WaitCommEvent`，
  投 `comEvCTS(3)` / `comEvDSR(4)` / `comEvCD(5)`；`EV_ERR` 走 `ClearCommError` 的 `CE_*`
  分档成 `1004` 帧错 / `1006` 超速 / `1009` 校验错 / `1008` 收缓冲溢出 —— 不再“所有读错误都是 1008”。
- `InBufferCount = 0` 除了清我们这层缓冲，还 `PurgeComm(PURGE_RXCLEAR)`。
- `OutBufferCount` 改问 `ClearCommError` 的 `cbOutQue`（驱动里还没发完的），不再恒答 0。

### 6.6 实跑读数（`tests/ctrlmscomm`，x64 与 x86 逐字一致，`exit=0`）

```
MC1..MC9   同 §5.4（属性回读 / 开 / 写 / 收空 / EOF / 事件 / 关 / 重开 / 累计）
MC10-close2=False        改设备名前先关
MC11-dev=True            CommDevice 属性往返（不经环境变量钩子）
MC12-reopen=True|True    用 CommDevice 重开；PortCount 不炸
MC13-to=20|2000          ReadIntervalTimeout 往返；WriteTimeout 默认值
MC14-bin=8               OutputBytes 每字节一个字符：5 + 3
MC15-inb=0               InputBytes 空缓冲 → 空串
MC16-ports=True          PortList 形状（空串或以 COM 开头）
```

### 6.7 这一批**没实测到**的（诚实口径）

本机没有串口硬件，下面三条只有代码、没有读数：

1. **Settings 硬失败**那条路（写错波特率 ⇒ `PortOpen=False`）：NUL 上走的是降级档，到不了这一支。
2. **`comEvCTS/DSR/CD` 与 `CE_*` 分档**：要真串口才产生线路变化/校验错，夹具造不出来。
3. **高波特率真下发**：`921600/1500000/3000000` 的解析在（`50..12000000` 都收），但“设上并在线上跑对”
   要有对应设备才能验。
4. 端口枚举只验了**形状**：这台机器可能一个 COM 都没有（`MC16` 两种都算对）。
