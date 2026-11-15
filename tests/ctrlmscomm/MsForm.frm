VERSION 5.00
Begin VB.Form MsForm 
   Caption         =   "MsForm"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "MsForm"
   ScaleHeight     =   2400
   ScaleWidth      =   4560
   Begin MSCommLib.MSComm MSComm1 
      Height          =   360
      Left            =   120
      Top             =   120
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   635
      _Version        =   393216
      CommPort        =   1
      DTREnable       =   -1  'True
      EOFEnable       =   -1  'True
      InputLen        =   0
      RThreshold      =   1
      Settings        =   "9600,N,8,1"
   End
   Begin VB.Timer tGo 
      Interval        =   60
      Left            =   1200
      Top             =   120
   End
   Begin VB.Label lblOut 
      Caption         =   "out"
      Height          =   255
      Left            =   120
      Top             =   720
      Width           =   3615
   End
End
Attribute VB_Name = "MsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' C30-4: MSComm 的无头判据（与 tests/ctrlmaskedbox 同构：Timer 起、Debug.Print 出）。
' 本机没有串口硬件 ⇒ 跑的时候设 C3_MSCOMM_DEVICE=\\.\NUL（C3 的测试钩子，不是 VB6 API），
' 让"开-写-计数-EOF-关"这一路真走一遍 CreateFileW / WriteFile / ReadFile。
Option Explicit

Private gEv As Long          ' OnComm 被调了几次
Private gLast As Long        ' 最近一次 CommEvent

Private Sub MSComm1_OnComm()
    gEv = gEv + 1
    gLast = MSComm1.CommEvent
End Sub

Private Sub tGo_Timer()
    Dim s As String
    tGo.Enabled = False

    ' 1) 设计期属性回读（.frm 上写的那几条）
    Debug.Print "MC1-prop=" & CStr(MSComm1.CommPort) & "|" & MSComm1.Settings & "|" & CStr(MSComm1.InputLen)

    ' 2) 真打开（降级设备 NUL 也认）
    MSComm1.PortOpen = True
    Debug.Print "MC2-open=" & CStr(MSComm1.PortOpen)

    ' 3) 真写出去：VB6 没有这条属性，C3 的证人 TxTotal（字节数）
    MSComm1.Output = "AT"
    Debug.Print "MC3-tx=" & CStr(MSComm1.TxTotal)

    ' 4) 收缓冲与 Input：NUL 上没有回包，应为 0 与空串
    s = MSComm1.Input
    Debug.Print "MC4-in=" & CStr(MSComm1.InBufferCount) & "|" & CStr(Len(s))

    ' 5) EOF：NUL 读到 0 字节 ⇒ 读线程报 EOF；EOFEnable=True 时 CommEvent=comEvEOF(7)
    Dim i As Long
    For i = 1 To 200
        DoEvents
        If gLast <> 0 Then Exit For
    Next
    Debug.Print "MC5-ev=" & CStr(gLast)

    ' 6) OnComm 至少到过一次（事件面真的接上了，不是只存了属性）
    Debug.Print "MC6-cnt=" & CStr(gEv > 0)

    ' 7) 关端口
    MSComm1.PortOpen = False
    Debug.Print "MC7-close=" & CStr(MSComm1.PortOpen)

    ' 8) 关掉再开（VB6 里这页能重开，是最常见的翻车点：句柄没清干净）
    MSComm1.PortOpen = True
    Debug.Print "MC8-reopen=" & CStr(MSComm1.PortOpen)

    ' 9) 第二次写：计数器继续累加（TxTotal 是累计，不是"这一次"）
    MSComm1.Output = "ABC"
    Debug.Print "MC9-tx2=" & CStr(MSComm1.TxTotal)

    ' 10) 收尾
    MSComm1.PortOpen = False
    Debug.Print "MS-DONE"
    Unload Me
End Sub
