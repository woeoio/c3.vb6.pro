VERSION 5.00
Begin VB.Form TmForm 
   Caption         =   "TmForm"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4000
   LinkTopic       =   "TmForm"
   ScaleHeight     =   2400
   ScaleWidth      =   4000
   Begin VB.Timer tOn 
      Enabled         =   -1   'True
      Interval        =   20
      Left            =   240
      Top             =   240
   End
   Begin VB.Timer tOff 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   720
      Top             =   240
   End
End
Attribute VB_Name = "TmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-T: VB.Timer 运行期真触发 + 精度提到 ms 级。
' 改之前的实测读数（029 §九 C29-T 那一格）：
'   * 设计期 Enabled=0 的 Timer 压根不挂表，运行期 `Timer1.Enabled = True` 发成
'     vb6_SetTimerEnabled(vb6_hwnd_<timer>, ...)，而 Timer 是无窗口控件、句柄恒 NULL
'     => SetPropW(NULL,...) 静默丢，tick 数永远是 0；
'   * `Timer1.Interval = 100` 同理没人重排周期，速率纹丝不动；
'   * 精度只有 SetTimer 那一档 ~15.6ms 地板（Interval=20 实得 34.5ms/tick、Interval=5 封顶 ~64/tick）。
' 现在 Enabled/Interval 真的起停与重排，底层走 winmm timeSetEvent（取不到则退回 SetTimer）。
'' 读数带区间是按'秒级墙钟窗口里的 tick 数'算的：20ms 名义 50 次，允许 [40,60]。

Declare Function GetTickCount Lib "kernel32" () As Long

Private mOn As Long
Private mOff As Long
Private mDlg As Long

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Function InBand(ByVal n As Long, ByVal lo As Long, ByVal hi As Long) As String
    If n >= lo And n <= hi Then InBand = "Y" Else InBand = "N"
End Function

Private Sub tOn_Timer()
    mOn = mOn + 1
End Sub

Private Sub tOff_Timer()
    mOff = mOff + 1
End Sub

' 账 #156: 第二枚窗体的 Timer 自己数，只有它自己的事件过程会调这里。
Public Sub SetDlg(ByVal n As Long)
    mDlg = n
End Sub

Private Function Spin(ByVal ms As Long) As Long
    Dim t0 As Long
    t0 = GetTickCount()
    Do While GetTickCount() - t0 < ms
        DoEvents
    Loop
    Spin = GetTickCount() - t0
End Function

Private Sub Form_Load()
    Dim spent As Long

    ' --- 1/2: 设计期开着的那枚要跑；设计期关着的，运行期 Enabled=True 必须起得来 ---
    tOff.Enabled = True
    spent = Spin(1000)
    Debug.Print "T1=" & InBand(mOn, 40, 60) & "/" & mOn
    Debug.Print "T2=" & InBand(mOff, 40, 60) & "/" & mOff

    ' --- 3/4: 改 Interval 立刻按新周期重排，改回来也跟着变 ---
    tOn.Interval = 100
    tOff.Interval = 100
    mOn = 0
    mOff = 0
    spent = Spin(1000)
    Debug.Print "T3=" & InBand(mOn, 8, 13) & "/" & mOn
    tOn.Interval = 20
    mOn = 0
    spent = Spin(1000)
    Debug.Print "T4=" & InBand(mOn, 40, 60) & "/" & mOn

    ' --- 5: Enabled=False 之后不该再来（最多一枚在途的） ---
    tOn.Enabled = False
    mOn = 0
    spent = Spin(500)
    Debug.Print "T5=" & TF(mOn <= 1) & "/" & mOn

    ' --- 6: 精度那一刀。SetTimer 地板下 Interval=5 也只到 ~64/tick；ms 级能到 ~200 ---
    tOn.Enabled = True
    tOn.Interval = 5
    mOn = 0
    spent = Spin(1000)
    Debug.Print "T6=" & TF(mOn >= 100) & "/" & mOn
    tOn.Interval = 20

    ' --- 7/8: 属性读数口径（含 SetProp 存 0 那一坑：Enabled=False 要读回 False）---
    Debug.Print "T7=" & TF(tOn.Enabled = True And tOn.Interval = 20)
    tOn.Enabled = False
    Debug.Print "T8=" & TF(tOn.Enabled = False And tOff.Enabled = True)
    tOn.Enabled = True

    ' --- 9: 两枚互不串：只改一枚的周期，另一枚速率不动 ---
    mOn = 0
    mOff = 0
    tOff.Interval = 200
    spent = Spin(1000)
    Debug.Print "T9=" & TF(InBand(mOn, 40, 60) = "Y" And InBand(mOff, 3, 7) = "Y")

    ' --- 10: 恢复同周期后两枚都按 20ms 跑 ---
    tOff.Interval = 20
    mOn = 0
    mOff = 0
    spent = Spin(600)
    Debug.Print "T10=" & TF(mOn >= 20 And mOff >= 20)

    ' --- 11/12: 账 #156 第二枚窗体的 Timer 自己跳、且不串进第一枚的事件过程 ---
    ' 改之前的实测读数：vb6_TimerAttach 的 id 取自控件 id 那个计数器，而每建一枚窗体
    ' 编译器都发一次 vb6_ResetControlId() ⇒ 第二枚窗体的 Timer 与第一枚的第一枚 Timer
    ' 同号；派发按 id 查进程内那张表、先建的那格先命中，于是 TmForm2 的到期全打在
    ' tOn_Timer 上：mDlg 恒 0（T11 红）、mOn 的速率翻倍（T12 红）。
    ' 两枚故意差 5 倍周期，翻红时差的是量级、不是抖动。
    tOff.Enabled = False
    tOn.Interval = 100
    mOn = 0
    mDlg = 0
    TmForm2.Show vbModeless
    spent = Spin(1000)
    Debug.Print "T11=" & TF(mDlg >= 25) & "/" & mDlg
    Debug.Print "T12=" & TF(mOn <= 20) & "/" & mOn
    Unload TmForm2

    Debug.Print "TIMERPROG-DONE"
    Unload Me
End Sub