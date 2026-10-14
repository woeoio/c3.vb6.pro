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
      Interval        =   100
      Left            =   240
      Top             =   240
   End
   Begin VB.Timer tOff 
      Enabled         =   0   'False
      Interval        =   100
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

' ai/029 C29-T + Fix <c3-menu3d-click>: VB.Timer 的运行期**机制**判据。
'
' 口径 (2026-10-08 改): 节拍由 Win32 SetTimer 给，也就是**系统计时 tick** 那一档 ——
' 本机与 CI 实测 tick 约 15.6 ms，Interval 不足一个 tick 就往上取整 (50 ms 会变成
' 62.5 ms)，另有一条 10 ms 钳位。这正是 VB6 Timer 控件自己的行为。曾经走过 winmm
' timeSetEvent 把精度提到 ms 级 (Interval=20 实得 ~50 拍/秒)，代价是自发出去的那条
' WM_TIMER 是**真实待处理消息**、会长期占住线程队列，把鼠标键盘饿死 (3DMenu 实测
' 点一下就不动，而 VB6 编译的同一份代码正常)，所以现在只作兜底。
'
' ⇒ 判据跟着改成**不钉绝对拍号**（那等于在钉这台机器的 tick），只钉机制：
'   T1/T2 开了就要跑；T3/T4 改 Interval 两个方向都按新周期重排；T5 关掉就停；
'   T9/T10 每格自己的周期互不串；T11/T12 (账 #156) 第二枚窗体的计时器自己跳、
'   且不串进第一枚的事件过程；T6 钉"小 Interval 到地板为止" —— 下限证明地板确实
'   在跳，上限证明没人绕过 10 ms 钳位 (winmm 那条路实测 199~200 拍/秒，一超就红)。
'   名义间隔一律取 100/200/500 ms 这一档：tick 是 15.6 ms 还是被别的进程用
'   timeBeginPeriod 提到 1 ms，取整误差都 <=7%，两种机器落进同一条带。
'
' 保留的历史读数（说明每条带在防什么）：
'   * 设计期 Enabled=0 不挂表 / Interval 没人重排 —— 负控 T2=0、T3=32、T5=16；
'   * 账 #156 的 id 撞车 —— TmForm2 的到期全打在 tOn_Timer 上：T11=0、T12 翻 ~5 倍；
'   * SetTimer 地板的本机现值 —— Interval=5 → 61 拍/秒 (CI 54~58)。

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
    Debug.Print "T1=" & InBand(mOn, 5, 13) & "/" & mOn
    Debug.Print "T2=" & InBand(mOff, 5, 13) & "/" & mOff

    ' --- 3/4: 改 Interval 立刻按新周期重排，改回来也跟着变 ---
    tOn.Interval = 200
    tOff.Interval = 200
    mOn = 0
    mOff = 0
    spent = Spin(1000)
    Debug.Print "T3=" & InBand(mOn, 3, 6) & "/" & mOn
    tOn.Interval = 100
    mOn = 0
    spent = Spin(1000)
    Debug.Print "T4=" & InBand(mOn, 5, 13) & "/" & mOn

    ' --- 5: Enabled=False 之后不该再来（最多一枚在途的） ---
    tOn.Enabled = False
    mOn = 0
    spent = Spin(500)
    Debug.Print "T5=" & TF(mOn <= 1) & "/" & mOn

    ' --- 6: 地板那一刀（Interval 小于 tick 就取整到 tick，再加 10 ms 钳位）---
    tOn.Enabled = True
    tOn.Interval = 5
    mOn = 0
    spent = Spin(1000)
    Debug.Print "T6=" & TF(mOn >= 35 And mOn <= 150) & "/" & mOn
    tOn.Interval = 100

    ' --- 7/8: 属性读数口径（含 SetProp 存 0 那一坑：Enabled=False 要读回 False）---
    Debug.Print "T7=" & TF(tOn.Enabled = True And tOn.Interval = 100)
    tOn.Enabled = False
    Debug.Print "T8=" & TF(tOn.Enabled = False And tOff.Enabled = True)
    tOn.Enabled = True

    ' --- 9: 两枚互不串：只改一枚的周期，另一枚速率不动 ---
    mOn = 0
    mOff = 0
    tOff.Interval = 500
    spent = Spin(1000)
    Debug.Print "T9=" & TF(InBand(mOn, 5, 13) = "Y" And InBand(mOff, 1, 4) = "Y")

    ' --- 10: 恢复同周期后两枚都按 100ms 跑 ---
    tOff.Interval = 100
    mOn = 0
    mOff = 0
    spent = Spin(1000)
    Debug.Print "T10=" & TF(mOn >= 5 And mOff >= 5)

    ' --- 11/12: 账 #156 第二枚窗体的 Timer 自己跳、且不串进第一枚的事件过程 ---
    ' 周期故意差 5 倍 (第二枚 100 vs 第一枚 500)，翻红时差的是量级、不是抖动。
    tOff.Enabled = False
    tOn.Interval = 500
    mOn = 0
    mDlg = 0
    TmForm2.Show vbModeless
    spent = Spin(1000)
    Debug.Print "T11=" & TF(mDlg >= 5) & "/" & mDlg
    Debug.Print "T12=" & TF(mOn <= 4) & "/" & mOn
    Unload TmForm2

    Debug.Print "TIMERPROG-DONE"
    Unload Me
End Sub
