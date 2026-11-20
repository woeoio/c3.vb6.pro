VERSION 5.00
Begin VB.Form UpForm 
   Caption         =   "UpForm"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   LinkTopic       =   "UpForm"
   ScaleHeight     =   3200
   ScaleWidth      =   4200
   Begin MSComctlLib.UpDown ud1 
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   400
      Min             =   5
      Max             =   25
      Value           =   10
      Increment       =   3
      _ExtentX        =   706
      _ExtentY        =   706
   End
   Begin MSComctlLib.UpDown ud2 
      Alignment       =   0
      Height          =   400
      Left            =   120
      Orientation     =   1
      TabIndex        =   1
      Top             =   720
      Width           =   800
      Wrap            =   -1  'True
      _ExtentX        =   1411
      _ExtentY        =   706
   End
   Begin VB.TextBox txtB 
      Height          =   300
      Left            =   720
      TabIndex        =   2
      Text            =   "10"
      Top             =   180
      Width           =   1200
   End
   Begin MSComctlLib.UpDown ud3 
      Height          =   400
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   400
      Min             =   0
      Max             =   50
      Increment       =   5
      BuddyControl    =   "txtB"
      SyncBuddy       =   -1  'True
      _ExtentX        =   706
      _ExtentY        =   706
   End
   Begin VB.Timer tGo 
      Enabled         =   -1  'True
      Interval        =   60
      Left            =   3600
      Top             =   2640
   End
End
Attribute VB_Name = "UpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' C30-UD-a 的判据。这一批只钉**属性面**：UpDown 没有几何假证人可踩，
' 值面四条（Min/Max/Value/Increment）与行为面四条（Wrap/Alignment/Orientation/SyncBuddy）
' 都能直接读回窗口；事件面（UpClick/DownClick/Change）要等 UD-b 的合成驱动，
' 这一件夹具**不碰**那三条处理器 —— 免得"编得过、没人送消息"那种绿。
'
' 边界那几条（越界写值、Min 越过 Max）刻意钉住"控件自己钳"，因为我们不再钳第二遍：
' 两遍钳会出现"我们以为的值"和"控件真在显示的值"两个答案。

' C30-UD-b: 事件面的三条计数器。处理器本身必须**真的存在**，否则 cgen 那条 -721 的
' arm 根本不会生成（`if (!upSym && !dnSym && !chSym) continue;`）—— 那么判据就会在
' 一个"没有接收方"的窗上验派发，绿得毫无意义。
Private gUp As Long
Private gDn As Long
Private gCh As Long
' C30-UD-c: 处理器里读到的 Value 是“新”还是“旧”，只有窗口里的真实时序能答。
' 真手势实测（tools/updown_manual）：Up/Down 在值变更前发（读旧值），
' Change 应该在值落定后发（读新值）—— 早期这里 Change 也读旧值，业务代码会差一格。
Private gUpVal As Long
Private gChVal As Long

Private Sub ud1_UpClick()
    gUp = gUp + 1
    gUpVal = ud1.Value
End Sub

Private Sub ud1_DownClick()
    gDn = gDn + 1
End Sub

Private Sub ud1_Change()
    gCh = gCh + 1
    gChVal = ud1.Value
End Sub

Private Sub tGo_Timer()
    tGo.Enabled = False
    Debug.Print "UD1-range=" & CStr(ud1.Min) & "/" & CStr(ud1.Max)
    Debug.Print "UD2-value=" & CStr(ud1.Value)
    Debug.Print "UD3-inc=" & CStr(ud1.Increment)
    ' ud1 没写 Wrap（默认 False=0），ud2 写了 -1 ⇒ 这一行同时钉两档。
    Debug.Print "UD4-wrap=" & CStr(ud1.Wrap) & "/" & CStr(ud2.Wrap)
    ' Alignment：ud1 没写 ⇒ VB6 默认 1（右）；ud2 写 0 ⇒ 左。
    Debug.Print "UD5-align=" & CStr(ud1.Alignment) & "/" & CStr(ud2.Alignment)
    ' Orientation：ud1 没写 ⇒ 0（竖）；ud2 写 1 ⇒ 横。
    Debug.Print "UD6-ori=" & CStr(ud1.Orientation) & "/" & CStr(ud2.Orientation)
    ' 越界由控件钳：25 是上限、5 是下限。
    ud1.Value = 30
    Debug.Print "UD7-clampmax=" & CStr(ud1.Value)
    ud1.Value = -4
    Debug.Print "UD8-clampmin=" & CStr(ud1.Value)
    ' 运行期改值面四条都该真生效（不是只改我们的自存）。
    ud1.Value = 7
    ud1.Increment = 4
    ud1.Max = 12
    Debug.Print "UD9-rt=" & CStr(ud1.Value) & "/" & CStr(ud1.Increment) & "/" & CStr(ud1.Max)
    ' 运行期翻行为面（样式位）也该读得回来。
    ud1.Wrap = True
    ud1.Alignment = 0
    ud1.Orientation = 1
    Debug.Print "UD10-rtstyle=" & CStr(ud1.Wrap) & "/" & CStr(ud1.Alignment) & "/" & CStr(ud1.Orientation)
    ' SyncBuddy 这一批只钉"写进去读得回来"；buddy 文本联动要真消息，归 UD-b。
    ud1.SyncBuddy = True
    Debug.Print "UD11-sync=" & CStr(ud1.SyncBuddy) & "/" & CStr(ud2.SyncBuddy)
    Debug.Print "UD12-type=" & TypeName(ud1.Value) & "/" & CStr(VarType(ud1.Value))

    ' ---------------- C30-UD-b：事件面 ----------------
    ' 用 SimDelta(格数) 造一条**真** UDN_DELTAPOS 递到父窗：它先按 Increment 把值移到位，
    ' 再把 NMUPDOWN 交给窗体的 WM_NOTIFY —— 「认来源 → 分流 → 调用」三段都要走一遍，
    ' 所以这三行钉的既是值也是那三个处理器有没有被真调到。
    ' 先把量程/步长复位（上面 UD9 把 Max 改成了 12、Increment 改成了 4）。
    ud1.Min = 5
    ud1.Max = 25
    ud1.Value = 10
    ud1.Increment = 3
    gUp = 0
    gDn = 0
    gCh = 0
    Debug.Print "UE1-up=" & CStr(ud1.SimDelta(1))
    Debug.Print "UE2-cnt=" & CStr(gUp) & "/" & CStr(gDn) & "/" & CStr(gCh)
    Debug.Print "UE3-dn=" & CStr(ud1.SimDelta(-1))
    Debug.Print "UE4-cnt=" & CStr(gUp) & "/" & CStr(gDn) & "/" & CStr(gCh)
    ' 一次两格：iDelta 的**大小**不进处理器（VB6 那三条本来就不带参数），只进值。
    Debug.Print "UE5-two=" & CStr(ud1.SimDelta(2))
    Debug.Print "UE6-cnt=" & CStr(gUp) & "/" & CStr(gDn) & "/" & CStr(gCh)
    ' C30-UD-c：同一串 E 段里，两条时间线分开钉 ——
    '   UE7 = Change 里读到的值（必须已经是**新值** 16）
    '   UE8 = UpClick 里读到的值（必须还是**旧值** 10，与真手势一致）
    ' 判据把两档写在一次 SimDelta(1) 上，改坏了任一侧都会立刻掉。
    Debug.Print "UE7-chgval=" & CStr(gChVal)
    Debug.Print "UE8-upval=" & CStr(gUpVal)

    ' ---------------- C30-UD-b：buddy 文本联动 ----------------
    ' ud3 是**设计期**挂的 buddy（BuddyControl = "txtB"，SyncBuddy = -1）：
    ' 值一变，UDS_SETBUDDYINT 就该把新位置写进 txtB 的正文。写值那一路与 SimDelta 那一路
    ' 各钉一次 —— 前者是 UDM_SETPOS32 的自发结果，后者走我们合成的那条通知。
    ud3.Value = 20
    Debug.Print "UB1-buddy=" & txtB.Text
    ud3.SimDelta(1)
    Debug.Print "UB2-buddy=" & txtB.Text & "/" & CStr(ud3.Value)

    Debug.Print "UPDOWN-DONE"
    Unload Me
End Sub
