VERSION 5.00
Begin VB.Form CbForm 
   Caption         =   "CbForm"
   ClientHeight    =   1900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   LinkTopic       =   "CbForm"
   ScaleHeight     =   1900
   ScaleWidth      =   4200
   Begin VB.Timer t1 
      Enabled         =   -1   'True
      Interval        =   40
      Left            =   3800
      Top             =   120
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   350
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   700
      Width           =   700
   End
   Begin VB.ListBox lb1 
      Height          =   450
      Left            =   1440
      TabIndex        =   2
      Top             =   700
      Width           =   600
   End
   Begin VB.ComboBox cb2 
      Height          =   285
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   700
      Width           =   600
   End
   Begin VB.ComboBox cb1 
      Height          =   285
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   600
   End
End
Attribute VB_Name = "CbForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #158 的 Combo 那一半：`_GotFocus` / `_LostFocus` 以前**永远不触发** —— 发码那两支撑的是
' `code == 1024 / 2048`，而那是 **CBEM_*（发给 ComboBoxEx 的消息号）**，不是 WM_COMMAND 的
' notification code；真正的码是 `CBN_SETFOCUS=3` / `CBN_KILLFOCUS=4`（同一段里 EN_=256/512、
' LBN_=4/5、BN_=6/7 三张表都是对的，只有 Combo 这一格抄错）。
' 判据一次问三件事，不只问"改对了"：
'   CF-cb1 / CF-cb2 = 六枚计数器各自恰好 1（**没改之前 Combo 那四条全是 0**）
'   CF-lb / CF-tx   = ListBox 与 TextBox 那两张表**没被这一刀带坏**（同一种"焦点通知走父窗
'                     WM_COMMAND"的机制，各用自己的 code，改 Combo 时很容易顺手碰）
'   CF-cross        = 串台检查：`id ==` 那道过滤还在起作用 —— ListBox 的 LBN_SETFOCUS=4 与
'                     Combo 的 CBN_KILLFOCUS=4 **同号**，要是哪天有人把 id 过滤挪走，
'                     lb1 拿焦点就会把 cb1/cb2 的 LostFocus 一起点着 ⇒ 这一条当场红。
' cb2 额外挂一条 `_Validate`：走的是"带 Validate 时 LostFocus 要问 `VB6_ValidateCancel`"那
' 一个分支（与不带的那条是两份发码），不覆盖就等于那条分支从没跑过。
Private gTick As Long
Private gCb1Got As Long
Private gCb1Lost As Long
Private gCb2Got As Long
Private gCb2Lost As Long
Private gCb2Val As Long
Private gLbGot As Long
Private gLbLost As Long
Private gTxGot As Long
Private gTxLost As Long

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Log1(ByVal s As String)
    Dim h As Long
    h = FreeFile
    Open App.Path & "\combofocus.log" For Append As #h
    Print #h, s
    Close #h
    Debug.Print s
End Sub

Private Sub cb1_GotFocus()
    gCb1Got = gCb1Got + 1
End Sub

Private Sub cb1_LostFocus()
    gCb1Lost = gCb1Lost + 1
End Sub

Private Sub cb2_GotFocus()
    gCb2Got = gCb2Got + 1
End Sub

Private Sub cb2_LostFocus()
    gCb2Lost = gCb2Lost + 1
End Sub

Private Sub cb2_Validate(Cancel As Boolean)
    gCb2Val = gCb2Val + 1
    ' 刻意**不**置 Cancel：这一条只为把"带 Validate 时 LostFocus 走那条问 `VB6_ValidateCancel`
    ' 的分支"跑起来；置了 Cancel 就会把 LostFocus 也吃掉，判据要问的就变成另一件事了。
End Sub

Private Sub lb1_GotFocus()
    gLbGot = gLbGot + 1
End Sub

Private Sub lb1_LostFocus()
    gLbLost = gLbLost + 1
End Sub

Private Sub txt1_GotFocus()
    gTxGot = gTxGot + 1
End Sub

Private Sub txt1_LostFocus()
    gTxLost = gTxLost + 1
End Sub

' 每拍一次程序化 SetFocus —— Combo 这两条通知的原生来源就是窗口管理器自己发的 WM_COMMAND
' （探针实测：`cb1.SetFocus` 之后 got/lost 各 1，改码之前各 0），所以判据不伪造通知。
Private Sub t1_Timer()
    gTick = gTick + 1
    If gTick = 1 Then
        cb1.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 2 Then
        cb2.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 3 Then
        cmdGo.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 4 Then
        lb1.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 5 Then
        txt1.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 6 Then
        cmdGo.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 7 Then
        Log1 "CF-cb1=" & CStr(gCb1Got) & "/" & CStr(gCb1Lost)
        Log1 "CF-cb2=" & CStr(gCb2Got) & "/" & CStr(gCb2Lost) & "/" & CStr(gCb2Val)
        Log1 "CF-lb=" & CStr(gLbGot) & "/" & CStr(gLbLost)
        Log1 "CF-tx=" & CStr(gTxGot) & "/" & CStr(gTxLost)
        Log1 "CF-each=" & TF(gCb1Got = 1 And gCb1Lost = 1 And gCb2Got = 1 And gCb2Lost = 1 _
                              And gLbGot = 1 And gLbLost = 1 And gTxGot = 1 And gTxLost = 1)
        ' 串台读数：每枚控件各算各的，一次进出焦点不该把邻居的计数也带上
        Log1 "CF-cross=" & TF(gCb1Got + gCb1Lost + gCb2Got + gCb2Lost + gLbGot + gLbLost + gTxGot + gTxLost = 8)
        Log1 "COMBOFOCUS-DONE"
        t1.Enabled = False
        Unload Me
    End If
End Sub
