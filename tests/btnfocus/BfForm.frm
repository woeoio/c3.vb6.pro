VERSION 5.00
Begin VB.Form BfForm 
   Caption         =   "BfForm"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "BfForm"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   Begin VB.Timer t1 
      Enabled         =   -1   'True
      Interval        =   40
      Left            =   4400
      Top             =   120
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   350
      Left            =   3600
      TabIndex        =   8
      Top             =   1800
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Caption         =   "F"
      Height          =   1000
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   2000
      Begin VB.CommandButton cmdIn 
         Caption         =   "IN"
         Height          =   350
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   800
      End
      Begin VB.TextBox txtIn 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   480
         Width           =   700
      End
   End
   Begin VB.TextBox txtA 
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   700
      Width           =   700
   End
   Begin VB.OptionButton optA 
      Caption         =   "opt"
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   700
      Width           =   800
   End
   Begin VB.CheckBox chkA 
      Caption         =   "chk"
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   700
      Width           =   800
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A"
      Height          =   350
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   800
   End
End
Attribute VB_Name = "BfForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #158 的按钮那一半：`_GotFocus` / `_LostFocus` 以前**永远不触发**，但根因不是码表
' （那两支撑的一直是对的 `BN_SETFOCUS=6` / `BN_KILLFOCUS=7`），缺的是创建时没挂
' **`BS_NOTIFY`** —— BUTTON 类不挂这一位就不把焦点变化作为 WM_COMMAND 报给父窗。
' 裸码探针 `.build/bnnotify/bnnotify.c`（同一父窗两枚按钮，只差这一位）实测：挂了的
' 那枚，程序化 `SetFocus` 与对话框管理器（VK_TAB）两条路都把 6/7 送到父窗；没挂的一条
' 都不送。而 `BN_CLICKED=0` **不需要**这一位 ⇒ 按钮的 `_Click` 一直是通的。
' 第二件事在同一次发码读数里量到：焦点码表那一趟以前**只遍历顶层控件**（ID 100+），
' 容器子控件（ID 200+）一条 arm 都没发过 ⇒ 放在 Frame 里的按钮/文本框连"有没有码表"
' 都还没轮到。本批把那一趟收成两处共用的一个 lambda，这里就按**两条路读同一串数**判。
' 判据问四件事，不许省：
'   BF-cmd / BF-chk / BF-opt = 顶层三型按钮各一进一出（这一刀本身）
'   BF-tx                    = TextBox 那张表（EN_=256/512）**没被带坏**
'   BF-in / BF-intx          = 容器那一趟：Frame 里的按钮与文本框各一进一出
'   BF-both                  = 两条创建路读同一串数（顶层按钮 1/1 == 容器按钮 1/1）
'   BF-noclick               = **挂上 BS_NOTIFY 之后焦点移动不许算成点击** —— 同一个 id
'                              现在会携 code=6/7 进来，Click 那条 arm 不加 `code == 0`
'                              过滤的话，这三枚按钮的 `_Click` 会被焦点各点一次
'                              账 #171 补上**单选钮那一半**：optA 本夹具里没勾着、第 3 拍把焦点
'                              移进它 —— OS 替父窗发的就是 `BN_CLICKED`（号 0），光按 code 筛不掉，
'                              要按「来路自己勾上了没有」筛（RTL `vb6_RadioClickCounts`）。
'                              修之前这一条读 False（焦点移动被算成一次点击），修后 True。
'   BF-each                  = 十二个计数器各自恰好 1（双发检查，本线已知缺陷族 #161）
Private gTick As Long
Private gCmdGot As Long
Private gCmdLost As Long
Private gCmdClk As Long
Private gChkGot As Long
Private gChkLost As Long
Private gChkClk As Long
Private gOptGot As Long
Private gOptClk As Long
Private gOptLost As Long
Private gTxGot As Long
Private gTxLost As Long
Private gInGot As Long
Private gInLost As Long
Private gInClk As Long
Private gInTxGot As Long
Private gInTxLost As Long

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Log1(ByVal s As String)
    Dim h As Long
    h = FreeFile
    Open App.Path & "\btnfocus.log" For Append As #h
    Print #h, s
    Close #h
    Debug.Print s
End Sub

Private Sub cmdA_GotFocus()
    gCmdGot = gCmdGot + 1
End Sub

Private Sub cmdA_LostFocus()
    gCmdLost = gCmdLost + 1
End Sub

Private Sub cmdA_Click()
    gCmdClk = gCmdClk + 1
End Sub

Private Sub chkA_GotFocus()
    gChkGot = gChkGot + 1
End Sub

Private Sub chkA_LostFocus()
    gChkLost = gChkLost + 1
End Sub

Private Sub chkA_Click()
    gChkClk = gChkClk + 1
End Sub

Private Sub optA_Click()
    ' 账 #171：单选钮是这一格里唯一「焦点进入也会送来 `BN_CLICKED`」的一型
    ' （#158 那道 `code == 0` 对它无效，因为号本身就是 0）⇒ 这一枚计数器就是那条口径的证人。
    gOptClk = gOptClk + 1
End Sub

Private Sub optA_GotFocus()
    gOptGot = gOptGot + 1
End Sub

Private Sub optA_LostFocus()
    gOptLost = gOptLost + 1
End Sub

Private Sub txtA_GotFocus()
    gTxGot = gTxGot + 1
End Sub

Private Sub txtA_LostFocus()
    gTxLost = gTxLost + 1
End Sub

Private Sub cmdIn_GotFocus()
    gInGot = gInGot + 1
End Sub

Private Sub cmdIn_LostFocus()
    gInLost = gInLost + 1
End Sub

Private Sub cmdIn_Click()
    gInClk = gInClk + 1
End Sub

Private Sub txtIn_GotFocus()
    gInTxGot = gInTxGot + 1
End Sub

Private Sub txtIn_LostFocus()
    gInTxLost = gInTxLost + 1
End Sub

' 每拍一次程序化 SetFocus，走完整条原生路：按钮 → OS 发 BN_SETFOCUS/BN_KILLFOCUS 到父窗
' → 生成的 WndProc 按 (id, code) 派发。不伪造通知，也不发 WM_COMMAND。
' 最后一拍落到 cmdGo（**没挂任何处理器**）：让 txtIn 的 LostFocus 有个收尾，
' 这样每一枚控件的进出都各是 1，`BF-each` 才是"恰好一次"而不是"至少一次"。
Private Sub t1_Timer()
    gTick = gTick + 1
    If gTick = 1 Then
        cmdA.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 2 Then
        chkA.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 3 Then
        optA.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 4 Then
        txtA.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 5 Then
        cmdIn.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 6 Then
        txtIn.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 7 Then
        cmdGo.SetFocus
        DoEvents
        Exit Sub
    End If
    If gTick = 8 Then
        Log1 "BF-cmd=" & CStr(gCmdGot) & "/" & CStr(gCmdLost)
        Log1 "BF-chk=" & CStr(gChkGot) & "/" & CStr(gChkLost)
        Log1 "BF-opt=" & CStr(gOptGot) & "/" & CStr(gOptLost)
        Log1 "BF-tx=" & CStr(gTxGot) & "/" & CStr(gTxLost)
        Log1 "BF-in=" & CStr(gInGot) & "/" & CStr(gInLost)
        Log1 "BF-intx=" & CStr(gInTxGot) & "/" & CStr(gInTxLost)
        Log1 "BF-both=" & TF(gCmdGot = gInGot And gCmdLost = gInLost And gTxGot = gInTxGot)
        Log1 "BF-noclick=" & TF(gCmdClk = 0 And gChkClk = 0 And gInClk = 0 And gOptClk = 0)
        Log1 "BF-each=" & TF(gCmdGot = 1 And gCmdLost = 1 And gChkGot = 1 And gChkLost = 1 _
                              And gOptGot = 1 And gOptLost = 1 And gTxGot = 1 And gTxLost = 1 _
                              And gInGot = 1 And gInLost = 1 And gInTxGot = 1 And gInTxLost = 1)
        Log1 "BTNFOCUS-DONE"
        t1.Enabled = False
        Unload Me
    End If
End Sub
