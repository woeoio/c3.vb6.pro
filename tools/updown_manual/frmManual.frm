VERSION 5.00
Begin VB.Form frmManual 
   Caption         =   "UpDown manual acceptance"
   ClientHeight    =   7600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11200
   LinkTopic       =   "frmManual"
   ScaleHeight     =   7600
   ScaleWidth      =   11200
   Begin VB.TextBox txtA 
      Height          =   300
      Left            =   3900
      TabIndex        =   0
      Text            =   "5"
      Top             =   600
      Width           =   900
   End
   Begin MSComctlLib.UpDown udA 
      Height          =   320
      Left            =   4860
      TabIndex        =   1
      Top             =   590
      Width           =   340
      Min             =   0
      Max             =   10
      Value           =   5
      Increment       =   1
      BuddyControl    =   "txtA"
      SyncBuddy       =   -1  'True
      _ExtentX        =   600
      _ExtentY        =   565
   End
   Begin VB.TextBox txtB 
      Height          =   300
      Left            =   3900
      TabIndex        =   2
      Text            =   "10"
      Top             =   1100
      Width           =   900
   End
   Begin MSComctlLib.UpDown udB 
      Height          =   320
      Left            =   4860
      TabIndex        =   3
      Top             =   1090
      Width           =   340
      Min             =   0
      Max             =   10
      Value           =   10
      Increment       =   1
      _ExtentX        =   600
      _ExtentY        =   565
   End
   Begin MSComctlLib.UpDown udC 
      Height          =   320
      Left            =   3900
      TabIndex        =   4
      Top             =   1590
      Width           =   340
      Min             =   0
      Max             =   10
      Value           =   10
      Increment       =   1
      Wrap            =   -1  'True
      _ExtentX        =   600
      _ExtentY        =   565
   End
   Begin MSComctlLib.UpDown udD 
      Height          =   320
      Left            =   3900
      TabIndex        =   5
      Top             =   2090
      Width           =   340
      Min             =   -20
      Max             =   20
      Value           =   0
      Increment       =   5
      _ExtentX        =   600
      _ExtentY        =   565
   End
   Begin MSComctlLib.UpDown udE 
      Alignment       =   0
      Height          =   300
      Left            =   3900
      TabIndex        =   6
      Top             =   2590
      Width           =   900
      Min             =   0
      Max             =   10
      Value           =   5
      Increment       =   1
      Orientation     =   1
      _ExtentX        =   1587
      _ExtentY        =   529
   End
   Begin VB.Label lblA 
      Caption         =   "G1 base + buddy  (Min0 Max10 Inc1 V5)"
      Height          =   240
      Left            =   300
      TabIndex        =   7
      Top             =   640
      Width           =   3500
   End
   Begin VB.Label lblB 
      Caption         =   "G2 top edge, Wrap=False (V10=Max)"
      Height          =   240
      Left            =   300
      TabIndex        =   8
      Top             =   1140
      Width           =   3500
   End
   Begin VB.Label lblC 
      Caption         =   "G3 Wrap=True (V10=Max)"
      Height          =   240
      Left            =   300
      TabIndex        =   9
      Top             =   1640
      Width           =   3500
   End
   Begin VB.Label lblD 
      Caption         =   "G4 Increment=5, Min=-20 (V0)"
      Height          =   240
      Left            =   300
      TabIndex        =   10
      Top             =   2140
      Width           =   3500
   End
   Begin VB.Label lblE 
      Caption         =   "G5 Horizontal + Align=Left (V5)"
      Height          =   240
      Left            =   300
      TabIndex        =   11
      Top             =   2640
      Width           =   3500
   End
   Begin VB.TextBox txtState 
      Height          =   2700
      Left            =   5600
      MultiLine       =   -1  'True
      ScrollBars      =   2
      TabIndex        =   12
      Top             =   590
      Width           =   5400
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset counters + log"
      Height          =   360
      Left            =   3900
      TabIndex        =   13
      Top             =   2960
      Width           =   2400
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   360
      Left            =   6420
      TabIndex        =   14
      Top             =   2960
      Width           =   1200
   End
   Begin VB.TextBox txtLog 
      Height          =   3000
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2
      TabIndex        =   15
      Top             =   3450
      Width           =   10700
   End
   Begin VB.Timer tmGo 
      Interval        =   300
      Left            =   7900
      Top             =   2960
   End
End
Attribute VB_Name = "frmManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' UpDown 手动验收夹具 (C30-UD-c)。存在的理由只有一条:
'   无头环境点不了鼠标箭头 ⇒ 判据只能用 vb6_UpDown_SimDelta 合成 UDN_DELTAPOS,
'   而"真手势到底发不发通知、值什么时候才变、量程边界上 Change 响不响"这三件事
'   只有真鼠标点击才能回答。本夹具把每一次事件按发生顺序记下来、并落盘到
'   App.Path\UpManual.log, 由人工按 tests/updown_manual/README.md 的脚本点一遍。
'
' 关键信息:**处理器里读到的 udX.Value 是"通知到时"的值**。它和 timer 快照里的
' 值不一样就说明通知发生在值变更之前 (原生 UDN_DELTAPOS 的文档语义就是"将要变"),
' 这条差异是判断我们派发时序对不对的唯一读数, 别在事后把它改平。

Private gSeq As Long
Private gLog As String
Private gLastWritten As Long

Private gUpA As Long
Private gDnA As Long
Private gChA As Long
Private gUpB As Long
Private gDnB As Long
Private gChB As Long
Private gUpC As Long
Private gDnC As Long
Private gChC As Long
Private gUpD As Long
Private gDnD As Long
Private gChD As Long
Private gUpE As Long
Private gDnE As Long
Private gChE As Long

Private Sub AddLog(ByVal s As String)
    gSeq = gSeq + 1
    gLog = gLog & "#" & CStr(gSeq) & " " & s & vbCrLf
End Sub

' -1 是故意的: 让 timer 第一拍就落一次盘 (否则窗口开着、一次没点时文件根本不存在,
' 人工会以为是写文件挂了)。Form_Load 里还会记一条起点, 回报时能看出"从哪开始算"。
Private Sub Form_Load()
    gLastWritten = -1
    AddLog "form loaded"
End Sub

Private Sub udA_UpClick()
    gUpA = gUpA + 1
    AddLog "udA UpClick    v=" & CStr(udA.Value)
End Sub

Private Sub udA_DownClick()
    gDnA = gDnA + 1
    AddLog "udA DownClick  v=" & CStr(udA.Value)
End Sub

Private Sub udA_Change()
    gChA = gChA + 1
    AddLog "udA Change     v=" & CStr(udA.Value) & "  buddy=" & txtA.Text
End Sub

Private Sub udB_UpClick()
    gUpB = gUpB + 1
    AddLog "udB UpClick    v=" & CStr(udB.Value)
End Sub

Private Sub udB_DownClick()
    gDnB = gDnB + 1
    AddLog "udB DownClick  v=" & CStr(udB.Value)
End Sub

Private Sub udB_Change()
    gChB = gChB + 1
    AddLog "udB Change     v=" & CStr(udB.Value)
End Sub

Private Sub udC_UpClick()
    gUpC = gUpC + 1
    AddLog "udC UpClick    v=" & CStr(udC.Value)
End Sub

Private Sub udC_DownClick()
    gDnC = gDnC + 1
    AddLog "udC DownClick  v=" & CStr(udC.Value)
End Sub

Private Sub udC_Change()
    gChC = gChC + 1
    AddLog "udC Change     v=" & CStr(udC.Value)
End Sub

Private Sub udD_UpClick()
    gUpD = gUpD + 1
    AddLog "udD UpClick    v=" & CStr(udD.Value)
End Sub

Private Sub udD_DownClick()
    gDnD = gDnD + 1
    AddLog "udD DownClick  v=" & CStr(udD.Value)
End Sub

Private Sub udD_Change()
    gChD = gChD + 1
    AddLog "udD Change     v=" & CStr(udD.Value)
End Sub

Private Sub udE_UpClick()
    gUpE = gUpE + 1
    AddLog "udE UpClick    v=" & CStr(udE.Value)
End Sub

Private Sub udE_DownClick()
    gDnE = gDnE + 1
    AddLog "udE DownClick  v=" & CStr(udE.Value)
End Sub

Private Sub udE_Change()
    gChE = gChE + 1
    AddLog "udE Change     v=" & CStr(udE.Value)
End Sub

Private Sub cmdReset_Click()
    gSeq = 0
    gLog = ""
    gLastWritten = 0
    gUpA = 0
    gDnA = 0
    gChA = 0
    gUpB = 0
    gDnB = 0
    gChB = 0
    gUpC = 0
    gDnC = 0
    gChC = 0
    gUpD = 0
    gDnD = 0
    gChD = 0
    gUpE = 0
    gDnE = 0
    gChE = 0
    txtLog.Text = ""
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub tmGo_Timer()
    Dim s As String
    s = "SNAPSHOT (timer 300ms)  seq=" & CStr(gSeq) & vbCrLf
    s = s & "A v=" & CStr(udA.Value) & "  up=" & CStr(gUpA) & " dn=" & CStr(gDnA) & " ch=" & CStr(gChA) & "  txtA=" & txtA.Text & vbCrLf
    s = s & "B v=" & CStr(udB.Value) & "  up=" & CStr(gUpB) & " dn=" & CStr(gDnB) & " ch=" & CStr(gChB) & vbCrLf
    s = s & "C v=" & CStr(udC.Value) & "  up=" & CStr(gUpC) & " dn=" & CStr(gDnC) & " ch=" & CStr(gChC) & vbCrLf
    s = s & "D v=" & CStr(udD.Value) & "  up=" & CStr(gUpD) & " dn=" & CStr(gDnD) & " ch=" & CStr(gChD) & vbCrLf
    s = s & "E v=" & CStr(udE.Value) & "  up=" & CStr(gUpE) & " dn=" & CStr(gDnE) & " ch=" & CStr(gChE) & vbCrLf
    s = s & "---" & vbCrLf
    s = s & "A: Min0 Max10 V5 Inc1 Wrap=False Buddy=txtA" & vbCrLf
    s = s & "B: Min0 Max10 V10(top) Wrap=False" & vbCrLf
    s = s & "C: Min0 Max10 V10(top) Wrap=True" & vbCrLf
    s = s & "D: Min-20 Max20 V0 Inc5" & vbCrLf
    s = s & "E: Min0 Max10 V5 Horizontal Align=Left" & vbCrLf
    txtState.Text = s
    txtLog.Text = gLog
    ' 每拍无条件落盘。曾经的版本是"事件数没增长就不写", 结果正好漏掉要害那一档:
    ' **值变了、事件一条没到** 时文件一直停在 #1, 看上去就像"什么都没发生"。
    ' On Error Resume Next: 人工可能拿记事本开着这个文件, 共享冲突不许弹框。
    On Error Resume Next
    Open App.Path & "\UpManual.log" For Output As #1
    Print #1, s
    Print #1, "---- events ----"
    Print #1, gLog
    Close #1
End Sub
