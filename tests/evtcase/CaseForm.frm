VERSION 5.00
Begin VB.Form CaseForm
   Caption         =   "EvtCase"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "CaseForm"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdRun
      Caption         =   "run"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.CheckBox chkOpt
      Caption         =   "opt"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   180
      Width           =   1455
   End
   Begin VB.CommandButton cmdSame
      Caption         =   "same"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer tmrStep
      Enabled         =   -1  'True
      Interval        =   150
   End
End
Attribute VB_Name = "CaseForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #190 的判据夹具。
'   发码里事件臂的**存在性**问的是 symTab_.lookup(控件名 & "_Click") —— 大小写无关, 命中;
'   可发出去的**函数名**却是拿控件的设计期拼写现拼的, 而过程定义用的是 Sub 自己的名字。
'   VB6 大小写不敏感、C 敏感 ⇒ 只要这两处拼写只差大小写, 产物就是"声明了一个没人定义的函数",
'   链接期 LNK2019 (实测 Charts 2020/ucChartBar 的 Form1: ChkAxisY / ChkAxisy_Click 与
'   cboLabelsPositions / CboLabelsPositions_Click 两枚, 正好 2 个无法解析的外部符号)。
'   本夹具把这两形各摆一枚 (cmdRun/CmdRun_Click、chkOpt/ChkOpt_Click), 再摆一枚**拼写一致**的
'   cmdSame/cmdSame_Click 当证人 —— 证人排除的是"BM_CLICK 这条路自己没驱动起来"这种假红。
'   自驱: Timer 第一拍对三枚控件发 BM_CLICK (按钮自己会向父窗发 BN_CLICKED, 不打真实鼠标,
'   见账 #79 那条级联偏移的教训), 第二拍汇总计数后自退。

Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As Long

Private mStep As Integer
Private nRun As Integer
Private nChk As Integer
Private nSame As Integer

Private Sub Form_Load()
    Debug.Print "EC00-LOAD"
End Sub

Private Sub tmrStep_Timer()
    Const BM_CLICK As Long = &HF5
    If mStep = 0 Then
        Call SendMessage(cmdRun.hwnd, BM_CLICK, 0, 0)
        Call SendMessage(chkOpt.hwnd, BM_CLICK, 0, 0)
        Call SendMessage(cmdSame.hwnd, BM_CLICK, 0, 0)
    ElseIf mStep = 2 Then
        Debug.Print "EC-CNT run=" & nRun & " chk=" & nChk & " same=" & nSame
        tmrStep.Enabled = False
        Debug.Print "EC-DONE"
        Unload Me
    End If
    mStep = mStep + 1
End Sub

' 过程名与控件名只差首字母大小写 —— 这两条就是本账要拦的形状。
Private Sub CmdRun_Click()
    nRun = nRun + 1
    Debug.Print "EC01-RUN-CLICK"
End Sub

Private Sub ChkOpt_Click()
    nChk = nChk + 1
    Debug.Print "EC02-CHK-CLICK chk=" & CStr(chkOpt.Value)
End Sub

' 证人: 拼写完全一致, 修复前后都该响 (它排除"夹具自己没驱动起来")。
Private Sub cmdSame_Click()
    nSame = nSame + 1
    Debug.Print "EC03-SAME-CLICK"
End Sub
