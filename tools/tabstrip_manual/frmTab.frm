VERSION 5.00
Begin VB.Form frmTab 
   Caption         =   "C3 TabStrip 手动验证 (P0 / C30-2)"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4300
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   7585
      _Version        =   393216
      HotTracking     =   -1  'True
      MultiRow        =   0   'False
      ShowTips        =   -1  'True
   End
   Begin VB.PictureBox picA 
      BackColor       =   &H000000FF&
      Height          =   3300
      Left            =   450
      ScaleHeight     =   3240
      ScaleWidth      =   5740
      TabIndex        =   1
      Top             =   870
      Width           =   5800
   End
   Begin VB.PictureBox picB 
      BackColor       =   &H0000C000&
      Height          =   3300
      Left            =   450
      ScaleHeight     =   3240
      ScaleWidth      =   5740
      TabIndex        =   2
      Top             =   870
      Width           =   5800
   End
   Begin VB.PictureBox picC 
      BackColor       =   &H00C00000&
      Height          =   3300
      Left            =   450
      ScaleHeight     =   3240
      ScaleWidth      =   5740
      TabIndex        =   3
      Top             =   870
      Width           =   5800
   End
   Begin VB.Label lblInfo 
      Caption         =   "启动后应显示 tabs=3 sel=1 cap=常规 clicks=0"
      Height          =   420
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   4380
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "加一页"
      Height          =   420
      Left            =   4740
      TabIndex        =   5
      Top             =   4680
      Width           =   760
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "删当前页"
      Height          =   420
      Left            =   5520
      TabIndex        =   6
      Top             =   4680
      Width           =   840
   End
End
Attribute VB_Name = "frmTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' C30-2 手动验证窗体（不进 tests/，不进语料/清单扫描）。
'
' 验证点（照 ai/内置控件/TabStrip 控件（选项卡条）.md §5 的形状写）：
'   1. Form_Load 里连续 Tabs.Add（index 省略、key 与 caption 都给）→ 标签条出现 3 页；
'   2. 三个 PictureBox 按 ClientLeft/ClientTop/ClientWidth/ClientHeight 贴到页区上
'      （TabStrip **不是容器**，页内容是自己摆的）；
'   3. 点标签 → TabStrip1_Click 把 PictureBox 的 Visible 翻过去，并刷新 lblInfo；
'   4. lblInfo 同时回读 Tabs.Count / SelectedItem.Index / Tabs(n).Caption 三处。
'
' ⚠ 已知缺口（下一刀补）：链式读**直接当内建函数实参**时（`CStr(TabStrip1.SelectedItem.Index)`）
'   仍走 COM 晚绑定 ⇒ 真跑会弹 Run-time error 438。本样例统一先读进 Long 变量再用，
'   这也是 VB6 工程里更常见的写法。

Dim nSel As Long
Dim nClicks As Long

Private Sub Form_Load()
    TabStrip1.Tabs.Add , "t1", "常规"
    TabStrip1.Tabs.Add , "t2", "视图"
    TabStrip1.Tabs.Add , "t3", "高级"
    LayoutPages
    ShowPage
End Sub

Private Sub LayoutPages()
    picA.Left = TabStrip1.Left + TabStrip1.ClientLeft
    picA.Top = TabStrip1.Top + TabStrip1.ClientTop
    picA.Width = TabStrip1.ClientWidth
    picA.Height = TabStrip1.ClientHeight
    picB.Left = picA.Left
    picB.Top = picA.Top
    picB.Width = picA.Width
    picB.Height = picA.Height
    picC.Left = picA.Left
    picC.Top = picA.Top
    picC.Width = picA.Width
    picC.Height = picA.Height
End Sub

Private Sub ShowPage()
    Dim nCnt As Long
    Dim nCur As Long
    Dim sCap As String
    nCnt = TabStrip1.Tabs.Count
    nCur = TabStrip1.SelectedItem.Index
    If nCur < 1 Then nCur = 1
    If nCur > nCnt Then nCur = nCnt
    nSel = nCur
    picA.Visible = (nCur = 1)
    picB.Visible = (nCur = 2)
    picC.Visible = (nCur = 3)
    sCap = TabStrip1.Tabs(nCur).Caption
    lblInfo.Caption = "tabs=" & CStr(nCnt) _
        & "  sel=" & CStr(nCur) _
        & "  cap=" & sCap _
        & "  clicks=" & CStr(nClicks)
End Sub

Private Sub TabStrip1_Click()
    nClicks = nClicks + 1
    ShowPage
End Sub

Private Sub cmdAdd_Click()
    Dim nCnt As Long
    nCnt = TabStrip1.Tabs.Count
    TabStrip1.Tabs.Add , "t" & CStr(nCnt + 1), "新页" & CStr(nCnt + 1)
    ShowPage
End Sub

Private Sub cmdRemove_Click()
    Dim nCnt As Long
    nCnt = TabStrip1.Tabs.Count
    If nCnt > 1 Then
        TabStrip1.Tabs.Remove nSel
        ShowPage
    End If
End Sub
