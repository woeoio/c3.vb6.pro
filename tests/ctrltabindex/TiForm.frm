VERSION 5.00
Begin VB.Form TiForm 
   Caption         =   "TiForm"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3600
   LinkTopic       =   "TiForm"
   ScaleHeight     =   2400
   ScaleWidth      =   3600
   Begin VB.Label lblTop 
      Caption         =   "tabindex"
      Height          =   255
      Left            =   120
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdFive 
      Caption         =   "5"
      Height          =   400
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   700
   End
   Begin VB.TextBox txtTwo 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   900
   End
   Begin VB.PictureBox picBox 
      Height          =   900
      Left            =   2160
      ScaleHeight     =   840
      ScaleMode       =   3  'String
      ScaleWidth      =   1212
      Top             =   480
      Width           =   1300
   End
   Begin VB.Frame fr 
      Caption         =   "fr"
      Height          =   1000
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1900
      Begin VB.CommandButton cmdIn0 
         Caption         =   "I0"
         Height          =   300
         Left            =   120
         TabIndex        =   0
         Top             =   300
         Width           =   600
      End
      Begin VB.Label lblIn 
         Caption         =   "in"
         Height          =   255
         Left            =   840
         Top             =   300
         Width           =   700
      End
      Begin VB.CommandButton cmdIn9 
         Caption         =   "I9"
         Height          =   300
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdSix 
      Caption         =   "6"
      Height          =   400
      Left            =   2160
      Top             =   1560
      Width           =   700
   End
   Begin VB.Timer t1 
      Enabled         =   -1   'True
      Interval        =   60
      Left            =   3100
      Top             =   120
   End
End
Attribute VB_Name = "TiForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #160 的夹具：`.frm` 里写着的 `TabIndex` 要能在运行期读回来。
' RTL 那一对 (`vb6_GetTabIndex` / `vb6_SetTabIndex`，存窗口属性 `VB6_TabIndex`) 早就在，
' 缺的只有"创建时没人发"这一刀 ⇒ 改之前**每一枚控件都读 0**（getter 没存过时答 0）。
' 三档判据各管一件事：
'   TI-explicit = 写了的照发（5 / 2 / 1 —— 故意与声明顺序相反，"照创建顺序"会全读成别的数）
'   TI-in       = **每个父窗自己从 0 编号**（框架里那两枚读 0 / 9，与窗体那一串互不干扰）
'   TI-fallback = 没写的用**同一父窗内的声明序号**兜底：picBox=3、lblIn=1、cmdSix=5。
'                 兜底必须可分辨 —— cmdSix 是窗体那一串里第 5 个声明的（lblTop 0、cmdFive 1、
'                 txtTwo 2、picBox 3、fr 4、cmdSix 5、t1 6），读回 5 才说明兜的是声明序而不是
'                 "没写就当 0"（那样会有好几枚同时声称 0）。
'   TI-runtime  = 运行期赋值照样生效（setter 那条路没被发码期那一句盖死）。
Private gTick As Long

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Log1(ByVal s As String)
    Dim h As Long
    h = FreeFile
    Open App.Path & "\tabindex.log" For Append As #h
    Print #h, s
    Close #h
    Debug.Print s
End Sub

Private Sub t1_Timer()
    gTick = gTick + 1
    If gTick <> 1 Then Exit Sub
    t1.Enabled = False
    Log1 "TI-explicit=" & CStr(cmdFive.TabIndex) & "/" & CStr(txtTwo.TabIndex) & "/" & CStr(fr.TabIndex)
    Log1 "TI-in=" & CStr(cmdIn0.TabIndex) & "/" & CStr(cmdIn9.TabIndex)
    Log1 "TI-fallback=" & CStr(picBox.TabIndex) & "/" & CStr(lblIn.TabIndex) & "/" & CStr(cmdSix.TabIndex)
    cmdSix.TabIndex = 40
    Log1 "TI-runtime=" & CStr(cmdSix.TabIndex) & "/" & TF(cmdSix.TabIndex = 40)
    Log1 "TABINDEX-DONE"
    Unload Me
End Sub
