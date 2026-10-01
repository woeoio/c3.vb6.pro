VERSION 5.00
Begin VB.Form CtrlStateForm 
   Caption         =   "CtrlStateForm"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4200
   Begin VB.Timer tProbe 
      Enabled         =   -1  'True
      Interval        =   50
      Left            =   3720
      Top             =   2400
   End
   Begin VB.Frame fr 
      Caption         =   "容器"
      Height          =   2100
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1800
      Begin VB.CheckBox cbIn 
         Caption         =   "内勾"
         Value           =   1  'Checked
         Height          =   300
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1500
      End
      Begin VB.CheckBox cbDisIn 
         Caption         =   "内禁"
         Enabled         =   0  'False
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1500
      End
      Begin VB.CheckBox cbHidIn 
         Caption         =   "内藏"
         Visible         =   0  'False
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1500
      End
      Begin VB.Label lbIn 
         Caption         =   "内标"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1500
      End
   End
   Begin VB.CheckBox cbOut 
      Caption         =   "外勾"
      Value           =   1  'Checked
      Height          =   300
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   1800
   End
   Begin VB.CheckBox cbDisOut 
      Caption         =   "外禁"
      Enabled         =   0  'False
      Height          =   300
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1800
   End
   Begin VB.CheckBox cbHidOut 
      Caption         =   "外藏"
      Visible         =   0  'False
      Height          =   300
      Left            =   2160
      TabIndex        =   7
      Top             =   840
      Width           =   1800
   End
   Begin VB.CheckBox cbNoTab 
      Caption         =   "跳格"
      TabStop         =   0  'False
      Height          =   300
      Left            =   2160
      TabIndex        =   12
      Top             =   2400
      Width           =   1800
   End
   Begin VB.CheckBox cbDef 
      Caption         =   "默认"
      Height          =   300
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   1800
   End
   Begin VB.OptionButton obDef 
      Caption         =   "默认选"
      Height          =   300
      Left            =   2160
      TabIndex        =   9
      Top             =   1560
      Width           =   1800
   End
   Begin VB.OptionButton obOn 
      Caption         =   "选中"
      Value           =   -1  'True
      Height          =   300
      Left            =   2160
      TabIndex        =   10
      Top             =   1920
      Width           =   1800
   End
   Begin VB.Label lbOut 
      Caption         =   "外标"
      Height          =   300
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1800
   End
End
Attribute VB_Name = "CtrlStateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 账 #125: 设计期 Enabled / Visible / CheckBox·OptionButton 的 Value 三件
' 以前两条创建路都没打到窗口上（发码里 BM_SETCHECK / EnableWindow / ShowWindow 各 0 次）。
' 判据纪律：**Visible 只能在窗体真显示之后问** —— Form_Load 里父窗还没 Show，
' IsWindowVisible 对任何控件都返回假（第一版探针就是这么假绿的）。所以读数放 Timer。
' 负向那一半同样是判据：没写过这三项的控件必须照旧（cbDef 不勾、不藏、不灰，
' fr 仍可见）—— 少了它，把"全都设一遍"的写错了也照样绿。

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Form_Load()
    ' 装载期问的是"窗口还没显示"那一态，只记不判（DS0 两条与 Timer 里的 DS5/DS6 对照）
    Debug.Print "DS0=" & CStr(cbHidIn.Visible) & CStr(cbHidOut.Visible)
End Sub

Private Sub tProbe_Timer()
    Dim v As Variant
    tProbe.Enabled = False
    Debug.Print "DS1=" & CStr(cbIn.Value) & CStr(cbOut.Value)          ' 设计期 Value=1（两条路）
    Debug.Print "DS2=" & CStr(cbDisIn.Enabled) & CStr(cbDisOut.Enabled) ' 设计期 Enabled=0；账 #124 后打 FalseFalse
    Debug.Print "DS3=" & CStr(cbHidIn.Visible) & CStr(cbHidOut.Visible) ' 设计期 Visible=0；账 #124 后打 FalseFalse
    Debug.Print "DS4=" & CStr(cbDef.Value) & CStr(cbDef.Visible) & CStr(cbDef.Enabled)
    Debug.Print "DS5=" & CStr(obDef.Value) & CStr(obOn.Value)           ' OptionButton 两个方向（VB6 是 Boolean，今仍数字档：见 029）
    Debug.Print "DS6=" & TF(fr.Visible) & TF(lbIn.Caption = "内标") & TF(lbOut.Caption = "外标")
    Debug.Print "DS7=" & CStr(obDef.Enabled)
    ' 账 #124: 控件布尔属性要在四个消费面都是 Boolean —— CStr 见 DS2/DS3、TypeName 与
    ' 装箱（VarType 11 = VT_BOOL）在这里、比较面在 DS6 的 TF(...)。VarType 11 是关键读数：
    ' 类型 oracle 归 Boolean 之后装箱才走 vb6_VariantBool（VB6 的 Enabled 装出来就是 VT_BOOL）。
    v = cbDisOut.Enabled
    Debug.Print "DS8=" & TypeName(cbDisOut.Enabled) & CStr(VarType(v)) & CStr(v)
    v = cbDef.Enabled
    Debug.Print "DS9=" & CStr(VarType(v)) & CStr(v)
    v = cbHidOut.Visible
    Debug.Print "DS10=" & CStr(VarType(v)) & CStr(v)
    ' 账 #128-b: OptionButton.Value 归 Boolean(读回 -1/0、写非 0 折回 BST_CHECKED)。
    ' obDef 设计期没写 Value、obOn 写了 True —— 两个方向各一面;DS12 是比较面。
    v = obOn.Value
    Debug.Print "DS11=" & CStr(obDef.Value) & "/" & TypeName(obOn.Value) & "/" & CStr(VarType(v))
    Debug.Print "DS12=" & TF(obOn.Value = True And obDef.Value = False)
    ' 账 #83(a)（C29-SL-r）: VB6 的 TabStop 默认是 True，而以前**两条创建路都不立 WS_TABSTOP**
    ' （029 的 C29-SL-r-0：顶层按钮、Frame 里的按钮与文本框一律读回 0）。ST1 一次问四枚 ——
    ' 顶层 / 容器里 / 容器里那枚 Label（拿不到焦点，不该立）/ 显式写了 `TabStop = 0` 的那枚。
    ' 这一格问的是窗口本身（`vb6_GetTabStop` 读 GWL_STYLE），没有自存 ⇒ 不存在"读我们存的数"那种自洽假绿。
    ' ST2 是读侧那半（#124 同族）：登记成 Boolean 之前 VarType 是 3（Integer）、TypeName 也是整数那一头。
    Debug.Print "ST1-tab=" & CStr(cbOut.TabStop) & "/" & CStr(cbIn.TabStop) & "/" _
        & CStr(lbIn.TabStop) & "/" & CStr(cbNoTab.TabStop)
    Debug.Print "ST2-bool=" & CStr(VarType(cbOut.TabStop)) & "/" & TypeName(cbOut.TabStop)
    Debug.Print "CTRLSTATE-DONE"
    Unload Me
End Sub
