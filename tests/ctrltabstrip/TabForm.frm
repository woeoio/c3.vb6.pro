VERSION 5.00
Begin VB.Form TabForm 
   Caption         =   "TabForm"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "TabForm"
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3000
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   5292
      _Version        =   393216
      HotTracking     =   -1  'True
      MultiRow        =   0   'False
      ShowTips        =   -1  'True
   End
   Begin VB.Timer tGo 
      Enabled         =   -1  'True
      Interval        =   60
      Left            =   4200
      Top            =   3120
   End
End
Attribute VB_Name = "TabForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' C30-2 TabStrip 的无头判据（不进 tools/，与 tests/ctrlupdown 同构：Timer 起、Debug.Print 出）。
'
' 这一件只钉**属性面 + Tabs 集合**：标签条的点击/切换要真鼠标（tools/tabstrip_manual
' 那一份 + .build/probe/verify_tab.ps1），这里不碰 _Click —— 免得“编得过、没人送消息”的绿。
'
' 判据面（全部用 VB6 原样语法，不做“先落变量”那种绕行）：
'   · 属性面（TS1..TS9）与 Tabs 集合（TS5/TS6/TS10）；
'   · 集合链**当内建函数实参 / 子过程实参**（TS11/TS12）—— 改前这两形走 COM 晚绑定，真跑 438；
'   · 写侧 Caption/Key/Image/Selected 与 SelectedItem.Index（TS13..TS15）；
'   · Remove 的选中项语义（TS16/TS17）。
'
' 1 基：Tabs(1..Count)、SelectedItem.Index(1..Count，无选中=0)。

Private Sub Form_Load()
    TabStrip1.Tabs.Add , "k1", "常规"
    TabStrip1.Tabs.Add , "k2", "视图"
    TabStrip1.Tabs.Add , "k3", "高级"
End Sub

Private Sub ShowIdx(ByVal i As Long)
    Debug.Print "TS12-sub=" & CStr(i)
End Sub

Private Sub tGo_Timer()
    Dim n As Long
    Dim n2 As Long
    Dim c1 As String
    Dim c2 As String
    Dim c3 As String
    Dim k1 As String
    tGo.Enabled = False

    n = TabStrip1.Tabs.Count
    Debug.Print "TS1-tabs=" & CStr(n)

    ' 设计期没给 SelectedItem，原生控件插第一项即选第一项 ⇒ 1 基口径下是 1。
    n = TabStrip1.SelectedItem.Index
    Debug.Print "TS2-sel=" & CStr(n)

    c1 = TabStrip1.Tabs(1).Caption
    c2 = TabStrip1.Tabs(2).Caption
    c3 = TabStrip1.Tabs(3).Caption
    ' 判据一律落成 ASCII：控制台是 GBK，中文 needle 在别的 locale 上必红；
    ' 所以这里钉的是**相等**（真值/假值）与长度，不是“字符串里出现过”。
    Debug.Print "TS3-cap=" & CStr(c1 = "常规") & "/" & CStr(c2 = "视图") & "/" & CStr(c3 = "高级")
    Debug.Print "TS3-len=" & CStr(Len(c1)) & "/" & CStr(Len(c2)) & "/" & CStr(Len(c3))

    k1 = TabStrip1.Tabs(1).Key
    Debug.Print "TS4-key=" & k1

    ' Tabs.Add 的 index 省略 ⇒ 追加到末尾。
    TabStrip1.Tabs.Add , "k4", "新页4"
    n = TabStrip1.Tabs.Count
    c1 = TabStrip1.Tabs(4).Caption
    Debug.Print "TS5-add=" & CStr(n) & "/" & CStr(c1 = "新页4")

    ' Tabs.Remove 2：删掉“视图”，后面的往前顶（3→2、4→3）。
    TabStrip1.Tabs.Remove 2
    n = TabStrip1.Tabs.Count
    c2 = TabStrip1.Tabs(2).Caption
    c3 = TabStrip1.Tabs(3).Caption
    Debug.Print "TS6-rm=" & CStr(n) & "/" & CStr(c2 = "高级") & "/" & CStr(c3 = "新页4")

    ' 类型面：Count/Index 是 Long，Caption/Key 是 String（Properties 窗口的判定口径一致）。
    n = TabStrip1.Tabs.Count
    c1 = TabStrip1.Tabs(1).Caption
    Debug.Print "TS7-type=" & TypeName(n) & "/" & TypeName(c1)

    ' 几何面：Client* 是标签**页区**，不是整窗（TabStrip 不是容器，页内容自己摆）。
    n = TabStrip1.ClientWidth
    n2 = TabStrip1.ClientHeight
    Debug.Print "TS8-client=" & CStr(n) & "x" & CStr(n2)

    ' 行为面三枚：设计期 MultiRow=False / HotTracking=True / ShowTips=True。
    Debug.Print "TS9-flags=" & CStr(TabStrip1.MultiRow) & "/" & CStr(TabStrip1.HotTracking) & "/" & CStr(TabStrip1.ShowTips)

    ' ---- 集合链当**内建函数实参** / 子过程实参（改前这两形走 COM 晚绑定 ⇒ 真跑 438）----
    ' 这一格是 C30-2 收尾的硬判据：不再“先落局部变量再用”。
    Debug.Print "TS11-arg=" & CStr(TabStrip1.Tabs.Count) & "/" & CStr(TabStrip1.SelectedItem.Index)
    ShowIdx TabStrip1.SelectedItem.Index

    ' ---- 写侧：Tabs(i) 的 Caption / Key / Image ----
    TabStrip1.Tabs(1).Caption = "改1"
    TabStrip1.Tabs(1).Key = "kk1"
    TabStrip1.Tabs(1).Image = 0
    c1 = TabStrip1.Tabs(1).Caption
    k1 = TabStrip1.Tabs(1).Key
    n = TabStrip1.Tabs(1).Image
    Debug.Print "TS13-wcap=" & CStr(c1 = "改1") & "/" & CStr(k1 = "kk1") & "/" & CStr(n)

    ' ---- 写侧：Tabs(i).Selected = True / SelectedItem.Index = n ----
    TabStrip1.Tabs(2).Selected = True
    n = TabStrip1.SelectedItem.Index
    n2 = TabStrip1.Tabs(2).Index
    Debug.Print "TS14-wsel=" & CStr(n) & "/" & CStr(n2)
    TabStrip1.SelectedItem.Index = 1
    n = TabStrip1.SelectedItem.Index
    Debug.Print "TS15-selset=" & CStr(n)

    ' ---- Remove 的选中项语义：选中第 3 页，删第 1 页 ⇒ 用户手上那页不变（序号 3→2）----
    ' 注：此刻三页的标题已变成 [改1, 高级, 新页4]（TS13 改过第 1 页），所以“不变的那页”是“新页4”。
    TabStrip1.SelectedItem.Index = 3
    TabStrip1.Tabs.Remove 1
    n = TabStrip1.Tabs.Count
    n2 = TabStrip1.SelectedItem.Index
    c1 = TabStrip1.SelectedItem.Caption
    Debug.Print "TS16-rmkeep=" & CStr(n) & "/" & CStr(n2) & "/" & CStr(c1 = "新页4")
    ' 删的就是当前页：顶上来那一页就地接管选中（VB6 的“删后选其下一位”）
    TabStrip1.Tabs.Remove 2
    n = TabStrip1.Tabs.Count
    n2 = TabStrip1.SelectedItem.Index
    Debug.Print "TS17-rmcur=" & CStr(n) & "/" & CStr(n2)

    ' Tabs.Clear：清空后 Count=0、SelectedItem.Index 回 0（VB6 的“无选中”）。
    TabStrip1.Tabs.Clear
    n = TabStrip1.Tabs.Count
    n2 = TabStrip1.SelectedItem.Index
    Debug.Print "TS10-clear=" & CStr(n) & "/" & CStr(n2)

    Debug.Print "TABSTRIP-DONE"
    Unload Me
End Sub
