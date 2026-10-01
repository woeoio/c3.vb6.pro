VERSION 5.00
Begin VB.Form C29ListView
   Caption         =   "C29ListView"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7200
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView ListView1
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6800
      _ExtentX        =   11991
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393216
   End
End
Attribute VB_Name = "C29ListView"
Option Explicit

' ai/029 C29-7: ListView 复刻的判据夹具 (数据面)。
' 覆盖: ColumnHeaders.Add (标题) / ListItems.Add (数据) / SubItems(i) 1 基 /
'       ListItem 与 ColumnHeader 的成员读写 / 两个集合的 Count 与 For Each /
'       ListView 自身的 View/GridLines 等属性 / 按 Key 与按下标取项 / Remove/Clear。
' 走 Test-Vbp (打完 Unload Me 拿 stdout 针)。
' 事件面 (ItemClick/ColumnClick) 另立, 见 LV20+ (需要 WM_NOTIFY, 无头环境点不了鼠标)。

Private Sub Form_Load()
    Dim col As ColumnHeader
    Dim itm As ListItem
    Dim e As Object
    Dim acc As String

    ' ---- 标题: ColumnHeaders.Add([index], [key], [text], [width]) ----
    Set col = ListView1.ColumnHeaders.Add(, "c1", "姓名", 1200)
    Debug.Print "LV1-COL-KEY=" & col.Key
    Debug.Print "LV2-COL-TEXT=" & col.Text
    Debug.Print "LV3-COL-IDX=" & col.Index
    Debug.Print "LV4-COLWIDTH=" & col.Width

    Set col = ListView1.ColumnHeaders.Add(, "c2", "部门", 1600)
    Debug.Print "LV5-COLCOUNT=" & ListView1.ColumnHeaders.Count

    ' ---- 数据: ListItems.Add([index], [key], [text]) ----
    Set itm = ListView1.ListItems.Add(, "r1", "张三")
    Debug.Print "LV6-ITEM-KEY=" & itm.Key
    Debug.Print "LV7-ITEM-TEXT=" & itm.Text
    Debug.Print "LV8-ITEM-IDX=" & itm.Index

    ' ---- SubItems: 1 基, SubItems(1) 就是第 2 列 ----
    itm.SubItems(1) = "销售部"
    Debug.Print "LV9-SUB1=" & itm.SubItems(1)

    Set itm = ListView1.ListItems.Add(, "r2", "李四")
    itm.SubItems(1) = "技术部"
    Debug.Print "LV10-ITEMCOUNT=" & ListView1.ListItems.Count

    ' ---- For Each 遍历行 (顺序 = 1 基下标序) ----
    acc = ""
    For Each e In ListView1.ListItems
        acc = acc & e.Text & "/" & e.SubItems(1) & ","
    Next
    Debug.Print "LV11-FOREACH=" & acc

    ' ---- For Each 遍历列 ----
    acc = ""
    For Each e In ListView1.ColumnHeaders
        acc = acc & e.Text & ","
    Next
    Debug.Print "LV12-COLS=" & acc

    ' ---- ListView 自身属性 (设计期 View=3 报表 / GridLines=True) ----
    Debug.Print "LV13-VIEW=" & ListView1.View
    Debug.Print "LV14-GRID=" & ListView1.GridLines

    ' ---- 取项: 按 Key 与按下标 ----
    Set itm = ListView1.ListItems("r2")
    Debug.Print "LV15-BYKEY=" & itm.Text
    Set itm = ListView1.ListItems(1)
    Debug.Print "LV16-ITEM1=" & itm.Text

    ' ---- 成员属性写: Selected ----
    itm.Selected = True
    Debug.Print "LV17-SEL=" & itm.Selected

    ' ---- 成员属性写: 列宽 ----
    Set col = ListView1.ColumnHeaders(1)
    col.Width = 900
    Debug.Print "LV18-COLW=" & ListView1.ColumnHeaders(1).Width

    ' ---- Remove / Clear ----
    ListView1.ListItems.Remove 1
    Debug.Print "LV19-AFTERRM=" & ListView1.ListItems.Count
    ListView1.ListItems.Clear
    Debug.Print "LV20-AFTERCLEAR=" & ListView1.ListItems.Count

    ' ---- 账 #128-b: 这六位是布尔属性, 三个面各钉一面 ----
    ' VB6 的答案是 True / Boolean / VT_BOOL(11); 归 Boolean 之前它们**根本没登记进**
    '类型表, 于是 `CStr(ListView1.GridLines)` 打的是 1。LV24 是**比较面**的鉴别针:
    '值必须是 -1 才与 True 相等 —— RTL 以前"写什么存什么", 设计期发的是 1, 这条就恒假。
    Debug.Print "LV23=" & CStr(ListView1.GridLines) & "/" & TypeName(ListView1.GridLines) & "/" & CStr(ListView1.MultiSelect)
    ListView1.CheckBoxes = True
    If ListView1.CheckBoxes = True Then Debug.Print "LV24=EQ" Else Debug.Print "LV24=NE"
    Debug.Print "LV25=" & CStr(ListView1.AllowColumnReorder) & "/" & CStr(ListView1.HideColumnHeaders)
    Unload Me
End Sub

' ---- C29-7 事件面 ----
' 这两个处理器**存在**就足以证明接线: cgen 看到它们才会生成 WM_NOTIFY 的
' -114 (LVN_ITEMACTIVATE) / -108 (LVN_COLUMNCLICK) 两个分支, 并把
' vb6_ListView_OnNotify → ListItemAt/ColumnHeaderAt → 回调 这条链接起来;
' 少了处理器就"分支不生成、回调符号不存在"(LNK2019)。
' 真实鼠标点击在无头环境里做不到 —— 与计划书 §五-5 "交互式使用兜底" 同一口径。
Private Sub ListView1_ItemClick(ByVal Item As ListItem)
    Debug.Print "LV21-ITEMCLICK=" & Item.Text
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    Debug.Print "LV22-COLCLICK=" & ColumnHeader.Text
End Sub
