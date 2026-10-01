VERSION 5.00
Begin VB.Form TvfForm 
   Caption         =   "TvfForm"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "TvfForm"
   ScaleHeight     =   3200
   ScaleWidth      =   6000
   Begin VB.TreeView tv1 
      CheckBoxes      =   -1  'True
      Height          =   1500
      HideSelection   =   0   'False
      HotTracking     =   -1  'True
      Indentation     =   300
      Left            =   240
      LineStyle       =   1   'RootLines
      TabIndex        =   0
      Top             =   240
      Width           =   2100
   End
   Begin VB.TreeView tv2 
      Height          =   1500
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2100
   End
   Begin VB.Timer evtTimer 
      Interval        =   200
      Left            =   240
      Top             =   1920
   End
End
Attribute VB_Name = "TvfForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ai/029 C29-8a: TreeView 的标量属性面走原生 SysTreeView32（不加载 MSCOMCTL.OCX）。
' 改之前这枚控件的**窗口本来就建得出来**（controlTypeToWin32Class 早就有格），缺的是属性面：
' cgen 的读写表里 TreeView 零格 ⇒ 属性一律落通用兜底 vb6_ComGetStringProp(裸 HWND, "…")，
' 实测读回空串、写进去静默丢（029 §九 那四条前置读数）。本批补五条：
'   LineStyle / Indentation / CheckBoxes / HotTracking / HideSelection
' 判据口径：这四条样式类的属性**真值就是窗口样式位**（RTL getter 读 GWL_STYLE），所以
' "设计期初值落位"、"运行期可逆赋值"、"默认值"三类读数都问的是同一个窗口；
' Indentation 的默认那条（TV7）问的是 TVM_GETINDENT —— 句柄为空只会读到 0，读到正数
' 就是"消息真打进了控件"，与本线"窗口存在性靠读数证"的口径一致。
' Nodes / Style / LabelEdit / Sorted 与事件不在本批 (8a)：Nodes/Node 那一大半在 8b 接上
' (TV13..TV27)，Style / LabelEdit / Sorted / NodeClick 还欠着。

' C29-8c: 事件面 (NodeClick / Expand / Collapse) 走父窗的 WM_NOTIFY，判据靠 RTL 的
' Sim* 助手程序化发**真通知** —— 无头环境点不了鼠标，而直接调 handler 会绕开整条派发链。
' ⚠ Sim* 是**判据专用**助手 (与 C29-4 的 StatusBar.SimClick 同一先例)，不对应任何 VB6 语义。
Private gClicks As Long
Private gHitB As Long
Private gHitC As Long
Private gExpands As Long
Private gCollapses As Long
Private gExpText As String
Private gOther As Long

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Sub Form_Load()

    ' --- 1. 设计期五值落位 (tv1 的 .frm 里全写过) ---
    Debug.Print "TV1=" & TF(tv1.CheckBoxes <> 0)
    Debug.Print "TV2=" & TF(tv1.HotTracking <> 0)
    Debug.Print "TV3=" & TF(tv1.LineStyle = 1)
    Debug.Print "TV4=" & TF(tv1.Indentation = 300)
    Debug.Print "TV5=" & TF(tv1.HideSelection = 0)

    ' --- 2. 没写过的控件走 VB6 默认 (复选框/跟踪关、根层不画线、失焦藏选中) ---
    Debug.Print "TV6=" & TF(tv2.CheckBoxes = 0 And tv2.HotTracking = 0 _
                            And tv2.LineStyle = 0 And tv2.HideSelection <> 0)

    ' --- 3. 默认缩进是从真窗口问出来的 (TVM_GETINDENT) ---
    Debug.Print "TV7=" & TF(tv2.Indentation > 0)

    ' --- 4. 运行期赋值 + 反向可逆 ---
    tv2.CheckBoxes = True
    tv2.LineStyle = 1
    tv2.HotTracking = True
    Debug.Print "TV8=" & TF(tv2.CheckBoxes <> 0 And tv2.LineStyle = 1 And tv2.HotTracking <> 0)
    tv2.CheckBoxes = False
    tv2.LineStyle = 0
    tv2.HotTracking = False
    tv2.HideSelection = False
    Debug.Print "TV9=" & TF(tv2.CheckBoxes = 0 And tv2.LineStyle = 0 _
                            And tv2.HotTracking = 0 And tv2.HideSelection = 0)

    ' --- 5. Indentation 的 VB6 侧口径是缇: 往返要原样读回, 超界也不回吐像素值 ---
    tv2.Indentation = 400
    Debug.Print "TV10=" & TF(tv2.Indentation = 400)
    tv2.Indentation = 100000
    Debug.Print "TV11=" & TF(tv2.Indentation = 100000)

    ' --- 6. 通用属性面没被这五条抢走 ---
    tv1.Visible = False
    Debug.Print "TV12=" & TF(tv1.Visible = False And tv1.Height = 1500)

    ' ============================================================
    ' --- 7. C29-8b: Nodes 集合与 Node 对象 (真 IDispatch, 见 vb6forms_memberobj.c) ---
    '     结构 (父子/兄弟) 现问原生树，字符串住这张表 —— 两边都从同一个窗口出发,
    '     所以"读数对"就等于"屏幕上对"。集合序 = 插入序 (029 §九 记的口径)。
    Dim ndA As Object, ndB As Object, ndC As Object
    Set ndA = tv1.Nodes.Add(, , "a", "根甲")
    Set ndB = tv1.Nodes.Add("a", 4, "b", "子乙")      ' relative 按 Key, 4 = tvwChild
    Set ndC = tv1.Nodes.Add(, , "c", "根丙")

    ' --- 8. Count / Add 返回的对象带 Index ---
    Debug.Print "TV13=" & TF(tv1.Nodes.Count = 3)
    Debug.Print "TV14=" & TF(ndA.Index = 1 And ndB.Index = 2 And ndC.Index = 3)

    ' --- 9. 下标取与按 Key 取 (同一条 Item 两种实参形态) ---
    '     账 #88 收线后这里**直接**比字符串成员。以前 cgen 在比较上下文里把链上的字符串成员
    '     按数值解包 (发 vb6_CStrLong(vb6_ComGetIntProp(o, L"Text"))) ⇒ 拿 BSTR 指针当数字
    '     去和字面量比 ⇒ 恒 False，判据只能先把值取进局部变量绕过去 (TV34 仍留那条路径)。
    '     TV37 钉的是选择性: 同一条链上的**数值**成员 (Index/Children) 必须照旧走 int 档，
    '     两档一起改是错的；TV36 钉的是"上面那些 Y 不是恒真"。
    Dim ndIx As Object, ndKy As Object
    Set ndIx = tv1.Nodes(2)
    Set ndKy = tv1.Nodes("b")
    Debug.Print "TV15=" & TF(ndIx.Text = "子乙" And ndKy.Index = 2)
    Debug.Print "TV35=" & TF(tv1.Nodes(2).Text = "子乙" And tv1.Nodes("b").Key = "b")
    Debug.Print "TV36=" & TF(tv1.Nodes(2).Text = "根甲")
    Debug.Print "TV37=" & TF(tv1.Nodes(2).Index = 2 And tv1.Nodes(1).Children = 1)
    Dim sIx As String
    sIx = ndIx.Text
    Debug.Print "TV34=" & TF(sIx = "子乙")

    ' --- 10. 导航读数全问原生树 ---
    Debug.Print "TV16=" & TF(ndA.Child = 2 And ndA.Children = 1)
    Debug.Print "TV17=" & TF(ndB.Parent = 1 And ndB.Next = 0 And ndC.Previous = 1)
    Debug.Print "TV18=" & TF(ndB.Root = 1 And ndC.Root = 3)

    ' --- 11. 写回去：Text 改完两边同步 (表 + TVM_SETITEMW) ---
    ndA.Text = "改名甲"
    ndA.Tag = "标签甲"
    Dim ndReread As Object
    Dim sRe As String, tRe As String
    Set ndReread = tv1.Nodes(1)          ' 换一枚对象再读: 证的是**控件侧**的状态, 不是变量里存的串
    sRe = ndReread.Text
    tRe = ndReread.Tag
    Debug.Print "TV19=" & TF(sRe = "改名甲" And tRe = "标签甲")

    ' --- 12. Checked 住在原生 state image 里 (tv1 设计期就 CheckBoxes=True) ---
    ndA.Checked = True
    ndC.Checked = False
    Debug.Print "TV20=" & TF((ndA.Checked <> 0) And (ndC.Checked = 0))

    ' --- 13. Expanded: 先关后开, 读数问 TVIS_EXPANDED ---
    ndA.Expanded = True
    Debug.Print "TV21=" & TF(ndA.Expanded <> 0)
    ndA.Expanded = False
    Debug.Print "TV22=" & TF(ndA.Expanded = 0)

    ' --- 14. EnsureVisible 把父节点撑开 (TVM_ENSUREVISIBLE 的真行为, 不是表里的位) ---
    ndB.EnsureVisible
    Debug.Print "TV23=" & TF(ndA.Expanded <> 0)

    ' --- 15. For Each 走 _NewEnum, 顺序与集合序同 ---
    Dim e As Object
    Dim acc As String
    acc = ""
    For Each e In tv1.Nodes
        acc = acc & e.Index & ":" & e.Text & "/"
    Next
    Debug.Print "TV24=" & TF(acc = "1:改名甲/2:子乙/3:根丙/")

    ' --- 16. Remove 按 Key 删的是**整棵子树** (原生删父带走子, 表要跟住) ---
    tv1.Nodes.Remove "a"
    Dim ndLeft As Object
    Dim kLeft As String
    Set ndLeft = tv1.Nodes(1)
    kLeft = ndLeft.Key
    Debug.Print "TV25=" & TF(tv1.Nodes.Count = 1 And kLeft = "c")

    ' --- 17. Clear 清表 ---
    tv1.Nodes.Clear
    Debug.Print "TV26=" & TF(tv1.Nodes.Count = 0)

    ' --- 18. 集合是**独立于控件属性面**的一条路: 标量读数没被 Nodes 抢走 ---
    Debug.Print "TV27=" & TF(tv1.CheckBoxes <> 0 And tv2.Nodes.Count = 0)

    ' DONE 与 Unload 挪到 evtTimer_Timer —— 事件必须在窗体载入完之后才派发得动
End Sub

' 派发链的触发点: Form_Load 阶段被 block events during form init 拦掉，Form_Activate 在无头
' 会话里永远不来 (GA #109 实测) => 照 C29-4/P20-42 的先例放 Timer。
Private Sub evtTimer_Timer()
    Static done As Integer
    If done Then Exit Sub
    done = 1

    ' Form_Load 的最后两栏 (TV25/TV26) 把表清干净了，所以事件判据自己重建目标树 ——
    ' 顺便这条也是"通知换算问的是活着的表"的证据: 表空的时候 Sim 打不进 handler。
    tv1.Nodes.Add , , "a", "根甲"
    tv1.Nodes.Add "a", 4, "b", "子乙"
    tv1.Nodes.Add , , "c", "根丙"

    ' 计数一律问**增量**: 光 Form_Load 那半截就把真通知发出来过 (下面 TV33 钉它)，
    ' 拿绝对值比就会把两件事混在一起。
    Dim baseClick As Long, baseExp As Long, baseCol As Long
    baseClick = gClicks
    baseExp = gExpands
    baseCol = gCollapses

    ' 下面 TV33 钉的是"控件自己发的通知也走完了同一条派发链":
    ' 实测到这里 gExpands 已经是 2 (TV21 那条 Expanded=True 与 TV23 那条
    ' EnsureVisible 各撑开一次，原生就发了 TVN_ITEMEXPANDEDW)。
    Debug.Print "TV33=" & TF(gExpands >= 2)

    tv1.SimNodeClick 2
    tv1.SimNodeClick 3
    Debug.Print "TV28=" & TF(gClicks - baseClick = 2)
    Debug.Print "TV29=" & TF(gHitB = 1 And gHitC = 1)
    ' 通知是从 tv1 发的：另一枚树不该收到任何东西
    Debug.Print "TV30=" & TF(gOther = 0 And tv2.Nodes.Count = 0)

    tv1.SimExpand 1, True
    tv1.SimExpand 1, False
    Debug.Print "TV31=" & TF(gExpands - baseExp = 1 And gCollapses - baseCol = 1)
    Debug.Print "TV32=" & TF(gExpText = "根甲")

        Debug.Print "TV38=" & CStr(tv1.CheckBoxes) & "/" & CStr(tv1.HideSelection) & "/" & TypeName(tv1.HotTracking)
Debug.Print "TREEVIEW-DONE"
    Unload Me
End Sub

' handler 里**真读** Node 对象的成员: 证的是造出来的那枚对象指向被点的那一格，
' 而不是 handler 被调了一次就算数。
Private Sub tv1_NodeClick(ByVal Node As Node)
    Dim sK As String
    gClicks = gClicks + 1
    sK = Node.Key
    If sK = "b" Then gHitB = 1
    If sK = "c" Then gHitC = 1
End Sub

Private Sub tv1_Expand(ByVal Node As Node)
    Dim sT As String
    gExpands = gExpands + 1
    sT = Node.Text
    gExpText = sT
End Sub

Private Sub tv1_Collapse(ByVal Node As Node)
    gCollapses = gCollapses + 1
End Sub

' tv2 一枚对照 handler: 它不该被 tv1 的通知打进来 (TV30 问的就是这个 gOther)
Private Sub tv2_NodeClick(ByVal Node As Node)
    gOther = gOther + 1
End Sub
