VERSION 5.00
Begin VB.Form HmForm 
   Caption         =   "HmForm"
   ClientHeight    =   3200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   LinkTopic       =   "HmForm"
   ScaleHeight     =   3200
   ScaleWidth      =   6000
   Begin VB.TreeView tvA 
      Height          =   1500
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2100
   End
   Begin VB.TreeView tvB 
      Height          =   1500
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2100
   End
   Begin RichTextLib.RichTextBox rt1 
      Height          =   900
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   2100
   End
   Begin VB.Timer evtTimer 
      Interval        =   200
      Left            =   4560
      Top             =   2760
   End
End
Attribute VB_Name = "HmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' M3 第三百五十一刀: OCX 宿主属性面的可观察夹具。
'
' 靶子不是"日志里不再出现 C3_HOST_MEMBER_MISS", 而是**改完之后运行期读得回、
' 窗口本身也确实变了**。所以每条读数都问两种证人:
'   ① 对象本身 (晚绑定读 `o.X` / 直接读 `tvA.X`)
'   ② 原生窗口的样式位或消息返回值 (GetWindowLongW / EM_GETMARGINS / GetWindowTextW)
' 两证人必须一致 —— 只改一张 RTL 自存表的实现过不了 ①+②, 只改样式位而读表没跟上
' 也过不了 `o.X` 那一问。tvB 是隔离证人: 对 tvA 的所有操作不许漏到 tvB。

Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const GW_HWNDNEXT As Long = 2
Private Const GW_HWNDPREV As Long = 3
Private Const WS_BORDER As Long = &H800000
Private Const WS_EX_CLIENTEDGE As Long = &H200
Private Const TVS_HASLINES As Long = &H2
Private Const TVS_LINESATROOT As Long = &H4

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal uCmd As Long) As LongPtr

Private Function TF(ByVal ok As Boolean) As String
    If ok Then TF = "Y" Else TF = "N"
End Function

Private Function StyleOf(ByVal h As LongPtr) As Long
    StyleOf = GetWindowLong(h, GWL_STYLE)
End Function

Private Function ExStyleOf(ByVal h As LongPtr) As Long
    ExStyleOf = GetWindowLong(h, GWL_EXSTYLE)
End Function

Private Sub Form_Load()
    Dim o As Object
    Dim p As Object
    Dim hA As LongPtr, hB As LongPtr
    Dim st As Long, ex As Long

    hA = tvA.hwnd
    hB = tvB.hwnd

    ' --- HM1..HM3 创建期的默认观感 (原生窗口自己就是判据) ---
    Debug.Print "HM1=" & TF((StyleOf(hA) And WS_BORDER) <> 0)
    Debug.Print "HM2=" & TF((StyleOf(hA) And TVS_HASLINES) <> 0)
    ' 本 RTL 的 TreeView 创建路不挂 WS_EX_CLIENTEDGE (Appearance 默认读 0);
    ' 钉的是"默认读数与窗口一致", 不是"VB6 的 Appearance 默认值"。
    Debug.Print "HM3=" & TF((ExStyleOf(hA) And WS_EX_CLIENTEDGE) = 0)

    Set o = Me.Controls("tvA")

    ' --- HM4..HM6 ShowLines: 关 -> 窗口位清; 开 -> 窗口位立; 同值重写仍读回 True ---
    o.ShowLines = False
    st = StyleOf(hA)
    Debug.Print "HM4=" & TF((st And TVS_HASLINES) = 0 And o.ShowLines = 0)
    o.ShowLines = True
    st = StyleOf(hA)
    Debug.Print "HM5=" & TF((st And TVS_HASLINES) <> 0 And o.ShowLines <> 0)
    ' 同值重写: 以前"样式位没变化"就落到 reportUnknown (C3_HOST_MEMBER_MISS 误报)。
    ' 这一条把"值读得回 + 位还在"钉住; 误报那条由 harness 的 host-miss 计数钉。
    o.ShowLines = True
    st = StyleOf(hA)
    Debug.Print "HM6=" & TF((st And TVS_HASLINES) <> 0 And o.ShowLines <> 0)

    ' --- HM7..HM8 Appearance: 1 -> WS_EX_CLIENTEDGE 立, 读回 1; 0 -> 清, 读回 0 ---
    o.Appearance = 1
    ex = ExStyleOf(hA)
    Debug.Print "HM7=" & TF((ex And WS_EX_CLIENTEDGE) <> 0 And o.Appearance = 1)
    o.Appearance = 0
    ex = ExStyleOf(hA)
    Debug.Print "HM8=" & TF((ex And WS_EX_CLIENTEDGE) = 0 And o.Appearance = 0)

    ' --- HM9..HM10 ShowBorder: False -> WS_BORDER 清, 读回 False; True -> 立, 读回 True ---
    o.ShowBorder = False
    st = StyleOf(hA)
    Debug.Print "HM9=" & TF((st And WS_BORDER) = 0 And o.ShowBorder = False)
    o.ShowBorder = True
    st = StyleOf(hA)
    Debug.Print "HM10=" & TF((st And WS_BORDER) <> 0 And o.ShowBorder <> 0)

    ' --- HM11..HM12 Tag: 晚绑定写、晚绑定读、直接读 —— 两条路必须同一份存储 ---
    o.Tag = "M3-TAG"
    Debug.Print "HM11=" & TF(CStr(o.Tag) = "M3-TAG")
    Debug.Print "HM12=" & TF(CStr(tvA.Tag) = "M3-TAG")
    ' 隔离: tvB 的 Tag 仍为空 (没被 tvA 那次写串过去)
    Debug.Print "HM13=" & TF(CStr(tvB.Tag) = "")

    ' --- HM14..HM15 ZOrder: 0 -> 最上 (上面没有兄弟); 1 -> 最下 (下面没有兄弟) ---
    o.ZOrder 0
    Debug.Print "HM14=" & TF(GetWindow(hA, GW_HWNDPREV) = 0)
    o.ZOrder 1
    Debug.Print "HM15=" & TF(GetWindow(hA, GW_HWNDNEXT) = 0)

    ' --- HM16..HM17 RightMargin: 缇值往返 + 归零 (读回问 EM_GETMARGINS on the control) ---
    Dim r As Object
    Set r = Me.Controls("rt1")
    r.RightMargin = 720
    Debug.Print "HM16=" & TF(r.RightMargin = 720)
    r.RightMargin = 0
    Debug.Print "HM17=" & TF(r.RightMargin = 0)

    ' --- HM18 隔离: 对 tvA 的样式写入一次都没落到 tvB ---
    Debug.Print "HM18=" & TF((StyleOf(hB) And WS_BORDER) <> 0 _
                            And (StyleOf(hB) And TVS_HASLINES) <> 0 _
                            And (ExStyleOf(hB) And WS_EX_CLIENTEDGE) = 0)

    ' --- HM19 Caption: 晚绑定写到窗体, 直接读回 (两边问的是同一枚窗口的 WM_SETTEXT) ---
    Set p = Me
    p.Caption = "M3-CAP"
    Debug.Print "HM19=" & TF(Me.Caption = "M3-CAP")

    ' --- HM20..HM21 LineStyle: 1 -> LINESATROOT 立 (HASLINES 同时必须在);
    '     0 -> 只清 LINESATROOT, HASLINES 留住 (写侧旧实现把这一位判反了) ---
    o.LineStyle = 1
    st = StyleOf(hA)
    Debug.Print "HM20=" & TF((st And TVS_LINESATROOT) <> 0 And (st And TVS_HASLINES) <> 0)
    o.LineStyle = 0
    st = StyleOf(hA)
    Debug.Print "HM21=" & TF((st And TVS_LINESATROOT) = 0 And (st And TVS_HASLINES) <> 0)

    Debug.Print "HM-LOAD-DONE"
End Sub

Private Sub evtTimer_Timer()
    Static done As Integer
    If done Then Exit Sub
    done = 1
    Debug.Print "HM-DONE"
    Unload Me
End Sub
