Option Explicit

' §11.177 第一百九十四刀: 一维 Variant 下标的错误档。
' 真机读数见 .temp/k194p/p223.txt / p224.txt:
'   槽不是数组（Empty / Long / Null / String / Boolean）→ 13
'   数组没分配 / 下标越界（上、下同）→ 9
'   界内照常取值
'   读侧（v(i)）与写侧（v(i) = x）同一口径；二维（v(i,j)）同样 13 / 9。
' 本夹具只钉**错误号**与界内取值：raise 之后目标槽是否被写（真机保持原值、
' 本仓仍写 Empty）是第一百九十二刀留下的 Variant 目标缺口，另记，不在此钉。

Private Sub Cell(ByVal tag As String, ByRef r As Variant, ByVal i As Long)
    On Error Resume Next
    Dim e As Long
    Err.Clear
    Debug.Print tag & "=" & r(i)
    e = Err.Number
    Debug.Print tag & "-err=" & e
    Err.Clear
End Sub

Private Sub CellSet(ByVal tag As String, ByRef r As Variant, ByVal i As Long)
    On Error Resume Next
    Dim e As Long
    Err.Clear
    r(i) = 77
    e = Err.Number
    Debug.Print tag & "-err=" & e
    Err.Clear
End Sub

Private Sub Cell2(ByVal tag As String, ByRef r As Variant, ByVal i As Long, ByVal j As Long)
    On Error Resume Next
    Dim e As Long
    Err.Clear
    Debug.Print tag & "=" & r(i, j)
    e = Err.Number
    Debug.Print tag & "-err=" & e
    Err.Clear
End Sub

Public Sub Main()
    On Error Resume Next

    Dim r As Variant
    Dim ar(0 To 2) As Variant
    Dim la(0 To 2) As Long
    Dim da() As Variant
    Dim a2(0 To 1, 0 To 1) As Variant
    Dim u As Long
    Dim hits As Long

    ' ---- 非数组槽: 一律 13 ----
    r = Empty
    Cell "V1 empty", r, 0
    r = 5
    Cell "V2 long", r, 0
    r = Null
    Cell "V3 null", r, 0
    r = "abc"
    Cell "V4 string", r, 0
    r = True
    Cell "V5 bool", r, 0

    ' ---- 一维数组: 界内 / 越界 ----
    ar(0) = 7: ar(1) = 8: ar(2) = 9
    r = ar
    Cell "V6 inrange", r, 0
    Cell "V7 oob_hi", r, 5
    Cell "V8 oob_lo", r, -1

    ' ---- 有类型数组装进 Variant ----
    la(0) = 11: la(1) = 12
    r = la
    Cell "V9 longarr", r, 0

    ' ---- 未分配的动态数组: 9 ----
    r = da
    Cell "V10 unalloc", r, 0

    ' ---- Array() 函数据 ----
    r = Array(21, 22)
    Cell "V11 arrayfn", r, 1

    ' ---- 写侧 ----
    r = Empty
    CellSet "V12 set_empty", r, 0
    r = ar
    CellSet "V13 set_oob", r, 5

    ' ---- UBound / LBound 对照 (归 RTL 的 vb6_UBound 一侧) ----
    r = Empty
    u = 0
    Err.Clear
    u = UBound(r)
    Debug.Print "V14 ubound_empty=" & u & "-err=" & Err.Number
    Err.Clear
    u = 0
    u = LBound(r)
    Debug.Print "V15 lbound_empty=" & u & "-err=" & Err.Number
    Err.Clear

    ' ---- 二维 (ND 出口) ----
    r = Empty
    Cell2 "V16 empty2d", r, 0, 0
    a2(0, 1) = 31
    a2(1, 0) = 32
    r = a2
    Cell2 "V17 inrange2d", r, 0, 1
    Cell2 "V18 oob2d", r, 2, 0

    ' ---- 出错之后下一条语句照跑, Err 留着 (隐式 resume 不清 Err) ----
    r = Empty
    hits = 0
    Err.Clear
    Debug.Print "V19 pre=" & r(0)
    hits = hits + 1
    Debug.Print "V19 hits=" & hits & "-err=" & Err.Number
End Sub
