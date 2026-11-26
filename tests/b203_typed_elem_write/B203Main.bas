Option Explicit

' §11.178 第一百九十五刀: Variant 里**有类型**数组元素的写侧。
' 真机读数见 .temp/k195p/p226.txt / p227.txt:
'   Long()   ← 99 / "42" / 1.9(→2)         err 0
'   String() ← 77(→"77") / "new"           err 0   ← Null 是 94
'   Boolean()← 2(→True)                    err 0
'   Double() ← 3                           err 0
'   Currency()← 1.23456 (→1.2346)          err 0
'   Byte()   ← 300 ⇒ err 6、元素不变；Integer() ← 40000 ⇒ err 6
'   Long()   ← Null ⇒ 94；← "abc" ⇒ 13；两种都不写目标
' 不钉「写完源数组变不变」：真机是副本（la(0) 不变），本仓这条数组装箱仍是共用，
' 属既有缺口，另记（§11.178 第五节）。

Private Sub Show(ByVal tag As String, ByRef r As Variant, ByVal i As Long)
    On Error Resume Next
    Err.Clear
    Debug.Print tag & "=" & CStr(r(i)) & "-err=" & Err.Number
    Err.Clear
End Sub

Private Sub W(ByVal tag As String, ByRef r As Variant, ByVal i As Long, ByRef v As Variant)
    On Error Resume Next
    Err.Clear
    r(i) = v
    Debug.Print tag & "-err=" & Err.Number
    Err.Clear
End Sub

Public Sub Main()
    On Error Resume Next

    Dim r As Variant
    Dim v As Variant
    Dim la(0 To 2) As Long
    Dim va(0 To 2) As Variant
    Dim sa(0 To 2) As String
    Dim ba(0 To 2) As Boolean
    Dim da(0 To 2) As Double
    Dim bya(0 To 2) As Byte
    Dim ia(0 To 2) As Integer
    Dim ca(0 To 2) As Currency

    ' ---- 正常写入: 按元素档转换 ----
    la(0) = 11: la(1) = 12: la(2) = 0
    r = la
    v = 99
    W "T1 long_set99", r, 0, v
    Show "T1 long_after", r, 0

    va(0) = 7
    r = va
    v = 99
    W "T2 variant_set99", r, 0, v
    Show "T2 variant_after", r, 0

    sa(0) = "old"
    r = sa
    v = 77
    W "T3 str_set77", r, 0, v
    Show "T3 str_after", r, 0

    r = sa
    v = "new"
    W "T4 str_setnew", r, 0, v
    Show "T4 str_after", r, 0

    ba(0) = False
    r = ba
    v = 2
    W "T5 bool_set2", r, 0, v
    Show "T5 bool_after", r, 0

    da(0) = 1.5
    r = da
    v = 3
    W "T6 dbl_set3", r, 0, v
    Show "T6 dbl_after", r, 0

    r = la
    v = "42"
    W "T7 long_setnumstr", r, 1, v
    Show "T7 long_after", r, 1

    r = la
    v = 1.9
    W "T8 long_round", r, 2, v
    Show "T8 long_after", r, 2

    ' ---- 越界 ----
    r = la
    v = 1
    W "T9 oob", r, 5, v

    ' ---- 溢出: 报 6 且元素不动 ----
    bya(0) = 1
    r = bya
    v = 300
    W "T10 byte300", r, 0, v
    Show "T10 byte_after", r, 0

    r = bya
    v = 200
    W "T11 byte200", r, 0, v
    Show "T11 byte_after", r, 0

    ia(0) = 1
    r = ia
    v = 40000
    W "T12 int40000", r, 0, v
    Show "T12 int_after", r, 0

    ' ---- Currency: 4 位定点 ----
    ca(0) = 0
    r = ca
    v = 1.23456
    W "T13 cy", r, 0, v
    Show "T13 cy_after", r, 0

    ' ---- Null / Empty / 不可转换文本 ----
    la(0) = 5
    r = la
    v = Null
    W "T14 long_null", r, 0, v
    Show "T14 long_after", r, 0

    r = la
    v = Empty
    W "T15 long_empty", r, 0, v
    Show "T15 long_after", r, 0

    sa(0) = "z"
    r = sa
    v = Null
    W "T16 str_null", r, 0, v
    Show "T16 str_after", r, 0

    ' 先把 la(1) 写回哨兵: 本仓 `r = la` 目前是共用而不是副本（既有缺口），
    ' 同前面几格写进去的值会留在 la 里；这一格只问「13 时元素动不动」。
    la(1) = 12
    r = la
    v = "abc"
    W "T17 long_abc", r, 1, v
    Show "T17 long_after", r, 1
End Sub
