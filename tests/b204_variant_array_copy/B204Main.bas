Option Explicit

' §11.179 第一百九十六刀: 数组装进 Variant 是**副本**, 不是共用一枚描述符。
' 真机读数见 .temp/k196p/p230.txt (VB6Mini /make 跑出来的):
'   T1 写槽 → 源不动 99/11      T2 写源 → 槽不动 55/11
'   T3 Variant→Variant 两边独立 11/77
'   T4 Variant() 数组 88/7      T5 String() 数组 bb/aa
'   T6 源被 ReDim 后槽留着旧的界与值 9/2/11
'   T7 槽换成另一档数组后标签跟着换 8200/"String()"
'   T8 界/身份照旧 8195/True/0/2
' 背账 (不在本刀): 未分配数组装箱后 VarType —— 真机答 8195 (保住元素档), 本仓答 8204。

Public Sub Main()
    On Error Resume Next

    Dim r As Variant, r2 As Variant
    Dim la(0 To 2) As Long
    Dim va(0 To 2) As Variant
    Dim sa(0 To 2) As String

    ' ---- T1: 写槽, 看源 ----
    la(0) = 11: la(1) = 12: la(2) = 13
    r = la
    r(0) = 99
    Debug.Print "T1 slot=" & r(0) & " src=" & la(0)

    ' ---- T2: 写源, 看槽 ----
    la(0) = 11
    r = la
    la(0) = 55
    Debug.Print "T2 src=" & la(0) & " slot=" & r(0)

    ' ---- T3: Variant → Variant 各持一份 ----
    la(0) = 11: la(1) = 12
    r = la
    r2 = r
    r2(0) = 77
    Debug.Print "T3 r1=" & r(0) & " r2=" & r2(0)

    ' ---- T4: Variant() 数组 (逐元素另拷) ----
    va(0) = 7: va(1) = 8
    r = va
    r(0) = 88
    Debug.Print "T4 slot=" & r(0) & " src=" & va(0)

    ' ---- T5: String() 数组 (BSTR 另分配一只) ----
    sa(0) = "aa"
    r = sa
    r(0) = "bb"
    Debug.Print "T5 slot=" & r(0) & " src=" & sa(0)

    ' ---- T6: 源被 ReDim 之后槽仍是旧的那一份 ----
    Dim dl() As Long
    ReDim dl(0 To 2)
    dl(0) = 11: dl(1) = 12: dl(2) = 13
    r = dl
    ReDim dl(0 To 9)
    dl(0) = 44
    Debug.Print "T6 src_ub=" & UBound(dl) & " slot_ub=" & UBound(r) & " slot0=" & r(0)

    ' ---- T7: 槽改成另一档数组, 标签跟着换 ----
    la(0) = 11
    r = la
    r = sa
    Debug.Print "T7 vt=" & VarType(r) & " tn=" & TypeName(r)

    ' ---- T8: 身份与界 ----
    la(0) = 11: la(1) = 12
    r = la
    Debug.Print "T8 isarr=" & IsArray(r) & " vt=" & VarType(r) & " lb=" & LBound(r) & " ub=" & UBound(r)
End Sub
