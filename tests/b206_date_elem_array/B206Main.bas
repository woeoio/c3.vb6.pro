Option Explicit

' §11.181 第一百九十八刀: Date 元素档。
' 真机读数 `.temp/k198p/p234.txt`（VB6Mini /make）：未分配 8199/"Date()"；ReDim 之后
' 8199/"Date()"；元素 VarType 7 / TypeName "Date"；整数组装箱 8199；Erase 之后源仍 8199；
' For Each 元素 7/"Date"；ReDim Preserve 之后标签与界都在。
' 刻意不钉两格（本刀明示不开口，见 §11.181 第五节）：
'   · `Erase q` 之后 `VarType(q)`/`TypeName(q)` 走的是「谓词实参」那一站，仍答 8204/"Variant()"
'     （真机 8199/"Date()"；这一格对**所有**元素档都一样，是 §11.180 留的那条背账的新读数，
'     见 .temp/k198p/p236.txt）；
'   · `q(0) = "2021-03-04"`（String→Date 隐式转换）本仓编译期就报 VB3003，真机 44259
'     （见 .temp/k198p/p235.txt）。

Public Sub Main()
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\b206.txt" For Output As #f

    Dim v As Variant
    Dim q() As Date
    Dim d As Date, d2 As Date

    ' ---- D1: 未分配 ----
    v = q
    Print #f, "D1 unalloc vt=" & VarType(v) & " isarr=" & IsArray(v) & " tn=" & TypeName(v)

    ' ---- D2: ReDim 之后 ----
    ReDim q(0 To 2)
    Print #f, "D2 alloc vt=" & VarType(q) & " tn=" & TypeName(q) & " lb=" & LBound(q) & " ub=" & UBound(q)

    ' ---- D3: 日期字面量写进元素 ----
    q(0) = #1/2/2020 3:04:05 PM#
    Print #f, "D3 elem vt=" & VarType(q(0)) & " tn=" & TypeName(q(0)) & " dbl=" & CStr(CDbl(q(0)))

    ' ---- D4: 纯数字（日期序列号）写进元素 ----
    q(1) = 45000.5
    Print #f, "D4 num vt=" & VarType(q(1)) & " tn=" & TypeName(q(1)) & " dbl=" & CStr(CDbl(q(1)))

    ' ---- D5: 元素装进 Variant ----
    v = q(1)
    Print #f, "D5 box vt=" & VarType(v) & " tn=" & TypeName(v) & " dbl=" & CStr(CDbl(v))

    ' ---- D6: 整个已分配数组装箱 ----
    v = q
    Print #f, "D6 arr vt=" & VarType(v) & " tn=" & TypeName(v) & " isarr=" & IsArray(v) & " lb=" & LBound(v) & " ub=" & UBound(v)

    ' ---- D7: 装箱之后再 Erase 源，槽里那份不动（副本语义 §11.179）----
    Erase q
    Print #f, "D7 aftererase vt=" & VarType(v) & " tn=" & TypeName(v) & " ub=" & UBound(v)

    ' ---- D8: 再 ReDim，写日期字面量 ----
    ReDim q(0 To 1)
    q(0) = #3/4/2021#
    Print #f, "D8 literal vt=" & VarType(q(0)) & " tn=" & TypeName(q(0)) & " dbl=" & CStr(CDbl(q(0)))

    ' ---- D9: For Each 取出来的档 ----
    For Each v In q
        Print #f, "D9 each vt=" & VarType(v) & " tn=" & TypeName(v)
    Next

    ' ---- D10: 纯量 Date 装箱做对照 ----
    d = #12/31/1999 11:59:59 PM#
    v = d
    Print #f, "D10 scalar vt=" & VarType(v) & " tn=" & TypeName(v) & " dbl=" & CStr(CDbl(v))

    ' ---- D11: Variant 里的日期写回元素 ----
    v = #2/3/2004 5:06:07 AM#
    q(1) = v
    Print #f, "D11 back vt=" & VarType(q(1)) & " tn=" & TypeName(q(1)) & " dbl=" & CStr(CDbl(q(1)))

    ' ---- D12: ReDim Preserve 之后标签还在不在 ----
    ReDim Preserve q(0 To 3)
    q(3) = #6/7/2008#
    Print #f, "D12 pres vt=" & VarType(q) & " tn=" & TypeName(q) & " ub=" & UBound(q) & " dbl=" & CStr(CDbl(q(3)))

    ' ---- D13: 元素赋给 Date 变量 ----
    d2 = q(3)
    Print #f, "D13 tovar dbl=" & CStr(CDbl(d2))

    Close #f
End Sub
