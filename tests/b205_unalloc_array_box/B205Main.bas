Option Explicit

' §11.180 第一百九十七刀: **未分配**（Dim 还没 ReDim / Erase 之后）的数组装进 Variant。
' 真机读数见 .temp/k197p/p233.txt（VB6Mini /make 跑出来的）：
'   Long 8195 / String 8200 / Double 8197 / Currency 8198 / Boolean 8203 / Byte 8209 /
'   Variant 8204 / Object 8201 —— 标签与 TypeName 全按**声明**那一档，UBound/LBound 抛 9。
'   Erase 之后同形；装箱后再 ReDim 源，槽的标签与界都不变。
' 刻意不钉的两格（本刀明示不开口，见 §11.180 第五节）：
'   · Date()：真机 8199/"Date()"，本仓枚举里没有 Date 档（§B159），答 8197/"Double()"；
'   · Object() 的 UBound：真机答 -1 且不报错，本仓抛 9（只在这一格上不同）。

Public Sub Main()
    On Error Resume Next

    Dim r As Variant
    Dim vt As Long, isa As Boolean, tn As String
    Dim e1 As Long, e2 As Long

    Dim dl() As Long
    Err.Clear: r = dl
    vt = VarType(r): isa = IsArray(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Err.Clear: e2 = LBound(r): e2 = Err.Number
    Debug.Print "T1 long vt=" & vt & " isarr=" & isa & " tn=" & tn & " ub=" & e1 & " lb=" & e2

    Dim ds() As String
    Err.Clear: r = ds
    vt = VarType(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T2 str vt=" & vt & " tn=" & tn & " ub=" & e1

    Dim dd() As Double
    Err.Clear: r = dd
    vt = VarType(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T3 dbl vt=" & vt & " tn=" & tn & " ub=" & e1

    Dim dc() As Currency
    Err.Clear: r = dc
    vt = VarType(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T4 cur vt=" & vt & " tn=" & tn & " ub=" & e1

    Dim db() As Boolean
    Err.Clear: r = db
    vt = VarType(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T5 bool vt=" & vt & " tn=" & tn & " ub=" & e1

    Dim dby() As Byte
    Err.Clear: r = dby
    vt = VarType(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T6 byte vt=" & vt & " tn=" & tn & " ub=" & e1

    Dim dv() As Variant
    Err.Clear: r = dv
    vt = VarType(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T7 variant vt=" & vt & " tn=" & tn & " ub=" & e1

    Dim dobj() As Object
    Err.Clear: r = dobj
    vt = VarType(r): tn = TypeName(r)
    Debug.Print "T8 object vt=" & vt & " tn=" & tn

    ' ---- T9: ReDim 之后再 Erase ----
    ReDim dl(0 To 2)
    dl(0) = 7
    Erase dl
    Err.Clear: r = dl
    vt = VarType(r): tn = TypeName(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T9 erased vt=" & vt & " tn=" & tn & " ub=" & e1

    ' ---- T10: 装箱之后再 ReDim 源，槽的标签与界都不变（副本语义，§11.179）----
    Dim dl2() As Long
    r = dl2
    ReDim dl2(0 To 2)
    vt = VarType(r): isa = IsArray(r)
    Err.Clear: e1 = UBound(r): e1 = Err.Number
    Debug.Print "T10 srcredim vt=" & vt & " isarr=" & isa & " ub=" & e1

    ' ---- T11: 槽被重新赋标量 ----
    Dim dl3() As String
    r = dl3
    r = 5
    Debug.Print "T11 overwrite vt=" & VarType(r) & " val=" & CStr(r)
End Sub
