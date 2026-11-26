Option Explicit

' §11.182 第一百九十九刀: **谓词实参站**的数组标签。
' 真机读数 `.temp/k199p/p238.txt`（VB6Mini /make，局部/模块级 × Long/Date/String/Byte/
' Variant/一维）与 `.temp/k199p/p238b.txt`（多维）：未分配与 Erase 之后的动态数组，
' VarType 一律答 8192|元素码、TypeName 答 "<元素>()"、IsArray=True，**Err 恒 0**；
' UBound/LBound 才是 9（那是 §11.180 那一支，本夹具只在 1D 上钉它）。
' 用 Debug.Print 出到 stdout（夹具 harness 比的就是 stdout）。
' 刻意不钉两格（不在本刀站内，记在 §11.182 第六节背账）：
'   · Erase 之后**多维**的 UBound/LBound：C3 答 0/err=0，真机抛 9
'     （.temp/k199p/p238b.txt 的 N5/N6；那一问不经装箱站，是 vb6_UBoundND 的缺口）；
'   · String→Date 隐式转换（.temp/k198p/p235.txt，编译期 VB3003）。

Private mq() As Long
Private md() As Date
Private ms() As String
Private m2() As Long
Private mv() As Variant

Public Sub Main()
    Dim lq() As Long
    Dim ld() As Date
    Dim ls() As String
    Dim lb() As Byte
    Dim lv() As Variant
    Dim e As Long

    On Error Resume Next

    ' ---- L1..L5: 局部动态数组，从未 ReDim ----
    Debug.Print "L1 vt=" & VarType(lq) & " tn=" & TypeName(lq) & " isa=" & IsArray(lq) & " err=" & Err.Number
    Err.Clear
    Debug.Print "L2 vt=" & VarType(ld) & " tn=" & TypeName(ld) & " isa=" & IsArray(ld) & " err=" & Err.Number
    Err.Clear
    Debug.Print "L3 vt=" & VarType(ls) & " tn=" & TypeName(ls) & " isa=" & IsArray(ls) & " err=" & Err.Number
    Err.Clear
    Debug.Print "L4 vt=" & VarType(lb) & " tn=" & TypeName(lb) & " isa=" & IsArray(lb) & " err=" & Err.Number
    Err.Clear
    Debug.Print "L5 vt=" & VarType(lv) & " tn=" & TypeName(lv) & " isa=" & IsArray(lv) & " err=" & Err.Number
    Err.Clear

    ' ---- M1..M5: 模块级动态数组，从未 ReDim（M2 = Date，改前发码编不过 C2440）----
    Debug.Print "M1 vt=" & VarType(mq) & " tn=" & TypeName(mq) & " isa=" & IsArray(mq) & " err=" & Err.Number
    Err.Clear
    Debug.Print "M2 vt=" & VarType(md) & " tn=" & TypeName(md) & " isa=" & IsArray(md) & " err=" & Err.Number
    Err.Clear
    Debug.Print "M3 vt=" & VarType(ms) & " tn=" & TypeName(ms) & " isa=" & IsArray(ms) & " err=" & Err.Number
    Err.Clear
    Debug.Print "M4 vt=" & VarType(m2) & " tn=" & TypeName(m2) & " isa=" & IsArray(m2) & " err=" & Err.Number
    Err.Clear
    Debug.Print "M5 vt=" & VarType(mv) & " tn=" & TypeName(mv) & " isa=" & IsArray(mv) & " err=" & Err.Number
    Err.Clear

    ' ---- E1/E2: Erase 之后仍按声明档答标签；UBound 抛 9；装箱进 Variant 同档 ----
    ReDim mq(0 To 2)
    Erase mq
    Err.Clear
    Debug.Print "E1 vt=" & VarType(mq) & " tn=" & TypeName(mq) & " isa=" & IsArray(mq) & " err=" & Err.Number
    Err.Clear: e = UBound(mq)
    Debug.Print "E1 ub err=" & Err.Number
    Err.Clear
    Dim v As Variant
    v = mq
    Debug.Print "E2 vt=" & VarType(v) & " tn=" & TypeName(v) & " isa=" & IsArray(v) & " err=" & Err.Number
    Err.Clear

    ' ---- A1: 已分配那一档（对照，标签问描述符自己）----
    ReDim mq(0 To 1)
    Debug.Print "A1 vt=" & VarType(mq) & " tn=" & TypeName(mq) & " isa=" & IsArray(mq) & " err=" & Err.Number
    Err.Clear
    Debug.Print "A1 ub=" & UBound(mq) & " err=" & Err.Number
    Err.Clear

    ' ---- N1..N5: 多维动态数组（未分配 / 已分配 / Erase 之后）----
    Dim a2() As Long
    Dim b2() As Date
    Dim c2() As String
    Debug.Print "N1 vt=" & VarType(a2) & " tn=" & TypeName(a2) & " isa=" & IsArray(a2) & " err=" & Err.Number
    Err.Clear
    Debug.Print "N2 vt=" & VarType(b2) & " tn=" & TypeName(b2) & " isa=" & IsArray(b2) & " err=" & Err.Number
    Err.Clear
    Debug.Print "N3 vt=" & VarType(c2) & " tn=" & TypeName(c2) & " isa=" & IsArray(c2) & " err=" & Err.Number
    Err.Clear
    ReDim a2(1 To 2, 1 To 3)
    Debug.Print "N4 vt=" & VarType(a2) & " tn=" & TypeName(a2) & " isa=" & IsArray(a2) & " err=" & Err.Number
    Err.Clear
    Erase a2
    Debug.Print "N5 vt=" & VarType(a2) & " tn=" & TypeName(a2) & " isa=" & IsArray(a2) & " err=" & Err.Number
    Err.Clear
End Sub
