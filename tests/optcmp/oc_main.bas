Option Compare Text
Option Explicit

' <vbeclipse>: `Option Compare Text` 是按**模块**的编译期属性。修复前 module.options
' 从没进过代码生成层 (那条 g_vb6_optionCompareText 赋值从来发不出来), 所以本模块里
' `a = b` / InStr / StrComp / Replace / Split / Filter / Like / Select Case 全部仍按
' 二进制比 —— 与 VB6 不一致, 而且同一个模块里 `=` 与显式 vbTextCompare 互相矛盾。
' 读数一律 ASCII 载荷; Binary 侧的对照在同工程的 oc_binary.bas (证明不串模块)。
Public Sub Main()
    Dim a As String
    Dim b As String
    a = "abc"
    b = "ABC"
    Debug.Print "OC-eq="; (a = b)
    Debug.Print "OC-lt="; (a < b)
    Debug.Print "OC-instr="; InStr(a, "ABC")
    Debug.Print "OC-instr-exp-b="; InStr(a, "ABC", vbBinaryCompare)
    Debug.Print "OC-instr3="; InStr(1, a, "ABC")
    Debug.Print "OC-sc="; StrComp(a, b)
    Debug.Print "OC-rep=["; Replace("aAaA", "a", "-"); "]"
    Dim sp As Variant
    sp = Split("aXbXc", "x")
    Debug.Print "OC-split-ub="; UBound(sp)
    Dim ft As Variant
    ft = Filter(Array("abc", "ABD"), "ab")
    Debug.Print "OC-filter-ub="; UBound(ft)
    Debug.Print "OC-like="; (a Like "[A-Z]*")
    Select Case a
        Case "ABC"
            Debug.Print "OC-case=hit"
        Case Else
            Debug.Print "OC-case=miss"
    End Select
    Debug.Print "OC-binary-mod="; OCBinaryProbe()
    Debug.Print "OC-DONE"
End Sub
