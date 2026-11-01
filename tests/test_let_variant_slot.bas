Attribute VB_Name = "test_let_variant_slot"
Option Explicit

' 账 #305 §B164 形⑤ 站③ (第一百八十四刀): 显式 `Let <变体> = …` 的七形。
' 改前: 非变体右值 (字面量 / String 变量 / 拼接) 三形发成裸赋值 = cl C2440 (BASE 实测 3 发);
'       变体槽与变体数组元素两形编得过但不清旧值, 且借来的值与槽共用同一只 BSTR。
' 改后: 目标是不是变体槽只问 variantSlotTargetOf, 命中就交给 emitVariantSlotAssign
'       (装箱 + 所有权 + 清旧值)。两个证人: s3/s6 = 源头改写后槽里必须还是旧值 (深拷贝)。
Public gv As Variant
Public ga(1 To 2) As Variant
Public sv As String

Sub Main()
    Dim v As Variant
    v = "src"
    sv = "seed"

    Let gv = "abc"
    Debug.Print "L171-s1=" & gv & "/" & VarType(gv)
    Let gv = sv
    Debug.Print "L171-s2=" & gv & "/" & VarType(gv)
    sv = "moved"
    Debug.Print "L171-s3=" & gv
    Let gv = sv & "!"
    Debug.Print "L171-s4=" & gv
    Let gv = v
    Debug.Print "L171-s5=" & gv
    v = "changed"
    Debug.Print "L171-s6=" & gv
    Let ga(1) = "arr"
    Debug.Print "L171-s7=" & ga(1)
    Let ga(1) = v
    Debug.Print "L171-s8=" & ga(1)
    Debug.Print "L171-DONE"
End Sub
