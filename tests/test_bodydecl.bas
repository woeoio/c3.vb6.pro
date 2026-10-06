Option Explicit

' <vbeclipse> 账 #215 = 体级声明收成一条声明符一条语句。
' 改前有四条体级声明路, 两种形状: `Dim a, b` 由 parseDimStmt 自己手写一遍展开(那份副本
' 漏了 parseVariableDecl 里的 WithEvents 与「后缀即类型」两步), 而 `Const A = 1, B = 2` /
' `Static p As Long, q As Long` / 体级 `Public x As Long, y As Long` 把 MultiDecl 原样
' 交给语义层, 那儿的 switch 不认这个 kind ⇒ 一枚名字都不登记, 每条使用报一条 VB3001
' (VBFlexGridDemo 一片就 276 条)。发码那侧一直是对的, 所以值没错 —— 但 `Dim a&, b&`
' 第二枚落回 Variant 是**真值差**(TypeName 与装箱档都跟着变)。
' 两头都要钉: 数值/类型那一面(BD-*), 和"诊断里不该再有 VB3001"那一面(登记在
' run_tests.ps1 的 [CODEGEN-NOTE] bodydecl_no_undeclared 里, -Absent VB3001)。

Public Sub Main()
    Dim dA As Long, dB As Long
    dA = 4
    dB = 5
    Debug.Print "BD-dim=" & (dA + dB)

    Dim sA&, sB&
    sA = 7
    sB = 8
    Debug.Print "BD-suffix=" & TypeName(sA) & "/" & TypeName(sB) & "/" & (sA + sB)

    ' 后缀档的第二枚以前落回 Variant, 而 TypeName 看不出来 (装箱后照样打 Long) ——
    ' 没赋值时 VarType 才分得开: 该是 3/3, 改前是 3/0 (第二枚是 Empty 的 Variant)。
    Dim u1&, u2&
    Debug.Print "BD-empty=" & VarType(u1) & "/" & VarType(u2)

    Const c1 = 3, c2 = 4, c3 = 5
    Debug.Print "BD-const=" & (c1 * 100 + c2 * 10 + c3)

    Static st1 As Long, st2 As Long
    st1 = 11
    st2 = 22
    Debug.Print "BD-static=" & (st1 + st2)

    Dim arr1(2) As Long, arr2(3) As Long
    arr1(1) = 6
    arr2(2) = 7
    Debug.Print "BD-arr=" & (arr1(1) + arr2(2)) & "/" & UBound(arr1) & "/" & UBound(arr2)

    Dim vA, vB
    vA = 1
    vB = "z"
    Debug.Print "BD-variant=" & TypeName(vA) & "/" & TypeName(vB) & "/" & Len(vB)

    Debug.Print "BD-DONE"
End Sub
