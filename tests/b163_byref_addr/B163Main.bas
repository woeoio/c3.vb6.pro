Option Explicit

Private ma(1 To 3) As String
Private mg As String

Public Sub Main()
    Dim box As B163Box
    Dim o As Object
    Dim la(1 To 2) As String
    Dim dyn() As String

    Set box = New B163Box
    Set o = box

    ' A1: 局部静态数组元素（发码 VB6_SA_AT(BSTR, la, i) —— 宏首参就是元素档）
    la(1) = "la"
    o.Bump la(1)
    Debug.Print "B163-A1=" & la(1)

    ' A2: 模块级数组元素
    ma(1) = "ma"
    o.Bump ma(1)
    Debug.Print "B163-A2=" & ma(1)

    ' A3: 动态数组元素
    ReDim dyn(1 To 2)
    dyn(2) = "dy"
    o.Swap dyn(2)
    Debug.Print "B163-A3=" & dyn(2)

    ' A4: 类型化类实例的 Public String 字段（发码 base->F 带 class var 注释）
    box.S = "fs"
    o.Bump box.S
    Debug.Print "B163-A4=" & box.S

    ' A5: With 块里的成员（发码 _vb6_with_N->F 带 With class 注释）
    With box
        .S = "ws"
        o.Bump .S
        Debug.Print "B163-A5=" & .S
    End With

    ' A6: 属性读 —— 不可取址，仍按值发（在册：VB6 会经 Let 写回）
    box.Tag = "pt"
    o.Bump box.Tag
    Debug.Print "B163-A6=" & box.Tag

    ' A7: 模块级裸 String 变量（基线，第一百六十一刀）
    mg = "mg"
    o.Bump mg
    Debug.Print "B163-A7=" & mg

    ' A8: 数组元素连发两次 —— 回写后槽里必须还是能用的串
    la(2) = "tw"
    o.Bump la(2)
    o.Bump la(2)
    Debug.Print "B163-A8=" & la(2)

    Debug.Print "B163-DONE"
    Set box = Nothing
    Set o = Nothing
End Sub
