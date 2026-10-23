Attribute VB_Name = "B190Main"
Option Explicit

Private Type Vert
    tU As Single
    X As Single
End Type

Sub Main()
    Dim o As New B190Holder
    Dim i As Long
    Dim bad As Long

    o.Init 4

    ' --- 写: 类成员动态数组 (声明档 Long) ---
    For i = 1 To 4
        o.Items(i) = i * 100
        o.Tags(i) = "T" & i
    Next i

    ' --- 读回来 ---
    For i = 1 To 4
        Debug.Print "R" & i & "=" & o.Items(i)
        Debug.Print "S" & i & "=" & o.GetDirect(i)
    Next i

    For i = 1 To 4
        Debug.Print "G" & i & "=" & o.Tags(i)
        Debug.Print "H" & i & "=" & o.GetTagDirect(i)
    Next i

    Debug.Print "SUM=" & o.SumAll()
    Debug.Print "UB=" & o.LastIdx()

    ' --- 装箱读 (走描述符自己的档) ---
    Dim v As Variant
    v = o.Items
    Debug.Print "VT=" & VarType(v)
    For i = 1 To 4
        Debug.Print "V" & i & "=" & v(i)
    Next i

    ' --- 控制组: 模块级动态数组 (声明档同样 Long) ---
    Dim g() As Long
    Dim j As Long
    ReDim g(1 To 4)
    For j = 1 To 4
        g(j) = j * 100
    Next j
    For j = 1 To 4
        Debug.Print "C" & j & "=" & g(j)
    Next j
    Debug.Print "CSUM=" & (g(1) + g(2) + g(3) + g(4))
    Debug.Print "CUB=" & UBound(g)

    ' --- 控制组: Variant 数组 (本来就该 16 字节一步) ---
    Dim q() As Variant
    ReDim q(1 To 2)
    q(1) = "hello"
    q(2) = 42
    Debug.Print "Q1=" & q(1)
    Debug.Print "Q2=" & q(2)

    ' --- 外部 UBound / With 成员索引 / 整数组读 ---
    Debug.Print "E1=" & UBound(o.Items)
    With o
        Debug.Print "E2=" & .Items(3)
        .Items(3) = 999
        Debug.Print "E3=" & .Items(3)
    End With
    Dim w As Variant
    w = o.Items
    Debug.Print "E4=" & VarType(w)
    ' --- LBound 各形状 ---
    Debug.Print "L1=" & LBound(o.Items)
    Debug.Print "L2=" & UBound(o.Items)
    Dim z() As Long
    ReDim z(5 To 9)
    Debug.Print "L3=" & LBound(z)
    Debug.Print "L4=" & UBound(z)
    ' --- 对照组: 模块级整数组读进 Variant (与类成员必须同一个 VT) ---
    Dim vg As Variant
    vg = g
    Debug.Print "MG=" & VarType(vg)
    Debug.Print "MC=" & VarType(w)

    Debug.Print "DONE"
End Sub
