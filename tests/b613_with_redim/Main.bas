Attribute VB_Name = "Main"
Option Explicit

' 账 #613 (第351刀): With 块里 ReDim .Member 的元素档 —— 修前回落 variant:
'   ① 元素档错 ⇒ 分配 24 字节一步、访问按声明档走 (UDT 元素 ⇒ 越界写)
'   ② 整数组读进 Variant 的 VarType 答 8204 (vbArray|vbVariant) 而不是 8195 (vbArray|vbLong)
' 两枚对照组 (模块级数组 / 第350刀那一档的类外 ReDim) 钉「没改过头」。

Private Type THost
    Items() As Long
End Type

Dim g() As Long

Sub Main()
    Dim o As K613
    Dim h As THost
    Dim v As Variant

    Set o = New K613

    ' 对照组 A: 模块级动态数组 (Long)
    ReDim g(1 To 2)
    g(2) = 3
    Debug.Print "613-ctl-mod=[" & VarType(g) & "]"

    ' ① With 块 + 工程类实例 owner: 元素档 (Long)
    With o
        ReDim .L(1 To 3)
        .L(2) = 7
    End With
    Debug.Print "613-with-cls=[" & o.L(2) & "]"
    v = o.L
    Debug.Print "613-with-cls-vt=[" & VarType(v) & "]"

    ' ② With 块 + UDT owner: 元素档 (Long)
    With h
        ReDim .Items(1 To 2)
        .Items(2) = 5
    End With
    Debug.Print "613-with-udt=[" & h.Items(2) & "]"

    ' ③ With 块 + 元素本身是 UDT: 步长必须等于 sizeof(TBig)=64 (修前 24 ⇒ 越界写)
    With o
        ReDim .Big(1 To 2)
        .Big(1).A = 111
        .Big(1).B = 222
        .Big(1).C = 333
        .Big(1).D = 444
        .Big(1).E = 555
        .Big(1).F = 666
        .Big(1).G = 777
        .Big(1).H = 888
        Debug.Print "613-udtelem=[" & .Big(1).A & "," & .Big(1).B & "," & .Big(1).C & "," _
                        & .Big(1).D & "," & .Big(1).E & "," & .Big(1).F & "," & .Big(1).G & "," _
                        & .Big(1).H & "]"
    End With

    ' 对照组 B: 第350刀那一档 (类外 ReDim, 不带 With) —— 不许退
    ReDim o.L(1 To 4)
    o.L(4) = 9
    v = o.L
    Debug.Print "613-plain-cls-vt=[" & VarType(v) & "]"

    Debug.Print "613-DONE"
End Sub
