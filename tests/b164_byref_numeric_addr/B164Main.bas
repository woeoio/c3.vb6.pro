Option Explicit

Private mn(1 To 3) As Long
Private gn As Long

Public Sub Main()
    Dim box As B164Box
    Dim o As Object
    Dim ln(1 To 2) As Long
    Dim in1(1 To 2) As Integer
    Dim dn(1 To 2) As Double
    Dim bl(1 To 2) As Boolean
    Dim by(1 To 2) As Byte

    Set box = New B164Box
    Set o = box

    ' N1: 局部静态 Long 数组元素
    ln(1) = 10
    o.Inc ln(1)
    Debug.Print "B164-N1=" & ln(1)

    ' N2: 模块级 Long 数组元素
    mn(1) = 20
    o.Inc mn(1)
    Debug.Print "B164-N2=" & mn(1)

    ' N3: 类实例 Public Long 字段
    box.N = 30
    o.Inc box.N
    Debug.Print "B164-N3=" & box.N

    ' N4: 模块级裸 Long（基线）
    gn = 40
    o.Inc gn
    Debug.Print "B164-N4=" & gn

    ' N5: Integer 数组元素（2 字节；按 I4 写回会越界写）
    in1(1) = 50
    o.IncI in1(1)
    Debug.Print "B164-N5=" & in1(1)

    ' N6: 类实例 Public Integer 字段
    box.I = 60
    o.IncI box.I
    Debug.Print "B164-N6=" & box.I

    ' N7: Double 数组元素
    dn(1) = 70
    o.IncD dn(1)
    Debug.Print "B164-N7=" & dn(1)

    ' N8: 类实例 Public Double 字段
    box.D = 80
    o.IncD box.D
    Debug.Print "B164-N8=" & box.D

    ' N9: Single 类字段（按 R4 写回，不能当 R8）
    box.S = 90
    o.IncS box.S
    Debug.Print "B164-N9=" & box.S

    ' N10: Boolean 数组元素（int16_t 与 Integer 同形，档靠声明表分）
    ' 用 If 判值而不是拼串: Boolean 数组元素走 `&` 折出来是 -1（读路径在册差异，
    ' 与本刀的回写无关），这里只钉「体里那一发写回来了」。
    bl(1) = False
    o.FlipB bl(1)
    If bl(1) Then
        Debug.Print "B164-N10=True"
    Else
        Debug.Print "B164-N10=False"
    End If

    ' N11: Boolean 类字段
    box.B = False
    o.FlipB box.B
    Debug.Print "B164-N11=" & box.B

    ' N12: Byte 数组元素 —— §11.146 (第一百六十四刀) 起也回写 (以前无 UI1 那一档，按值)
    by(1) = 1
    o.SetY by(1)
    Debug.Print "B164-N12=" & by(1)

    ' N13: 负对照 —— 属性读不可取址，仍按值（在册）
    box.Tag = 5
    o.Inc box.Tag
    Debug.Print "B164-N13=" & box.Tag

    ' N14: 同一元素连发两次 —— 回写后槽里必须还是能用的值
    ln(2) = 0
    o.Inc ln(2)
    o.Inc ln(2)
    Debug.Print "B164-N14=" & ln(2)

    Debug.Print "B164-DONE"
    Set box = Nothing
    Set o = Nothing
End Sub
