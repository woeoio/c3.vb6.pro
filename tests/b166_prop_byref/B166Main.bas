Option Explicit

' 账 #305 §11.147 (第一百六十五刀): 属性读当 ByRef 实参 —— 读-改-写。
' 每一格都是「晚绑定目标 + 属性读实参」，答案 = 调用前后的属性值。
' P8 是**反面证人**: String 属性暂不开 (临时量要管 BSTR 所有权)，仍按值 ⇒ "ab"。
Public Sub Main()
    Dim b As B166Box
    Dim s As String
    Dim o As Object

    Set b = New B166Box
    Set o = b

    b.Tag = 40
    o.Inc b.Tag
    Debug.Print "B166-P1=" & b.Tag
    o.Inc b.Tag
    Debug.Print "B166-P2=" & b.Tag

    b.INum = 20
    o.IncI b.INum
    Debug.Print "B166-P3=" & b.INum

    b.DNum = 30
    o.IncD b.DNum
    Debug.Print "B166-P4=" & b.DNum

    b.YNum = 3
    o.SetY b.YNum
    Debug.Print "B166-P5=" & b.YNum

    b.FNum = False
    o.FlipB b.FNum
    Debug.Print "B166-P6=" & b.FNum

    b.Tm = DateSerial(2020, 1, 1)
    o.AddDay b.Tm, 2
    Debug.Print "B166-P7=" & Format(b.Tm, "yyyy-mm-dd")

    b.Label2 = "ab"
    o.Append b.Label2, "cd"
    Debug.Print "B166-P8=" & b.Label2

    s = "xy"
    o.Append s, "z"
    Debug.Print "B166-P9=" & s

    Debug.Print "B166-DONE"
End Sub
