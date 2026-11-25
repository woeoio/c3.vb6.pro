Option Explicit

Public Sub Main()
    Dim box As B160Box
    Dim o As Object
    Dim n As Long
    Dim d As Double
    Dim s As String

    Set box = New B160Box
    Set o = box

    ' W1: 数值 → Double 属性
    o.D = 2
    Debug.Print "B160-W1=" & box.D

    ' W2: 数值 → String 属性 (VB6: "2")
    o.S = 2
    Debug.Print "B160-W2=" & box.S

    ' W3: 串 → Double 属性 (VB6: 3.5)
    o.D = "3.5"
    Debug.Print "B160-W3=" & box.D

    ' W4: Double → Integer 属性 (VB6 四舍六入五成双: 2.5 -> 2)
    o.I = 2.5
    Debug.Print "B160-W4=" & box.I

    ' W5: 数值 → Boolean 属性 (非零 = True)
    o.B = 7
    Debug.Print "B160-W5=" & box.B

    ' W6: Boolean 属性 → 数值
    n = box.B
    Debug.Print "B160-W6=" & n

    ' W7: 串 → Integer 属性
    o.I = "42"
    Debug.Print "B160-W7=" & box.I

    ' W8: 变量走一遍 (Variant 源)
    Dim v As Variant
    v = 9.75
    o.I = v
    Debug.Print "B160-W8=" & box.I

    s = box.S
    d = box.D
    Debug.Print "B160-Z=" & s & "/" & d

    Set box = Nothing
    Set o = Nothing
    Debug.Print "B160-DONE"
End Sub
