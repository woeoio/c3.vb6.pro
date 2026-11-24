Option Explicit

Public Sub Main()
    Dim box As B158Box
    Dim o As Object
    Dim n As Long
    Dim s As String
    Dim v As Variant

    Set box = New B158Box
    Set o = box

    ' L1/L2/L3: Variant 型属性读进 Long / 参与算术 / 折成串
    box.V = 11
    n = o.V
    Debug.Print "B158-L1=" & n
    n = o.V + 1
    Debug.Print "B158-L2=" & n
    s = CStr(o.V)
    Debug.Print "B158-L3=" & s

    ' L4: 属性里揣着串时按 VB6 的隐式转换取数
    box.V = "42"
    n = o.V
    Debug.Print "B158-L4=" & n

    ' L5: 读进 Variant —— 型别必须是服务端交回的那一档 (8 = VT_BSTR)
    v = o.V
    Debug.Print "B158-L5=" & VarType(v) & "/" & v

    Set box = Nothing
    Set o = Nothing
    Debug.Print "B158-DONE"
End Sub
