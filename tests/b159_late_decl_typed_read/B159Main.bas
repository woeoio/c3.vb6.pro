Option Explicit

Public gTerms As Long

Public Sub Main()
    Dim box As B159Box
    Dim o As Object
    Dim n As Long
    Dim d As Double
    Dim s As String
    Dim v As Variant
    Dim it As Object

    Set box = New B159Box
    Set o = box
    box.S = "101"
    box.D = 1.5

    ' S1/S2/S3: String 属性读进 Long / 参与算术 / 读进 Double
    n = o.S
    Debug.Print "B159-S1=" & n
    n = o.S + 1
    Debug.Print "B159-S2=" & n
    d = o.S
    Debug.Print "B159-S3=" & d

    ' S4/S5: Double 属性读进 Double / 读进 Long
    d = o.D
    Debug.Print "B159-S4=" & d
    n = o.D
    Debug.Print "B159-S5=" & n

    ' S6: String 属性进 Variant
    v = o.S
    Debug.Print "B159-S6=" & VarType(v) & "/" & v

    ' S7: Object 属性读出来再取成员
    Set it = New B159Item
    it.ID = 77
    Set box.O = it
    Set it = Nothing
    Set it = o.O
    Debug.Print "B159-S7=" & it.ID

    Set it = Nothing
    Set box = Nothing
    Set o = Nothing
    Debug.Print "B159-Z-terms=" & gTerms
    Debug.Print "B159-DONE"
End Sub
