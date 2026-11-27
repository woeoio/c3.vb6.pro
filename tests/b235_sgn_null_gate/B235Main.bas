Option Explicit

' 11.211 (knife 228): Sgn has no Variant outlet. Real machine (.temp/k225p/P313):
'   Sgn(Null) = error 94 (target unchanged), Sgn("")/Sgn("x") = error 13,
'   Sgn("2.5") = 1, Sgn(Empty) = 0, Sgn(True) = -1.
' Abs/Int/Fix are the opposite: Null in -> Null out (P314 K2/K3, VarType 1), so they keep
' their *V outlets; only the plain numeric extract had to change for Sgn.
'
' Before: vb6_Sgn(vb6_VariantToDouble(v)) -- VariantToDouble must stay lenient (Val family),
' so Null and non-numeric text both answered 0 with error 0.
Public Sub Main()
    Dim v As Variant
    Dim d As Double

    On Error Resume Next

    v = Null
    Err.Clear: d = -99: d = Sgn(v): Debug.Print "G1=" & d & "/" & Err.Number
    Err.Clear: d = -99: d = Abs(v): Debug.Print "G2=" & d & "/" & Err.Number
    Err.Clear: d = -99: d = Int(v): Debug.Print "G3=" & d & "/" & Err.Number
    Err.Clear: d = -99: d = Fix(v): Debug.Print "G4=" & d & "/" & Err.Number

    v = ""
    Err.Clear: d = -99: d = Sgn(v): Debug.Print "G5=" & d & "/" & Err.Number
    v = "x"
    Err.Clear: d = -99: d = Sgn(v): Debug.Print "G6=" & d & "/" & Err.Number
    v = "2.5"
    Err.Clear: d = -99: d = Sgn(v): Debug.Print "G7=" & d & "/" & Err.Number

    Err.Clear: d = -99: d = Sgn(""): Debug.Print "G8=" & d & "/" & Err.Number
    Err.Clear: d = -99: d = Sgn(Null): Debug.Print "G9=" & d & "/" & Err.Number

    Err.Clear: d = Abs(True): Debug.Print "G10=" & d & "/" & Err.Number
    Err.Clear: d = Sgn(True): Debug.Print "G11=" & d & "/" & Err.Number

    ' Abs/Int/Fix 的 Null 传播仍是 Null（这一刀不许把它们改成 94）
    v = Null
    Debug.Print "N1=" & IsNull(Abs(v)) & IsNull(Int(v)) & IsNull(Fix(v))
    ' Empty 四枚都是 0
    v = Empty
    Debug.Print "N2=" & Abs(v) & Int(v) & Fix(v) & Sgn(v)
    ' 负值/零/文本数值
    Err.Clear: Debug.Print "N3=" & Sgn(-3.5) & Sgn(0) & Sgn("2.5") & "/" & Err.Number

    On Error GoTo 0
End Sub