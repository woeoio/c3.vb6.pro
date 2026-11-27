Option Explicit

' 11.213 (knife 230): Abs/Int/Fix on a whole Variant slot are STRICT on every non-Variant
' consumer (string target, CStr, concatenation) -- Null is 94, non-numeric text is 13 --
' while the Variant consumer keeps Null (w = Abs(Null) is Null).
' Real machine (.temp/k225p/P323):
'   S1/S2 = [keep]/13, S3 = [2.5]/0, S4 = [keep]/94, S5 = [0]/0
'   T1/T2 = [keep]/13, T3 = [keep]/94, T4 = [-2]/0
'   U1..U4 = same as S1..S4 (direct string assignment), V1..V3 = concatenation
'   W1 = []/0/True (Variant target keeps Null)
Public Sub Main()
    Dim v As Variant
    Dim s As String
    Dim w As Variant

    On Error Resume Next

    v = "": Err.Clear: s = "keep": s = CStr(Abs(v)): Debug.Print "S1=[" & s & "]/" & Err.Number
    v = "x": Err.Clear: s = "keep": s = CStr(Abs(v)): Debug.Print "S2=[" & s & "]/" & Err.Number
    v = "2.5": Err.Clear: s = "keep": s = CStr(Abs(v)): Debug.Print "S3=[" & s & "]/" & Err.Number
    v = Null: Err.Clear: s = "keep": s = CStr(Abs(v)): Debug.Print "S4=[" & s & "]/" & Err.Number
    v = Empty: Err.Clear: s = "keep": s = CStr(Abs(v)): Debug.Print "S5=[" & s & "]/" & Err.Number

    v = "": Err.Clear: s = "keep": s = CStr(Int(v)): Debug.Print "T1=[" & s & "]/" & Err.Number
    v = "x": Err.Clear: s = "keep": s = CStr(Fix(v)): Debug.Print "T2=[" & s & "]/" & Err.Number
    v = Null: Err.Clear: s = "keep": s = CStr(Int(v)): Debug.Print "T3=[" & s & "]/" & Err.Number
    v = "-2.5": Err.Clear: s = "keep": s = CStr(Fix(v)): Debug.Print "T4=[" & s & "]/" & Err.Number

    v = "": Err.Clear: s = "keep": s = Abs(v): Debug.Print "U1=[" & s & "]/" & Err.Number
    v = "x": Err.Clear: s = "keep": s = Abs(v): Debug.Print "U2=[" & s & "]/" & Err.Number
    v = "-2.5": Err.Clear: s = "keep": s = Abs(v): Debug.Print "U3=[" & s & "]/" & Err.Number
    v = Null: Err.Clear: s = "keep": s = Abs(v): Debug.Print "U4=[" & s & "]/" & Err.Number

    v = "": Err.Clear: s = "keep": s = "n=" & Abs(v): Debug.Print "V1=[" & s & "]/" & Err.Number
    v = "x": Err.Clear: s = "keep": s = "n=" & Abs(v): Debug.Print "V2=[" & s & "]/" & Err.Number
    v = "-2.5": Err.Clear: s = "keep": s = "n=" & Abs(v): Debug.Print "V3=[" & s & "]/" & Err.Number

    v = Null: Err.Clear: w = "keep": w = Abs(v): Debug.Print "W1=[" & w & "]/" & Err.Number & "/" & IsNull(w)

    v = "": Err.Clear: s = "keep": s = Sgn(v) & "": Debug.Print "X1=[" & s & "]/" & Err.Number

    On Error GoTo 0
End Sub