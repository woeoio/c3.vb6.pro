Option Explicit

' 11.212 (knife 229): whole-array argument to Hex/Oct is error 13 (Type mismatch) and the
' target keeps its old value. Real machine (.temp/k225p/P320/P321):
'   Hex(a) / Oct(a) / Hex(a()) / Hex(b) (String elements) / Hex(c) (never ReDim'd) all 13;
'   Hex(a(1)) is still FFFFFFFB; a Variant holding the array is 13 too.
' Before: the array name was handed to vb6_Hex(double) as a pointer -> 0 (or the descriptor
' address printed as a number) with error 0.
Public Sub Main()
    Dim a() As Long
    Dim b() As String
    Dim c() As Long
    Dim s As String
    Dim v As Variant

    ReDim a(1 To 3)
    a(1) = -5
    a(2) = 255
    a(3) = 0
    ReDim b(0 To 1)
    b(0) = "12"
    b(1) = "x"

    On Error Resume Next
    Err.Clear: s = "keep": s = Hex(a): Debug.Print "H1=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Oct(a): Debug.Print "H2=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Hex(a()): Debug.Print "H3=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Hex(b): Debug.Print "H4=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Hex(c): Debug.Print "H5=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Hex(a(1)): Debug.Print "H6=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Hex(255): Debug.Print "H7=[" & s & "]/" & Err.Number
    Err.Clear: v = a: s = "keep": s = Hex(v): Debug.Print "H8=[" & s & "]/" & Err.Number
    Err.Clear: v = a: s = "keep": s = Oct(v): Debug.Print "H9=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Oct(-5): Debug.Print "H10=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Hex(a(2)): Debug.Print "H11=[" & s & "]/" & Err.Number
    Err.Clear: s = "keep": s = Hex(CLng(a(1))): Debug.Print "H12=[" & s & "]/" & Err.Number
    On Error GoTo 0
End Sub