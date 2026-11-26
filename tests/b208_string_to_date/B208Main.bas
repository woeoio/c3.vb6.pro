Option Explicit

' 11.183 (knife 200): String -> Date at the assignment station.
' Oracle: VB6Mini /make .temp/k200p/P241.bas -> p241.txt, eight lines:
'   S1..S5 = 44259 err=0 (literal / array elem / String var / array elem from var / CDate)
'   S6..S7 = keeps the previous value (45000) and raises 13 (bad text, scalar / elem)
'   S8     = vt=7 tn=Date
' Before this knife the array-element form emitted a raw BSTR into a double slot
' (cl C2440, no exe) and all three forms reported a bogus VB3003 warning.
Public Sub Main()
    Dim d As Date
    Dim q() As Date
    Dim s As String

    On Error Resume Next

    Err.Clear
    d = "2021-03-04"
    Debug.Print "S1 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    ReDim q(0 To 1)
    Err.Clear
    q(0) = "2021-03-04"
    Debug.Print "S2 d=" & CStr(CDbl(q(0))) & " err=" & CStr(Err.Number)

    Err.Clear
    s = "2021-03-04"
    d = s
    Debug.Print "S3 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    Err.Clear
    q(1) = s
    Debug.Print "S4 d=" & CStr(CDbl(q(1))) & " err=" & CStr(Err.Number)

    Err.Clear
    d = CDate("2021-03-04")
    Debug.Print "S5 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    d = 45000
    Err.Clear
    d = "abc"
    Debug.Print "S6 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    q(0) = 45000
    Err.Clear
    q(0) = "abc"
    Debug.Print "S7 d=" & CStr(CDbl(q(0))) & " err=" & CStr(Err.Number)

    Err.Clear
    Debug.Print "S8 vt=" & CStr(VarType(d)) & " tn=" & TypeName(d)
End Sub
