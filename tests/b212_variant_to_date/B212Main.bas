Option Explicit

' §11.188 (knife 205): a Variant on the right of a Date target goes through
' CDate's own conversion, not through the numeric-only unwrap.
' Oracle: VB6Mini /make .temp/k205p/P245.bas -> p245.txt, nine lines:
'   V1..V4 = 44259 / 44259 / 45000 / 45000 err=0 (Variant holding a text date,
'            then a Variant holding a numeric date; scalar and array element)
'   V5..V6 = keeps the value (45000) and raises 13 (Variant holding "abc")
'   V7     = 44259 err=0, the Variant itself is VT_DATE (7)
'   V8     = Empty assigns 0 (err=0)
'   V9     = Null raises 94 and keeps the old value (45000)
' Before this knife all of them emitted vb6_VariantToDouble(v): the text date
' silently became 2021, and "abc"/Null silently became 0 with Err=0.
Public Sub Main()
    Dim v As Variant
    Dim d As Date
    Dim q() As Date

    On Error Resume Next

    v = "2021-03-04"
    Err.Clear
    d = v
    Debug.Print "V1 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    ReDim q(0 To 1)
    Err.Clear
    q(0) = v
    Debug.Print "V2 d=" & CStr(CDbl(q(0))) & " err=" & CStr(Err.Number)

    v = 45000
    Err.Clear
    d = v
    Debug.Print "V3 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    Err.Clear
    q(1) = v
    Debug.Print "V4 d=" & CStr(CDbl(q(1))) & " err=" & CStr(Err.Number)

    d = 45000
    v = "abc"
    Err.Clear
    d = v
    Debug.Print "V5 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    q(0) = 45000
    Err.Clear
    q(0) = v
    Debug.Print "V6 d=" & CStr(CDbl(q(0))) & " err=" & CStr(Err.Number)

    v = CDate("2021-03-04")
    Err.Clear
    d = v
    Debug.Print "V7 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number) & " vt=" & CStr(VarType(v))

    v = Empty
    Err.Clear
    d = v
    Debug.Print "V8 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    v = Null
    d = 45000
    Err.Clear
    d = v
    Debug.Print "V9 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)
End Sub