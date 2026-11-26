Option Explicit

' §11.189 (knife 206): the Variant on the right of a Date target may also come
' from a call. Oracle: VB6Mini /make .temp/k206p/P246.bas -> p246.txt, nine lines,
' identical step by step with the local-variable group of §11.188:
'   W1/W2 = 44259 err=0   (a Variant-returning call holding a text date; scalar
'                          and array element)
'   W3    = 45000 err=0   (the same call holding a numeric date)
'   W4/W5 = 45000 err=13  ("abc" keeps the old value)
'   W6    = 44259 err=0   (the call holding a VT_DATE)
'   W7    = 0 err=0       (Empty)
'   W8    = 45000 err=94  (Null keeps the old value)
'   W9    = vt=8 tn=String (the call really returns a Variant)
' Before this knife all of them emitted vb6_VariantToDouble(call): the text date
' silently became 2021, "abc" and Null silently became 0 with Err=0.
Private Function FText() As Variant
    FText = "2021-03-04"
End Function

Private Function FNum() As Variant
    FNum = 45000
End Function

Private Function FBad() As Variant
    FBad = "abc"
End Function

Private Function FDate() As Variant
    FDate = CDate("2021-03-04")
End Function

Private Function FEmpty() As Variant
    FEmpty = Empty
End Function

Private Function FNull() As Variant
    FNull = Null
End Function

Public Sub Main()
    Dim d As Date
    Dim q() As Date

    On Error Resume Next
    ReDim q(0 To 1)

    Err.Clear
    d = FText()
    Debug.Print "W1 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    Err.Clear
    q(0) = FText()
    Debug.Print "W2 d=" & CStr(CDbl(q(0))) & " err=" & CStr(Err.Number)

    Err.Clear
    d = FNum()
    Debug.Print "W3 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    d = 45000
    Err.Clear
    d = FBad()
    Debug.Print "W4 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    q(0) = 45000
    Err.Clear
    q(0) = FBad()
    Debug.Print "W5 d=" & CStr(CDbl(q(0))) & " err=" & CStr(Err.Number)

    Err.Clear
    d = FDate()
    Debug.Print "W6 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    d = 45000
    Err.Clear
    d = FEmpty()
    Debug.Print "W7 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    d = 45000
    Err.Clear
    d = FNull()
    Debug.Print "W8 d=" & CStr(CDbl(d)) & " err=" & CStr(Err.Number)

    Debug.Print "W9 vt=" & CStr(VarType(FText())) & " tn=" & TypeName(FText())
End Sub