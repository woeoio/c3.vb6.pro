Option Explicit

' §11.192 (knife 209): Select Case where the CASE LABEL is a Variant and the
' subject is a Date (A*) or where the subject is a Variant and the label is a
' date literal (B*). Oracle: VB6Mini /make, .temp/k209p/P250.bas -> p250.txt
' and P251.bas -> p251.txt:
'   A1 hit  err=0   text label equal to d matches
'   A2 else err=0   text label, other date
'   A3 hit  err=13  unparsable text label: body runs, error 13 stays in Err
'   A4 hit  err=94  Null label: body runs, error 94 stays in Err
'   A5 hit  err=0   numeric label 44259 == d
'   A6 hit  err=0   Variant holding a real date
'   A7 hit  err=0   Variant subject holding text vs Case #3/4/2021#
'   A8 else err=0   Null subject vs a date label: no match, NO error
'   B1 hit  / B2 else   Case Is > Variant label
'   B3 hit              Case lo To hi (both Variant)
'   B4 else             Variant holding Empty is 0, never equals d
'   B5 hit              mixed range 44000 To Variant text label
' Before this knife the Date subject took the int32_t carrier (truncated), and
' a Variant subject with a date label took vb6_VariantToLong (A7 answered 0,
' A8 raised 94 and wrongly matched).
Public Sub Main()
    Dim d As Date
    Dim v As Variant
    Dim lo As Variant
    Dim hi As Variant

    On Error Resume Next
    d = CDate("2021-03-04")

    v = "2021-03-04"
    Err.Clear
    Select Case d
        Case v
            Debug.Print "A1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A1 else err=" & CStr(Err.Number)
    End Select

    v = "2020-01-01"
    Err.Clear
    Select Case d
        Case v
            Debug.Print "A2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A2 else err=" & CStr(Err.Number)
    End Select

    v = "abc"
    Err.Clear
    Select Case d
        Case v
            Debug.Print "A3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A3 else err=" & CStr(Err.Number)
    End Select

    v = Null
    Err.Clear
    Select Case d
        Case v
            Debug.Print "A4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A4 else err=" & CStr(Err.Number)
    End Select

    v = 44259
    Err.Clear
    Select Case d
        Case v
            Debug.Print "A5 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A5 else err=" & CStr(Err.Number)
    End Select

    v = CDate("2021-03-04")
    Err.Clear
    Select Case d
        Case v
            Debug.Print "A6 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A6 else err=" & CStr(Err.Number)
    End Select

    v = "2021-03-04"
    Err.Clear
    Select Case v
        Case #3/4/2021#
            Debug.Print "A7 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A7 else err=" & CStr(Err.Number)
    End Select

    v = Null
    Err.Clear
    Select Case v
        Case #3/4/2021#
            Debug.Print "A8 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "A8 else err=" & CStr(Err.Number)
    End Select

    v = "2021-01-01"
    Err.Clear
    Select Case d
        Case Is > v
            Debug.Print "B1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "B1 else err=" & CStr(Err.Number)
    End Select

    v = "2022-01-01"
    Err.Clear
    Select Case d
        Case Is > v
            Debug.Print "B2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "B2 else err=" & CStr(Err.Number)
    End Select

    lo = "2021-01-01"
    hi = "2021-12-31"
    Err.Clear
    Select Case d
        Case lo To hi
            Debug.Print "B3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "B3 else err=" & CStr(Err.Number)
    End Select

    v = Empty
    Err.Clear
    Select Case d
        Case v
            Debug.Print "B4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "B4 else err=" & CStr(Err.Number)
    End Select

    v = "2021-03-04"
    Err.Clear
    Select Case d
        Case 44000 To v
            Debug.Print "B5 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "B5 else err=" & CStr(Err.Number)
    End Select
End Sub