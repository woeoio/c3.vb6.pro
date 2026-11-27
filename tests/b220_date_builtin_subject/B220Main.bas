Option Explicit

Private Type TFix
    a(0 To 1) As Date
End Type

Private Type TDyn
    a() As Date
End Type

Private Type TRec
    d As Date
    a() As Date
End Type

' §11.196 (knife 213): the SUBJECT of a Select Case is a date BUILTIN. In VB6
' Now / Date / Time / DateSerial / DateValue / TimeValue / DateAdd / DateDiff /
' DatePart / TimeSerial return Variant (Date), not Date, so a String label is
' compared AS TEXT while a Date label still compares as a date.
' Oracle: VB6Mini /make, .temp/k213p/P260.bas -> p260.txt, P261.bas -> p261.txt,
' P262.bas -> p262.txt:
'   F1..F4 hit   Date element inside a UDT member array (fixed and dynamic)
'                as the subject — already right before this knife.
'   G1 G2 G5 hit With-block member / With-block array element — already right.
'   G3 hit       CDate(s) subject (CDate really returns Date)
'   G4 else      DateValue(s) subject with the text "2021-03-04": the subject is a
'                Variant, so the label is text ("2021/3/4" <> "2021-03-04")
'   H1 else      DateSerial(2021,3,4) vs Case "2021-03-04"   (text compare)
'   H2 hit       DateSerial(2021,3,4) vs Case #3/4/2021#
'   H3 hit       CDate(...) vs Case Is > #1/1/2000#
'   H5 hit       Now vs Case Is > #1/1/2000#
'   H6 hit       Date vs Case CStr(Date)
' Known defect left for the next knife: `CStr(Now)` in C3 prints the bare serial
' ("46353.4167824074") instead of VB6's "2026/11/27 9:59:29", so the H4 cell
' (`Select Case Now: Case CStr(Now)`) is NOT pinned here.
Public Sub Main()
    Dim u As TFix
    Dim v As TDyn
    Dim w As TRec
    Dim s As String

    On Error Resume Next
    u.a(0) = CDate("2021-03-04")
    ReDim v.a(0 To 1)
    v.a(0) = CDate("2021-03-04")
    w.d = CDate("2021-03-04")
    ReDim w.a(0 To 1)
    w.a(0) = CDate("2021-03-04")
    s = "2021-03-04"

    Err.Clear
    Select Case u.a(0)
        Case #3/4/2021#
            Debug.Print "F1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "F1 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case u.a(0)
        Case "2021-03-04"
            Debug.Print "F2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "F2 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case v.a(0)
        Case #3/4/2021#
            Debug.Print "F3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "F3 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case v.a(0)
        Case Is > #1/1/2021#
            Debug.Print "F4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "F4 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    With w
        Select Case .d
            Case #3/4/2021#
                Debug.Print "G1 hit err=" & CStr(Err.Number)
            Case Else
                Debug.Print "G1 else err=" & CStr(Err.Number)
        End Select
    End With

    Err.Clear
    With w
        Select Case .d
            Case "2021-03-04"
                Debug.Print "G2 hit err=" & CStr(Err.Number)
            Case Else
                Debug.Print "G2 else err=" & CStr(Err.Number)
        End Select
    End With

    Err.Clear
    Select Case CDate(s)
        Case #3/4/2021#
            Debug.Print "G3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "G3 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case DateValue(s)
        Case "2021-03-04"
            Debug.Print "G4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "G4 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    With w
        Select Case .a(0)
            Case #3/4/2021#
                Debug.Print "G5 hit err=" & CStr(Err.Number)
            Case Else
                Debug.Print "G5 else err=" & CStr(Err.Number)
        End Select
    End With

    Err.Clear
    Select Case DateSerial(2021, 3, 4)
        Case "2021-03-04"
            Debug.Print "H1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "H1 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case DateSerial(2021, 3, 4)
        Case #3/4/2021#
            Debug.Print "H2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "H2 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case CDate("2021-03-04")
        Case Is > #1/1/2021#
            Debug.Print "H3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "H3 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case Now
        Case Is > #1/1/2000#
            Debug.Print "H5 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "H5 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case Date
        Case CStr(Date)
            Debug.Print "H6 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "H6 else err=" & CStr(Err.Number)
    End Select
End Sub