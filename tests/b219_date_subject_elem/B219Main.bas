Option Explicit

Private Type TRec
    d As Date
End Type

' §11.195 (knife 212): the SUBJECT of a Select Case is a Date ARRAY ELEMENT or a
' UDT FIELD. Oracle: VB6Mini /make, .temp/k212p/P259.bas -> p259.txt:
'   E1 hit   Select Case q(0)  Case #3/4/2021#
'   E2 hit   Select Case q(0)  Case "2021-03-04"     (text label parsed as a date)
'   E3 hit   Select Case q(0)  Case Is > #1/1/2021#
'   E4 hit   Select Case u.d   Case #3/4/2021#
'   E5 hit   Select Case u.d   Case "2021-03-04"
' Before this knife `inferExprType(q(0))` answers Variant, so the subject took the
' String carrier (all labels were strings/date text) and C3 emitted
'   `BSTR _vb6_select_1 = VB6_SA_AT(double, q, 0);`   (C2440).
' The subject judgement now asks the same three authorities as the binary-compare
' operand side (§11.190): inferExprType, then the bare-name array table, then
' inferUdtFieldVb6Type (with its isArray flag).
Public Sub Main()
    Dim q() As Date
    Dim u As TRec

    On Error Resume Next
    ReDim q(0 To 1)
    q(0) = CDate("2021-03-04")
    u.d = CDate("2021-03-04")

    Err.Clear
    Select Case q(0)
        Case #3/4/2021#
            Debug.Print "E1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "E1 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case q(0)
        Case "2021-03-04"
            Debug.Print "E2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "E2 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case q(0)
        Case Is > #1/1/2021#
            Debug.Print "E3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "E3 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case u.d
        Case #3/4/2021#
            Debug.Print "E4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "E4 else err=" & CStr(Err.Number)
    End Select

    Err.Clear
    Select Case u.d
        Case "2021-03-04"
            Debug.Print "E5 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "E5 else err=" & CStr(Err.Number)
    End Select
End Sub