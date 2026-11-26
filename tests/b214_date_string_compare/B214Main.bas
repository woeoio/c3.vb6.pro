Option Explicit

' §11.190 (knife 207): a Date compared with a String compares DATE SERIALS, not
' formatted text. Oracle: VB6Mini /make .temp/k207p/P248.bas -> p248.txt, seven
' lines:
'   Y1 = 1  (d > "2020-12-31", scalar Date vs text)
'   Y2 = 1  (d = "2021-03-04", equality: text parses to the same serial)
'   Y3 = 1  (q(0) < "2021-03-05", Date ARRAY ELEMENT vs text)
'   Y4 = 1  (text on the left, Date on the right)
'   Y5 = 1 err=13  (unparsable text: raises 13, branch still taken)
'   Y6 = 1  (Select Case d: Case "2021-03-04")
'   Y7 = 0  (d > "2021-03-04 12:00:00 PM": midnight vs noon)
' Before this knife every comparison was emitted as
' vb6_StrCmp(vb6_CStrDate(d), "…") — comparing the FORMATTED text, so Y2 and Y3
' answered 0, Y7 answered 1, and the Select Case selector was truncated to int32_t.
Public Sub Main()
    Dim d As Date
    Dim q() As Date

    On Error Resume Next
    ReDim q(0 To 1)

    d = CDate("2021-03-04")
    q(0) = d

    Err.Clear
    If d > "2020-12-31" Then
        Debug.Print "Y1 gt=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Y1 gt=0 err=" & CStr(Err.Number)
    End If

    Err.Clear
    If d = "2021-03-04" Then
        Debug.Print "Y2 eq=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Y2 eq=0 err=" & CStr(Err.Number)
    End If

    Err.Clear
    If q(0) < "2021-03-05" Then
        Debug.Print "Y3 lt=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Y3 lt=0 err=" & CStr(Err.Number)
    End If

    Err.Clear
    If "2020-12-31" < d Then
        Debug.Print "Y4 rev=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Y4 rev=0 err=" & CStr(Err.Number)
    End If

    Err.Clear
    If d > "abc" Then
        Debug.Print "Y5 bad=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Y5 bad=0 err=" & CStr(Err.Number)
    End If

    Err.Clear
    Select Case d
        Case "2021-03-04"
            Debug.Print "Y6 sel=1 err=" & CStr(Err.Number)
        Case Else
            Debug.Print "Y6 sel=0 err=" & CStr(Err.Number)
    End Select

    Err.Clear
    If d > "2021-03-04 12:00:00 PM" Then
        Debug.Print "Y7 time=1 err=" & CStr(Err.Number)
    Else
        Debug.Print "Y7 time=0 err=" & CStr(Err.Number)
    End If
End Sub