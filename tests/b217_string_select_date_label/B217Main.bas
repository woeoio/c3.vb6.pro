Option Explicit

' §11.193 (knife 210): Select Case where the subject is a String and the label
' is a Date. Oracle: VB6Mini /make, .temp/k210p/P252.bas -> p252.txt,
' P253.bas -> p253.txt and P255.bas -> p255.txt.
'   The comparison is done AS TEXT, with the Date label formatted to VB6's own
'   short-date text (what CStr(Date) gives — "2021/3/4" on a zh-CN machine,
'   "3/4/2021" on en-US). Every cell pinned here is invariant across both
'   layouts, so the fixture travels:
'     S1 else      "2021-03-04" == #3/4/2021#      (not the formatted text)
'     S4 else      "abc"        == #3/4/2021#
'     S6 else      ""           == #3/4/2021#
'     T1 hit       "9999"  > #1/1/2021#            ("9" > any digit)
'     T2 else      "0"     > #1/1/2021#
'     T3 hit       subject == CStr(CDate("2021-03-04"))
'     T4 hit       that text < #3/5/2021#          (same layout on both sides)
'     T5 else      "2021-03-05" > #3/4/2021#       ("2" < first digit of the text)
'     U1 hit       "zzz" > #1/1/2021#              ("z" > any digit)
'     U2 else      "zzz" inside #1/1/2021# To #12/31/2021#
'     U3 else      "!"   > #1/1/2021#              ("!" < any digit)
'     U4 else      "zzz" == #1/1/2021#
' Before this knife a String subject with a Date label did not even compile:
'   `BSTR _vb6_select_0 = s; ... vb6_StrCmp(_vb6_select_0, 44259.0)`  (C2440 x5).
Public Sub Main()
    Dim s As String

    On Error Resume Next

    s = "2021-03-04"
    Err.Clear
    Select Case s
        Case #3/4/2021#
            Debug.Print "S1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "S1 else err=" & CStr(Err.Number)
    End Select

    s = "abc"
    Err.Clear
    Select Case s
        Case #3/4/2021#
            Debug.Print "S4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "S4 else err=" & CStr(Err.Number)
    End Select

    s = ""
    Err.Clear
    Select Case s
        Case #3/4/2021#
            Debug.Print "S6 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "S6 else err=" & CStr(Err.Number)
    End Select

    s = "9999"
    Err.Clear
    Select Case s
        Case Is > #1/1/2021#
            Debug.Print "T1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "T1 else err=" & CStr(Err.Number)
    End Select

    s = "0"
    Err.Clear
    Select Case s
        Case Is > #1/1/2021#
            Debug.Print "T2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "T2 else err=" & CStr(Err.Number)
    End Select

    ' Same text VB6 itself produces for that date (locale short date).
    s = CStr(CDate("2021-03-04"))
    Debug.Print "T3 text=" & s
    Err.Clear
    Select Case s
        Case #3/4/2021#
            Debug.Print "T3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "T3 else err=" & CStr(Err.Number)
    End Select

    s = CStr(CDate("2021-03-04"))
    Err.Clear
    Select Case s
        Case Is < #3/5/2021#
            Debug.Print "T4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "T4 else err=" & CStr(Err.Number)
    End Select

    s = "2021-03-05"
    Err.Clear
    Select Case s
        Case Is > #3/4/2021#
            Debug.Print "T5 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "T5 else err=" & CStr(Err.Number)
    End Select

    s = "zzz"
    Err.Clear
    Select Case s
        Case Is > #1/1/2021#
            Debug.Print "U1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "U1 else err=" & CStr(Err.Number)
    End Select

    s = "zzz"
    Err.Clear
    Select Case s
        Case #1/1/2021# To #12/31/2021#
            Debug.Print "U2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "U2 else err=" & CStr(Err.Number)
    End Select

    s = "!"
    Err.Clear
    Select Case s
        Case Is > #1/1/2021#
            Debug.Print "U3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "U3 else err=" & CStr(Err.Number)
    End Select

    s = "zzz"
    Err.Clear
    Select Case s
        Case #1/1/2021#
            Debug.Print "U4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "U4 else err=" & CStr(Err.Number)
    End Select
End Sub