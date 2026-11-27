Option Explicit
Option Compare Text

' §11.194 (knife 211): module-level `Option Compare Text` together with a Date
' label in a Select Case whose subject is a String. Oracle: VB6Mini /make,
' .temp/k211p/P256.bas -> p256.txt:
'   V1 hit   "ABC" vs Case "abc"        (text compare, as expected)
'   V2 else  "ABC" vs Case #3/4/2021#   (Date label still becomes short-date text,
'                                        so a case-insensitive compare cannot match)
'   V3 hit   CStr(CDate(...)) vs Case #3/4/2021#
'   V4 hit   "ZZZ" > #1/1/2021#         ("Z"/"z" > any digit, either layout)
'   V5 else  "!"   > #1/1/2021#
' C3 already answered all five identically before this knife: the Date fold
' (strLabel193 -> vb6_CStrDate) composes with the Text entry (vb6_StrCmpT)
' without a second rule. This fixture exists to keep it that way.
' Cells are pinned locale-invariant only (see §11.193): the answers must not
' depend on whether the short date reads "2021/3/4" or "3/4/2021".
Public Sub Main()
    Dim s As String

    On Error Resume Next

    s = "ABC"
    Err.Clear
    Select Case s
        Case "abc"
            Debug.Print "V1 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "V1 else err=" & CStr(Err.Number)
    End Select

    s = "ABC"
    Err.Clear
    Select Case s
        Case #3/4/2021#
            Debug.Print "V2 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "V2 else err=" & CStr(Err.Number)
    End Select

    s = CStr(CDate("2021-03-04"))
    Err.Clear
    Select Case s
        Case #3/4/2021#
            Debug.Print "V3 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "V3 else err=" & CStr(Err.Number)
    End Select

    s = "ZZZ"
    Err.Clear
    Select Case s
        Case Is > #1/1/2021#
            Debug.Print "V4 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "V4 else err=" & CStr(Err.Number)
    End Select

    s = "!"
    Err.Clear
    Select Case s
        Case Is > #1/1/2021#
            Debug.Print "V5 hit err=" & CStr(Err.Number)
        Case Else
            Debug.Print "V5 else err=" & CStr(Err.Number)
    End Select
End Sub