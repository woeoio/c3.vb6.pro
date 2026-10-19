Option Explicit

' <vbeclipse> regression fixture (knife 95, ledger 305 #418).
' The Is* predicates asked ONLY the Variant's vt tag, so any text answered "no":
' IsDate("2026/1/1") and IsNumeric("123") were both False while VB6 answers True for
' both. The text tier now asks the same two parse outlets the conversions ask (knife 93
' for numbers, knife 94 for dates) -- but as a SUCCESS QUESTION, so it must never raise
' (Err stays 0 through every line below; that is a needle, not a side effect).
' The two accept-sets are genuinely different and both are pinned: a time-only text IS a
' date ("14:30") while a date text is NOT a number ("2026/1/1").
' Readings came from the on-machine VB-family runtime: .build/k253_isfam.vbs + k262_ispins.vbs.
' The last needle is the reason this cell exists: VBFlexGrid sorts its date columns with
'   If IsDate(Text) Then ... = CDate(Text)
' and that Then-arm could never run while IsDate was text-blind -- knife 94 fixed CDate,
' knife 95 opens the door in front of it.

Sub Main()
    On Error Resume Next
    Dim s As String
    Dim v As Variant
    Dim d As Date
    Dim r As Long

    ' --- IsDate: the three shapes VB6 distinguishes
    Debug.Print "B418-01=" & IsDate("2026/1/1")
    Debug.Print "B418-02=" & IsDate("14:30")
    Debug.Print "B418-03=" & IsDate("46000")
    Debug.Print "B418-04=" & IsDate("abc")
    Debug.Print "B418-05=" & IsDate("")
    ' --- witnesses: the non-text tiers were already right and must not move
    d = CDate("2026/1/1")
    Debug.Print "B418-06=" & IsDate(d)
    Debug.Print "B418-07=" & IsDate(123)

    ' --- IsNumeric: text now asks the strict numeric outlet (knife 93's owner)
    Debug.Print "B418-08=" & IsNumeric("123")
    Debug.Print "B418-09=" & IsNumeric("&HFF")
    Debug.Print "B418-10=" & IsNumeric("123abc")
    Debug.Print "B418-11=" & IsNumeric("2026/1/1")
    Debug.Print "B418-12=" & IsNumeric(123)
    Debug.Print "B418-13=" & IsNumeric(d)

    ' --- the Variant route asks the SAME question as the String route
    s = "123"
    v = s
    Debug.Print "B418-14=" & IsNumeric(s) & "/" & IsNumeric(v)
    s = "2026/1/1"
    v = s
    Debug.Print "B418-15=" & IsDate(s) & "/" & IsDate(v)

    ' --- a predicate is not a conversion: none of the questions above may raise
    Err.Clear
    r = 0
    If IsDate("abc") Then r = r + 1
    If IsNumeric("123abc") Then r = r + 1
    Debug.Print "B418-16=" & Err.Number & "/" & r

    ' --- the door in front of CDate (VBFlexGrid's shape)
    s = "2026/1/1"
    If IsDate(s) Then
        d = CDate(s)
        r = CLng(d)
    Else
        r = -1
    End If
    Debug.Print "B418-17=" & r

    Debug.Print "B418-DONE"
End Sub
