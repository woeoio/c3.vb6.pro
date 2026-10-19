Option Explicit

' <vbeclipse> regression fixture (knife 94, ledger 305 #417).
' Text fed to a date conversion must get the DATE question asked, and the two tiers
' VB6 distinguishes must stay distinguished:
'   CDate(text)    = a date/time writing OR a strict numeric serial (VB6 accepts
'                    CDate("46000") = 46000 and keeps the time of day),
'   DateValue/TimeValue(text) = a date/time writing ONLY -- DateValue("46000") is 13.
' Before this knife CDate's text branch asked the NUMERIC parser, so CDate("2026/1/1")
' answered 2026 (the year, not the day) and CDate("abc") answered 0 with Err still 0,
' while DateValue("2026/1/1") in the very same program was right. Unparseable text is
' 13 Type mismatch in VB6; this repo was silent. Range/overflow is not this knife's
' subject. Val() stays the lenient authority on purpose (pinned as a witness).
' Readings came from the on-machine VB-family runtime: .build/k234/k235/k236 *.vbs.

Sub Main()
    On Error Resume Next
    Dim d As Date
    Dim v As Variant
    Dim r As Long

    ' --- the day, not the year: this is the line this knife exists for
    d = CDate("2026/1/1")
    Debug.Print "B417-01=" & CLng(d)
    ' --- and the time of day survives. 12:45 is deliberately used instead of 14:30:
    '     14:30/24h is not exactly representable as a binary fraction and the Hour/Minute/Second
    '     getters truncate their product (that answer is 29 here while the VB-family runtime says
    '     30 -- ledger #418, a separate cell). 12:45 lands on 17/32 of a day, so this knife's
    '     needle cannot be confused with that other gap.
    d = CDate("2026/1/1 12:45")
    Debug.Print "B417-02=" & Hour(d)
    Debug.Print "B417-03=" & Minute(d)

    ' --- CDate's own fallback tier: a numeric string IS a serial (VB6 does this too)
    d = CDate("46000")
    Debug.Print "B417-04=" & CLng(d)
    d = CDate("&H1")
    Debug.Print "B417-05=" & CLng(d)

    ' --- the three shapes VB6 answers with 13 (before: silent 0 with Err = 0)
    Err.Clear
    d = CDate("abc")
    Debug.Print "B417-06=" & Err.Number
    Err.Clear
    d = CDate("")
    Debug.Print "B417-07=" & Err.Number
    Err.Clear
    d = CDate("123abc")
    Debug.Print "B417-08=" & Err.Number

    ' --- the projections keep their own halves
    Debug.Print "B417-09=" & CLng(DateValue("2026/1/1 14:30"))
    Debug.Print "B417-10=" & Hour(TimeValue("2026/1/1 12:45"))
    ' --- and they do NOT accept a bare serial the way CDate does
    Err.Clear
    r = DateValue("46000")
    Debug.Print "B417-11=" & Err.Number
    Err.Clear
    r = TimeValue("abc")
    Debug.Print "B417-12=" & Err.Number

    ' --- witnesses: nothing else in the family moved
    Debug.Print "B417-13=" & CLng(DateValue("2026/1/1"))
    Debug.Print "B417-14=" & Val("123abc")
    d = CDate("2026/1/1")
    v = d
    Debug.Print "B417-15=" & CLng(CDate(v))
    Err.Clear
    d = CDate("2026/1/1")
    Debug.Print "B417-16=" & Err.Number
    Debug.Print "B417-17=" & Year(d) & "-" & Month(d) & "-" & Day(d)

    Debug.Print "B417-DONE"
End Sub
