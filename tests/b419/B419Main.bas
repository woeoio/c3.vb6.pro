Option Explicit

' <vbeclipse> regression fixture (knife 96, ledger 305 #419).
' A DATE is "whole days + a fraction of a day", and that fraction is not exact in binary
' (14:30 = 0.6041666666666666).  The split used to TRUNCATE, so every getter that asked
' "which second of the day is this" could land one second low -- measured against the
' VB-family runtime on this machine (.build/k272_timeparts.vbs, k277_round.vbs,
' k278_roll.vbs): it rounds the second, and Minute(CDate("2026/1/1 14:30")) is 30 there
' while our old code answered 29 (and Second of "23:59:59" answered 58).
' DateAdd is the worse half: it splits, adds, and writes the split BACK, so a second lost
' by truncation stays lost in the stored value.
' One outlet now owns the split (vb6_SecondsOfDay in vb6rtl_date.c); the three getters,
' DateAdd's decomposition and -- for free -- DatePart's "h"/"n"/"s" arms all ask it.
' Deliberately NOT pinned here: the day half at the end of the day (the runtime also rolls
' the DATE across midnight for 23:59:59.9; ours zeroes the seconds-of-day and leaves the
' day part alone) and pre-epoch serials (negative fraction sign) -- both recorded as their
' own ledger cells instead of being guessed at.

Sub Main()
    Dim d As Date
    Dim e As Date
    Dim base45 As Date

    ' --- the defect: 14:30 is not dyadic, truncation answered 14/29/59
    d = CDate("2026/1/1 14:30")
    Debug.Print "B419-01=" & Hour(d)
    Debug.Print "B419-02=" & Minute(d)
    Debug.Print "B419-03=" & Second(d)

    ' --- end of day: truncation lost a second here too
    d = CDate("2026/1/1 23:59:59")
    Debug.Print "B419-04=" & Hour(d) & "/" & Minute(d) & "/" & Second(d)

    ' --- witnesses: answers that were already right must not move
    d = CDate("2026/1/1 14:30:45")
    Debug.Print "B419-05=" & Hour(d) & "/" & Minute(d) & "/" & Second(d)
    Debug.Print "B419-06=" & Minute(CDate("2026/1/1 12:45")) & "/" & Second(CDate("2026/1/1 12:45"))
    Debug.Print "B419-07=" & Hour(CDate("2026/1/1")) & "/" & Minute(CDate("2026/1/1")) & "/" & Second(CDate("2026/1/1"))
    Debug.Print "B419-08=" & Year(d) & "-" & Month(d) & "-" & Day(d)

    ' --- the caliber: the runtime rounds the second (+0.4 does not carry, +0.6 does)
    base45 = CDate("2026/1/1 14:30:45")
    e = base45 + 0.6 / 86400
    Debug.Print "B419-09=" & Second(e)
    e = base45 + 0.4 / 86400
    Debug.Print "B419-10=" & Second(e)

    ' --- DateAdd writes the split back: a truncated second would stay truncated forever
    d = CDate("2026/1/1 14:30")
    e = DateAdd("d", 1, d)
    Debug.Print "B419-11=" & Hour(e) & "/" & Minute(e)
    e = DateAdd("n", 60, d)
    Debug.Print "B419-12=" & Hour(e) & "/" & Minute(e) & "/" & Second(e)
    e = DateAdd("h", -15, d)
    Debug.Print "B419-13=" & Hour(e) & "/" & Minute(e) & " day=" & Day(e) & "-" & Month(e) & "-" & Year(e)
    e = DateAdd("s", 1, CDate("2026/1/1 23:59:59"))
    Debug.Print "B419-14=" & Hour(e) & "/" & Minute(e) & "/" & Second(e) & " day=" & Day(e)

    ' --- DatePart's h/n/s arms delegate to the same three getters
    Debug.Print "B419-15=" & DatePart("n", CDate("2026/1/1 14:30")) & "/" & DatePart("s", CDate("2026/1/1 14:30"))

    ' --- a value that keeps real sub-second parts still round-trips through the day serial
    d = CDate("2026/1/1 06:15:30")
    Debug.Print "B419-16=" & Hour(d) & "/" & Minute(d) & "/" & Second(d)

    Debug.Print "B419-DONE"
End Sub
