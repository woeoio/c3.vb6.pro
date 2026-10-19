Option Explicit

' <vbeclipse> regression fixture (knife 97, ledger 305 #422).
' DatePart used to answer the wrong question in two arms:
'   (1) case vb6_di_day fell through into the day-of-year body, so DatePart("d", ...) answered
'       "which day of the year" (2026/3/1 -> 60).  The runtime answers 1 there -- "d" is
'       day-of-month and only "y" is day-of-year (ruler .build/k309_datepart.vbs).
'   (2) ("ww", ...) copied a formula in place, (dayOfYear-1)/7+1, and threw away
'       firstDayOfWeek / firstWeekOfYear -- measured against the runtime that disagreed on
'       208 of 365 days in 2026 and 52 of 366 days in 2024 (.build/k310_ww.vbs).
' One outlet vb6_WeekOfYear now owns the week question.  Its caliber is fitted day-by-day
' against the runtime over 14 years x 21 (firstDayOfWeek, firstWeekOfYear) combos =
' 107394 days with zero mismatch (.build/k321/k324/k330_ww_grid*.vbs + k323_ww_fit.py).
' Every number below is the runtime's own reading (k331_needle_readings.txt /
' k335_more_readings.txt), not our output.  B422-13..16 are the leap witnesses: 2023-12-29
' is a Friday (week start for fdow=6) whose aligned week has 4 days in 2024, yet the runtime
' keeps the boundary day itself as 2023's last week number -- because 2024 is a leap year.
' Deliberately NOT pinned here: DateDiff("ww") still counts whole 7-day spans and ignores the
' same two parameters (its own ledger cell), and the day half at the end of the day stays #420.

Sub Main()
    Dim a As Date
    Dim b As Date

    ' --- arm (1): "d" is day-of-month, "y" is day-of-year
    a = CDate("2026/3/1")
    Debug.Print "B422-01=" & DatePart("d", a) & "/" & Day(a)
    Debug.Print "B422-02=" & DatePart("y", a)
    Debug.Print "B422-03=" & DatePart("m", a)
    Debug.Print "B422-04=" & DatePart("q", a)
    Debug.Print "B422-05=" & DatePart("yyyy", a)
    Debug.Print "B422-06=" & DatePart("w", a)

    ' --- arm (2): the default pair is (vbSunday, vbFirstJan1)
    b = CDate("2024/1/1")
    Debug.Print "B422-07=" & DatePart("ww", b)
    b = CDate("2024/1/7")
    Debug.Print "B422-08=" & DatePart("ww", b)
    b = CDate("2026/1/4")
    Debug.Print "B422-09=" & DatePart("ww", b)
    b = CDate("2026/3/1")
    Debug.Print "B422-10=" & DatePart("ww", b)
    b = CDate("2024/6/15")
    Debug.Print "B422-11=" & DatePart("ww", b)

    ' --- the two parameters now participate (2024-12-29 is a Sunday = a week start)
    b = CDate("2024-12-29")
    Debug.Print "B422-12=" & DatePart("ww", b, 1, 1) & "/" & DatePart("ww", b, 1, 2) & "/" & DatePart("ww", b, 1, 3)
    b = CDate("2024-12-31")
    Debug.Print "B422-13=" & DatePart("ww", b, 1, 1) & "/" & DatePart("ww", b, 1, 2) & "/" & DatePart("ww", b, 1, 3)

    ' --- leap witnesses: 2023-12-29 (Friday) keeps 2023's last number, 30/31 already belong to 2024
    b = CDate("2023-12-29")
    Debug.Print "B422-14=" & DatePart("ww", b, 6, 2)
    b = CDate("2023-12-30")
    Debug.Print "B422-15=" & DatePart("ww", b, 6, 2)
    b = CDate("2023-12-31")
    Debug.Print "B422-16=" & DatePart("ww", b, 6, 2) & "/" & DatePart("ww", b, 1, 2)

    ' --- head witnesses: days before week 1 answer the previous year's last week
    b = CDate("2026/1/1")
    Debug.Print "B422-17=" & DatePart("ww", b) & "/" & DatePart("ww", b, 1, 2) & "/" & DatePart("ww", b, 1, 3)
    b = CDate("2032/1/1")
    Debug.Print "B422-18=" & DatePart("ww", b, 1, 2)
    b = CDate("2031-12-31")
    Debug.Print "B422-19=" & DatePart("ww", b, 1, 2)

    ' --- a Monday start and a Saturday start on the same date (three-arg form keeps fwoy default)
    b = CDate("2026/3/1")
    Debug.Print "B422-20=" & DatePart("ww", b, 2, 1) & "/" & DatePart("ww", b, 2, 2)
    Debug.Print "B422-21=" & DatePart("ww", b, 7, 3)
    b = CDate("2024-12-31")
    Debug.Print "B422-22=" & DatePart("ww", b, 2, 1)
    b = CDate("2023-12-31")
    Debug.Print "B422-23=" & DatePart("ww", b, 2, 1)
    b = CDate("2026/3/1")
    Debug.Print "B422-24=" & DatePart("ww", b, 1, 2) & "/" & DatePart("ww", b, 1, 3)

    Debug.Print "B422-DONE"
End Sub
