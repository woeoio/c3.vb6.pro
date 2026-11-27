Option Explicit

' 11.203 (knife 220): DateDiff / DatePart must read "which day" through the same
' carry-aware splitter as DateAdd. Oracle .temp/k220p/P280.bas -> p280.txt (VB6),
' with t = CDate(44259.9999999) = 2021-03-05 00:00:00:
'   DateDiff("d", t, CDate(44260)) = 0        (raw truncation answered 1)
'   DatePart("d", t)               = 5        (raw truncation answered 4)
'   DateAdd("d",1,t) = 3/6 ; ("m",1) = 4/5 ; ("yyyy",1) = 2022/3/5
'   DateAdd("h",-1,t) = 3/4 23:00 ; ("n",-1,t) = 3/4 23:59
'
' Cells are plain integers plus one "has /" shape -> region-independent.
Public Sub Main()
    Dim t As Date
    Dim a As String
    On Error Resume Next

    t = CDate(44259.9999999)

    Debug.Print "A1 day=" & CStr(Day(DateAdd("d", 1, t)))
    Debug.Print "A2 month=" & CStr(Month(DateAdd("m", 1, t)))
    Debug.Print "A3 day=" & CStr(Day(DateAdd("m", 1, t)))
    Debug.Print "A4 year=" & CStr(Year(DateAdd("yyyy", 1, t)))
    Debug.Print "A5 hour=" & CStr(Hour(DateAdd("h", -1, t)))
    Debug.Print "A6 min=" & CStr(Minute(DateAdd("n", -1, t)))

    Debug.Print "B1 diff=" & CStr(DateDiff("s", t, CDate(44260)))
    Debug.Print "B2 diff=" & CStr(DateDiff("d", t, CDate(44260)))
    Debug.Print "B3 diff=" & CStr(DateDiff("h", t, CDate(44260)))
    Debug.Print "B4 diff=" & CStr(DateDiff("n", t, CDate(44260)))
    Debug.Print "B5 diff=" & CStr(DateDiff("s", CDate(44260), t))

    Debug.Print "C1 d=" & CStr(DatePart("d", t))
    Debug.Print "C2 h=" & CStr(DatePart("h", t))
    Debug.Print "C3 n=" & CStr(DatePart("n", t))
    Debug.Print "C4 s=" & CStr(DatePart("s", t))
    Debug.Print "C5 m=" & CStr(DatePart("m", t))
    Debug.Print "C6 y=" & CStr(DatePart("yyyy", t))

    a = CStr(DateAdd("d", 1, t))
    Debug.Print "D1 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "D2 nocolon=" & CStr(InStr(a, ":") = 0)

    ' Weekday 也问同一句（真机 P282: 6 / 5 / 1）
    Debug.Print "E1 wd=" & CStr(Weekday(t))
    Debug.Print "E2 wd=" & CStr(Weekday(t, vbMonday))
    Debug.Print "E3 wd=" & CStr(Weekday(CDate(-0.9999999)))
    Debug.Print "E4 wd=" & CStr(DatePart("w", t))
End Sub