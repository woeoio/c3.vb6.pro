Option Explicit

' 11.202 (knife 219): DateAdd must take both halves (day, second-of-day) from the
' same carry-aware splitter. Oracle .temp/k219p/P277.bas -> p277*.txt / P278.bas ->
' p278.txt (VB6):
'   DateAdd("s", 1, CDate(44259.9999999)) = 2021/3/5 0:00:01   (never 3/4)
'   44259.9999999 itself = 2021-03-05 00:00:00  (Day=5, Hour=0)
'   DateAdd("s", -1, CDate(44260.0000001)) = 2021/3/4 23:59:59
' Before this knife DateAdd mixed a raw (int32)date with the carry-aware seconds,
' so every result near midnight landed one day early.
'
' Cells are plain integers plus "has /" / "has :" -> region-independent.
Public Sub Main()
    Dim t As Date
    Dim a As String
    On Error Resume Next

    t = CDate(44259.9999999)
    Debug.Print "D1 day=" & CStr(Day(t))
    Debug.Print "D2 hour=" & CStr(Hour(t))

    Debug.Print "A1 day=" & CStr(Day(DateAdd("s", 1, CDate(44259.9999999))))
    Debug.Print "A2 hour=" & CStr(Hour(DateAdd("s", 1, CDate(44259.9999999))))
    Debug.Print "A3 sec=" & CStr(Second(DateAdd("s", 1, CDate(44259.9999999))))
    Debug.Print "A4 day=" & CStr(Day(DateAdd("s", 1, CDate(0.9999999))))
    Debug.Print "A5 sec=" & CStr(Second(DateAdd("s", 1, CDate(0.9999999))))
    Debug.Print "A6 sec=" & CStr(Second(DateAdd("s", -1, CDate(44260.0000001))))
    Debug.Print "A7 day=" & CStr(Day(DateAdd("s", -1, CDate(44260.0000001))))
    Debug.Print "A8 sec=" & CStr(Second(DateAdd("n", 1, CDate(44259.9999999))))
    Debug.Print "A9 min=" & CStr(Minute(DateAdd("n", 1, CDate(44259.9999999))))

    Debug.Print "B1 day=" & CStr(Day(DateAdd("h", 1, CDate(44259.9999999))))
    Debug.Print "B2 sec=" & CStr(Second(DateAdd("s", 1, CDate(-0.9999999))))
    Debug.Print "B3 year=" & CStr(Year(DateAdd("s", 1, CDate(-0.9999999))))
    Debug.Print "B4 day=" & CStr(Day(DateAdd("s", 1, CDate(0))))
    Debug.Print "B5 sec=" & CStr(Second(DateAdd("s", 1, CDate(0))))

    a = CStr(DateAdd("s", 1, CDate(44259.9999999)))
    Debug.Print "C1 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "C2 colon=" & CStr(InStr(a, ":") > 0)

    a = CStr(DateAdd("s", -1, CDate(44260.0000001)))
    Debug.Print "C3 slash=" & CStr(InStr(a, "/") > 0)
End Sub