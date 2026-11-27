Option Explicit

' 11.198 (knife 215): the GENERAL DATE text shape of date/time values.
' VB6 rule (oracle .temp/k215p/P269.bas -> p269.txt):
'   serial in [0,1)  (VB6's zero date 1899-12-30)  -> TIME only   "0:00:00"
'   serial 1 or -1   (no time part)                -> DATE only   "1899/12/31"
'   both parts                                     -> "date time"
' Before this knife C3 always emitted "short date + space + long time", so
' CStr(TimeSerial(12,30,0)) printed "1899/12/30 12:30:00" instead of "12:30:00"
' (oracle P268 D3/D5).
'
' Cells are REGION-INDEPENDENT on purpose (CI is en-US, the oracle is zh-CN):
' only "has /", "has :", IsDate, VarType and lengths. Never pin the text.
Public Sub Main()
    Dim a As String
    Dim b As String
    Dim c As String
    Dim d As String
    Dim e As String
    Dim g As String
    Dim h As String
    Dim k As String
    On Error Resume Next

    a = CStr(TimeSerial(12, 30, 0))
    Debug.Print "A1 colon=" & CStr(InStr(a, ":") > 0)
    Debug.Print "A2 noslash=" & CStr(InStr(a, "/") = 0)
    Debug.Print "A3 isdate=" & CStr(IsDate(a))
    Debug.Print "A4 len=" & CStr(Len(a) >= 7)

    b = CStr(DateSerial(2021, 3, 4))
    Debug.Print "B1 slash=" & CStr(InStr(b, "/") > 0)
    Debug.Print "B2 nocolon=" & CStr(InStr(b, ":") = 0)

    c = CStr(CDate(0.5))
    Debug.Print "C1 noslash=" & CStr(InStr(c, "/") = 0)
    Debug.Print "C2 colon=" & CStr(InStr(c, ":") > 0)
    Debug.Print "C3 len=" & CStr(Len(c) >= 7)

    d = CStr(CDate(1))
    Debug.Print "D1 slash=" & CStr(InStr(d, "/") > 0)
    Debug.Print "D2 nocolon=" & CStr(InStr(d, ":") = 0)

    e = Format(CDate(0.5), "General Date")
    Debug.Print "E1 noslash=" & CStr(InStr(e, "/") = 0)
    Debug.Print "E2 colon=" & CStr(InStr(e, ":") > 0)

    g = CStr(DateAdd("d", 1, CDate("2021-03-04")))
    Debug.Print "G1 slash=" & CStr(InStr(g, "/") > 0)
    Debug.Print "G2 nocolon=" & CStr(InStr(g, ":") = 0)

    h = CStr(TimeValue("12:30:00"))
    Debug.Print "H1 noslash=" & CStr(InStr(h, "/") = 0)
    Debug.Print "H2 colon=" & CStr(InStr(h, ":") > 0)

    k = CStr(DateValue("2021-03-04"))
    Debug.Print "K1 slash=" & CStr(InStr(k, "/") > 0)
    Debug.Print "K2 nocolon=" & CStr(InStr(k, ":") = 0)

    Debug.Print "V1 vt=" & CStr(VarType(TimeSerial(12, 30, 0)) = 7)
    Debug.Print "V2 vt=" & CStr(VarType(DateValue("2021-03-04")) = 7)
End Sub