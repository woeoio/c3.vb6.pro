Option Explicit

' 11.201 (knife 218): the zero-day test must follow the time rounding.
' Oracle .temp/k218p/P275.bas -> p275p/p275w/p275c (VB6):
'   CStr(CDate(0.9999999))  = "1899/12/31"      time rounds up to 24:00 -> carry
'   CStr(CDate(-0.9999999)) = "1899/12/31"      carry direction is always +1
'   Print #f, CDate(0.9999999)  = "1899/12/31 "
'   Print #f, CDate(0)          = "0:00:00 "
'   Write #f, CDate(-0.0000001) = "#1899-12-30#"
' Before this knife the (0,1) window was decided on the raw serial, so the two
' 0.9999999 shapes printed "0:00:00" / "23:59:59 " instead of the carried date.
'
' Cells are region-independent (CI is en-US): only "has /", "has :", lengths.
Public Sub Main()
    Dim f As Integer
    Dim l1 As String
    Dim l2 As String
    Dim l3 As String
    Dim a As String
    On Error Resume Next

    a = CStr(CDate(0.9999999))
    Debug.Print "C1 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "C2 nocolon=" & CStr(InStr(a, ":") = 0)

    a = CStr(CDate(-0.9999999))
    Debug.Print "C3 slash=" & CStr(InStr(a, "/") > 0)

    a = CStr(CDate(0.0000001))
    Debug.Print "C4 noslash=" & CStr(InStr(a, "/") = 0)
    Debug.Print "C5 colon=" & CStr(InStr(a, ":") > 0)

    a = CStr(CDate(0))
    Debug.Print "C6 noslash=" & CStr(InStr(a, "/") = 0)

    a = Format(CDate(0.9999999), "General Date")
    Debug.Print "G1 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "G2 nocolon=" & CStr(InStr(a, ":") = 0)

    a = FormatDateTime(CDate(0.9999999), vbShortDate)
    Debug.Print "G3 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "G4 nocolon=" & CStr(InStr(a, ":") = 0)

    f = FreeFile
    Open App.Path & "\b225.tmp" For Output As #f
    Print #f, CDate(0)
    Print #f, CDate(0.9999999)
    Print #f, CDate(-0.9999999)
    Close #f
    f = FreeFile
    Open App.Path & "\b225.tmp" For Input As #f
    Line Input #f, l1
    Line Input #f, l2
    Line Input #f, l3
    Close #f
    Debug.Print "P1 noslash=" & CStr(InStr(l1, "/") = 0)
    Debug.Print "P2 colon=" & CStr(InStr(l1, ":") > 0)
    Debug.Print "P3 slash=" & CStr(InStr(l2, "/") > 0)
    Debug.Print "P4 nocolon=" & CStr(InStr(l2, ":") = 0)
    Debug.Print "P5 slash=" & CStr(InStr(l3, "/") > 0)

    f = FreeFile
    Open App.Path & "\b225w.tmp" For Output As #f
    Write #f, CDate(0.9999999)
    Write #f, CDate(-0.0000001)
    Write #f, CDate(0.4999999)
    Close #f
    f = FreeFile
    Open App.Path & "\b225w.tmp" For Input As #f
    Line Input #f, l1
    Line Input #f, l2
    Line Input #f, l3
    Close #f
    Debug.Print "W1 nocolon=" & CStr(InStr(l1, ":") = 0)
    Debug.Print "W2 len=" & CStr(Len(l1) = 12)
    Debug.Print "W3 nocolon=" & CStr(InStr(l2, ":") = 0)
    Debug.Print "W4 colon=" & CStr(InStr(l3, ":") > 0)
    Debug.Print "W5 len=" & CStr(Len(l3) = 21)
End Sub