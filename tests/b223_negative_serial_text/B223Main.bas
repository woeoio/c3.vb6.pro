Option Explicit

' 11.199 (knife 216): negative serials in the general-date text, and the
' Print # station's own zero-date window.
' Oracle .temp/k216p/P271.bas -> p271.txt (VB6, zh-CN):
'   CStr(CDate(-0.5))  = "12:00:00"                 (trunc part 0 -> time only)
'   CStr(CDate(-0.25)) = "6:00:00"
'   CStr(CDate(-1.5))  = "1899/12/29 12:00:00"      (trunc part -1 -> date kept)
'   Print #f, CDate(-0.5)  -> "12:00:00 "
'   Print #f, CDate(-1.5)  -> "1899/12/29 12:00:00 "
'   FormatDateTime(CDate(0.5), vbShortDate) = "1899/12/30"  (named arm NOT affected)
' Before this knife C3 used the half-open window [0,1), so every negative value
' with |serial| < 1 printed its 1899/12/30 prefix.
'
' Cells are region-independent (CI is en-US): only "has /", "has :", lengths.
Public Sub Main()
    Dim a As String
    Dim b As String
    Dim c As String
    Dim f As Integer
    Dim l1 As String
    Dim l2 As String
    On Error Resume Next

    a = CStr(CDate(-0.5))
    Debug.Print "A1 noslash=" & CStr(InStr(a, "/") = 0)
    Debug.Print "A2 colon=" & CStr(InStr(a, ":") > 0)

    a = CStr(CDate(-0.25))
    Debug.Print "A3 noslash=" & CStr(InStr(a, "/") = 0)
    Debug.Print "A4 colon=" & CStr(InStr(a, ":") > 0)

    a = CStr(CDate(-0.9999))
    Debug.Print "A5 noslash=" & CStr(InStr(a, "/") = 0)

    b = CStr(CDate(-1.5))
    Debug.Print "B1 slash=" & CStr(InStr(b, "/") > 0)
    Debug.Print "B2 colon=" & CStr(InStr(b, ":") > 0)

    c = Format(CDate(-1.5), "General Date")
    Debug.Print "C1 slash=" & CStr(InStr(c, "/") > 0)
    Debug.Print "C2 colon=" & CStr(InStr(c, ":") > 0)

    c = Format(CDate(-0.5), "General Date")
    Debug.Print "C3 noslash=" & CStr(InStr(c, "/") = 0)

    c = FormatDateTime(CDate(0.5), vbShortDate)
    Debug.Print "D1 slash=" & CStr(InStr(c, "/") > 0)
    Debug.Print "D2 nocolon=" & CStr(InStr(c, ":") = 0)

    c = FormatDateTime(CDate(1), vbShortDate)
    Debug.Print "E1 slash=" & CStr(InStr(c, "/") > 0)

    c = FormatDateTime(CDate(0.5), vbShortTime)
    Debug.Print "F1 colon=" & CStr(InStr(c, ":") > 0)

    f = FreeFile
    Open App.Path & "\b223.tmp" For Output As #f
    Print #f, CDate(-0.5)
    Print #f, CDate(-1.5)
    Close #f
    Open App.Path & "\b223.tmp" For Input As #f
    Line Input #f, l1
    Line Input #f, l2
    Close #f
    Debug.Print "G1 noslash=" & CStr(InStr(l1, "/") = 0)
    Debug.Print "G2 slash=" & CStr(InStr(l2, "/") > 0)
    Debug.Print "G3 colon=" & CStr(InStr(l1, ":") > 0)
End Sub