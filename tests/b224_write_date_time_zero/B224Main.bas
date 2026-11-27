Option Explicit

' 11.200 (knife 217): the Write # text of a date. Oracle .temp/k217p/P273.bas ->
' p273w.txt (VB6):  the time part is omitted entirely when it is exactly 00:00:00
'   Write #f, CDate(0)         -> #1899-12-30#
'   Write #f, CDate(0.5)       -> #1899-12-30 12:00:00#
'   Write #f, CDate(1)         -> #1899-12-31#
'   Write #f, CDate(-1.5)      -> #1899-12-29 12:00:00#
'   Write #f, CDate(0.0000001) -> #1899-12-30#      (time rounds to zero too)
' Before this knife C3 always printed " HH:MM:SS", so every whole date carried
' a spurious " 00:00:00".  Write # is ISO-shaped, so these cells are
' region-independent; only "has :", "len", and the surrounding # are pinned.
Public Sub Main()
    Dim f As Integer
    Dim l0 As String
    Dim l1 As String
    Dim l2 As String
    Dim l3 As String
    Dim l4 As String
    Dim a As String
    On Error Resume Next

    f = FreeFile
    Open App.Path & "\b224.tmp" For Output As #f
    Write #f, CDate(0)
    Write #f, CDate(0.5)
    Write #f, CDate(1)
    Write #f, CDate(-1.5)
    Write #f, CDate(0.0000001)
    Close #f

    f = FreeFile
    Open App.Path & "\b224.tmp" For Input As #f
    Line Input #f, l0
    Line Input #f, l1
    Line Input #f, l2
    Line Input #f, l3
    Line Input #f, l4
    Close #f

    Debug.Print "A1 nocolon=" & CStr(InStr(l0, ":") = 0)
    Debug.Print "A2 len=" & CStr(Len(l0) = 12)
    Debug.Print "A3 hash=" & CStr(Left$(l0, 1) = "#" And Right$(l0, 1) = "#")

    Debug.Print "B1 colon=" & CStr(InStr(l1, ":") > 0)
    Debug.Print "B2 len=" & CStr(Len(l1) = 21)
    Debug.Print "B3 hash=" & CStr(Left$(l1, 1) = "#" And Right$(l1, 1) = "#")

    Debug.Print "C1 nocolon=" & CStr(InStr(l2, ":") = 0)
    Debug.Print "C2 len=" & CStr(Len(l2) = 12)

    Debug.Print "D1 colon=" & CStr(InStr(l3, ":") > 0)
    Debug.Print "D2 hash=" & CStr(Left$(l3, 1) = "#" And Right$(l3, 1) = "#")

    Debug.Print "E1 nocolon=" & CStr(InStr(l4, ":") = 0)
    Debug.Print "E2 len=" & CStr(Len(l4) = 12)

    ' guard cells: the CStr side of the same near-zero serials (knife 216 grid)
    a = CStr(CDate(-0.0000001))
    Debug.Print "F1 noslash=" & CStr(InStr(a, "/") = 0)
    Debug.Print "F2 colon=" & CStr(InStr(a, ":") > 0)
End Sub