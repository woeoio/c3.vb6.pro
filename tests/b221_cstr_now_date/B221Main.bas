Option Explicit

' 11.197 (knife 214): CStr(<zero-arg date builtin>) must produce locale date/
' time TEXT, not the raw serial number. Oracle: VB6Mini /make,
' .temp/k214p/P264.bas -> p264.txt  (CStr(Now) = "2026/11/27 10:22:18",
' CStr(Date) = "2026/11/27") and P267.bas -> p267.txt (the twelve cells below).
'
' Before this knife CStr(Now) emitted vb6_CStr(vb6_VariantFromValue(vb6_Now()))
' -- a VT_R8 boxed serial -- so the program printed "46353.4325925926".
'
' The cells are deliberately REGION-INDEPENDENT (the CI runs in en-US, the
' oracle ran in zh-CN): only "does it contain / or :", "does CDate accept it",
' "does it round-trip to the same instant", and lengths. Never pin the text.
Public Sub Main()
    Dim s As String
    Dim d As String
    Dim t As String
    On Error Resume Next
    s = CStr(Now)
    Debug.Print "W1 len=" & CStr(Len(s) > 10)
    Debug.Print "W2 slash=" & CStr(InStr(s, "/") > 0)
    Debug.Print "W3 colon=" & CStr(InStr(s, ":") > 0)
    Debug.Print "W4 isdate=" & CStr(IsDate(s))
    Debug.Print "W5 str=" & CStr(VarType(s) = vbString)
    Debug.Print "W6 round=" & CStr(Abs(CDate(s) - Now) < 0.01)
    d = CStr(Date)
    Debug.Print "X1 isdate=" & CStr(IsDate(d))
    Debug.Print "X2 len=" & CStr(Len(d) >= 6)
    Debug.Print "X3 nocolon=" & CStr(InStr(d, ":") = 0)
    t = CStr(Time)
    Debug.Print "Y1 colon=" & CStr(InStr(t, ":") > 0)
    Debug.Print "Y2 len=" & CStr(Len(t) >= 4)
    Debug.Print "Y3 diff=" & CStr(s <> d)
End Sub