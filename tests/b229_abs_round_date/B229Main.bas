Option Explicit

' 11.205 (knife 222): Abs / Round keep the Date identity too (same family as Int/Fix).
' Oracle .temp/k222p/P288.bas / P289.bas (VB6), t = CDate(44259.6242685185):
'   VarType(Abs(t)) = 7 | CStr(Abs(t)) = "2021/3/4 14:58:57"
'   VarType(Round(t,0)) = 7 | CStr(Round(t,0)) = "2021/3/5"
'   VarType(Abs(CDate(-0.5))) = 7 | CStr(...) = "12:00:00"
'   VarType(Abs(2.5)) = 5, VarType(Round(2.5)) = 5   (numeric args stay numeric)
'   CLng(t) = 44260|3, CCur(t) = 6, CInt(t) raises 6 (Overflow)  -- unchanged guards
' Still open after this knife (separate accounts): Hex(t) rounds (VB6 ACE4) while C3
' truncates (ACE3), and Round(-0.5,0) is banker's-rounded to 0 by VB6 while C3 gives -1.
'
' Cells are VarType numbers, small integers, "2.5" text and "/"-":" shapes.
Public Sub Main()
    Dim t As Date
    Dim a As String
    Dim v As Variant
    On Error Resume Next

    t = CDate(44259.6242685185)

    Debug.Print "E1=" & CStr(VarType(Abs(t)))
    Debug.Print "E2=" & CStr(VarType(Round(t, 0)))
    Debug.Print "E3=" & CStr(VarType(Round(t, 2)))
    Debug.Print "E4=" & CStr(VarType(Abs(CDate(-0.5))))
    Debug.Print "E5=" & CStr(VarType(Abs(2.5)))
    Debug.Print "E6=" & CStr(VarType(Round(2.5)))

    a = CStr(Abs(t))
    Debug.Print "F1 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "F2 colon=" & CStr(InStr(a, ":") > 0)
    a = CStr(Round(t, 0))
    Debug.Print "F3 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "F4 nocolon=" & CStr(InStr(a, ":") = 0)
    Debug.Print "F5=" & CStr(Day(Round(t, 0)))
    Debug.Print "F6=" & CStr(Hour(Abs(CDate(-0.5))))
    Debug.Print "F7=[" & CStr(Abs(2.5)) & "]"

    Debug.Print "G1=" & CStr(VarType(CLng(t)))
    Debug.Print "G2=[" & CStr(CLng(t)) & "]"
    Debug.Print "G3=" & CStr(VarType(CCur(t)))
    v = Empty: Err.Clear: v = CInt(t)
    Debug.Print "G4 err=" & CStr(Err.Number)
End Sub