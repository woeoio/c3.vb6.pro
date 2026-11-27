Option Explicit

' 11.204 (knife 221): Int / Fix keep the Date identity of a Date argument.
' Oracle .temp/k221p/P283.bas / P284.bas (VB6), t = CDate(44259.9999999):
'   VarType(Int(t)) = 7 (Date)      VarType(Int(2.5)) = 5 (Double)
'   CStr(Int(t)) = "2021/3/4"       IsDate(Int(t)) = True, IsDate(Int(2.5)) = False
'   Day(Int(CDate(-0.5))) = 29 (Int = floor) ; Day(Fix(CDate(-1.5))) = 29, Day(Int(...)) = 28
' Before this knife every Int/Fix answered Double, so CStr printed the bare serial
' "44259" and VarType answered 5.
'
' Cells are VarType numbers, integer readings and "has /" -> region-independent.
Public Sub Main()
    Dim t As Date
    Dim a As String
    On Error Resume Next

    t = CDate(44259.9999999)

    Debug.Print "F1=" & CStr(VarType(Int(t)))
    Debug.Print "F2=" & CStr(VarType(Fix(t)))
    Debug.Print "F3=" & CStr(VarType(Int(2.5)))
    Debug.Print "F4=" & CStr(VarType(Int(CDate(1.5))))

    a = CStr(Int(t))
    Debug.Print "G1 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "G2 nocolon=" & CStr(InStr(a, ":") = 0)
    Debug.Print "G3 day=" & CStr(Day(Int(CDate(-0.5))))
    Debug.Print "G4 day=" & CStr(Day(Fix(CDate(-1.5))))
    Debug.Print "G5 day=" & CStr(Day(Int(CDate(-1.5))))

    Debug.Print "H1=[" & CStr(Int(2.5)) & "]"
    Debug.Print "H2=[" & CStr(Int(-2.5)) & "]"
    Debug.Print "H3=[" & CStr(Fix(-2.5)) & "]"
    Debug.Print "H4=[" & CStr(Int(44259.75)) & "]"
    a = CStr(Int(CDate(44259.75)))
    Debug.Print "H5 slash=" & CStr(InStr(a, "/") > 0)

    Debug.Print "I1=" & CStr(IsDate(Int(t)))
    Debug.Print "I2=" & CStr(IsDate(Int(2.5)))
    Debug.Print "J1=" & CStr(Day(Int(t)))
    Debug.Print "J2=" & CStr(Hour(Int(t)))

    ' Int(t) 的 Date 身份穿得过去：算术、赋值、混合小数（真机 P285）
    a = CStr(Int(t) + 1)
    Debug.Print "K1 slash=" & CStr(InStr(a, "/") > 0)
    Debug.Print "K2 nocolon=" & CStr(InStr(a, ":") = 0)
    Debug.Print "K3=[" & CStr(Int(t) * 2) & "]"
    a = CStr(Int(t) + 0.5)
    Debug.Print "K4 colon=" & CStr(InStr(a, ":") > 0)
End Sub