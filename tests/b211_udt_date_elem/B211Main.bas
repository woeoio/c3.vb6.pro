Option Explicit

' §11.187 (knife 204): a Date array that lives INSIDE a user-defined type keeps
' the same String -> Date conversion, 1-D and multi-dim.
' Oracle: VB6Mini /make .temp/k204p/P244.bas -> p244.txt, eight lines:
'   U1..U4 = 44259 / 44259 / 44687 / 44259 err=0 (member 1-D literal, member 1-D
'            from a String variable, member 2-D literal, member 2-D from a variable)
'   U5..U6 = keeps the previous value (45000 / 0) and raises 13
'   U7..U8 = vt=7 tn=Date
' Before this knife all six assignments emitted a raw BSTR into a double slot
' (cl C2440, no exe): the Date leg only recognised a bare array name as callee,
' and a UDT member array element comes in as MemberAccess (`u.dt`).
Public Type TDates
    dt() As Date
    dt2() As Date
    n As Long
End Type

Public Sub Main()
    Dim u As TDates
    Dim s As String

    On Error Resume Next

    ReDim u.dt(0 To 1)
    ReDim u.dt2(0 To 1, 0 To 1)

    Err.Clear
    u.dt(0) = "2021-03-04"
    Debug.Print "U1 d=" & CStr(CDbl(u.dt(0))) & " err=" & CStr(Err.Number)

    s = "2021-03-04"
    Err.Clear
    u.dt(1) = s
    Debug.Print "U2 d=" & CStr(CDbl(u.dt(1))) & " err=" & CStr(Err.Number)

    Err.Clear
    u.dt2(0, 1) = "2022-05-06"
    Debug.Print "U3 d=" & CStr(CDbl(u.dt2(0, 1))) & " err=" & CStr(Err.Number)

    Err.Clear
    u.dt2(1, 0) = s
    Debug.Print "U4 d=" & CStr(CDbl(u.dt2(1, 0))) & " err=" & CStr(Err.Number)

    u.dt(0) = 45000
    Err.Clear
    u.dt(0) = "abc"
    Debug.Print "U5 d=" & CStr(CDbl(u.dt(0))) & " err=" & CStr(Err.Number)

    Err.Clear
    u.dt2(0, 0) = "abc"
    Debug.Print "U6 d=" & CStr(CDbl(u.dt2(0, 0))) & " err=" & CStr(Err.Number)

    Debug.Print "U7 vt=" & CStr(VarType(u.dt(0))) & " tn=" & TypeName(u.dt(0))
    Debug.Print "U8 vt=" & CStr(VarType(u.dt2(0, 1))) & " tn=" & TypeName(u.dt2(0, 1))
End Sub