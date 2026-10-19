' Fix 432 (ledger 305 #432, knife 103): FormatDateTime's two formals now ask the ONE authority.
'
' Before this knife the arg-tier question answered "" for both positions (neither kNoTableFormalKinds176
' nor the runtime C-signature table had a vb6_FormatDateTime row), so a whole VARIANT struct was handed
' straight to `double dateSerial` / `int32_t namedFormat` => cl C2440 and the project did not build at
' all. Eight shapes were hard errors; four (As Date local, date literal, DateSerial, As Date + named
' format 1) compiled because the RTL return of those is already a plain double.
'
' Caliber: same as b431/b432 - every needle is relational or locale-free. This box is zh-CN and the CI
' runner is en-US, so a spelled date is never compared to a literal; it is compared to the SAME named
' format tier asked of a value that already worked (the date literal / As Date local).
'
' The String-shaped half is deliberately NOT in here: `FormatDateTime(t, 0)` with t As String is still a
' hard error after this row (a BSTR is not a slot, so the station leaves it alone) and the engine parses
' such a text as a DATE ("46027" -> 2026/1/5, "abc" -> error 13). That half is ledger #433.
Option Explicit

Private Function GiveV() As Variant
    GiveV = #1/5/2026#
End Function

Sub Main()
    Dim i As Long
    Dim v As Variant
    Dim n As Variant
    Dim d As Date
    Dim s As String
    Dim b As String
    Dim allSame As Boolean
    Dim c As B433Holder
    Dim o As Object

    d = #1/5/2026#
    v = d
    b = FormatDateTime(d, 0)

    Debug.Print "B433-01=" & CStr(FormatDateTime(v, 0) = b) & "/" & CStr(Len(b) > 0)
    Debug.Print "B433-02=" & CStr(VarType(FormatDateTime(v, 0)) = 8) & "/" & _
        CStr(TypeName(FormatDateTime(v, 0)) = "String")

    allSame = True
    For i = 0 To 4
        If FormatDateTime(v, i) <> FormatDateTime(d, i) Then allSame = False
    Next i
    Debug.Print "B433-03=" & CStr(allSame)

    ' the four shapes that compiled before this knife must keep answering the same thing
    Debug.Print "B433-04=" & CStr(FormatDateTime(#1/5/2026#, 0) = b) & "/" & _
        CStr(FormatDateTime(CDate(46027), 0) = b) & "/" & _
        CStr(FormatDateTime(DateSerial(2026, 1, 5), 0) = b) & "/" & _
        CStr(FormatDateTime(d, 2) = FormatDateTime(v, 2))

    ' slot in the SECOND formal too (that position was a bare struct into int32_t)
    n = 2
    s = FormatDateTime(v, n)
    Debug.Print "B433-05=" & CStr(s = FormatDateTime(d, 2)) & "/" & CStr(Len(s) > 0)

    ' numeric-carrying slot: the engine folds it as a serial number, same as the date
    v = CDbl(46027)
    Debug.Print "B433-06=" & CStr(FormatDateTime(v, 0) = b)

    ' omitted second formal (the "0" default) with a slot first argument
    v = d
    s = FormatDateTime(v)
    Debug.Print "B433-07=" & CStr(s = b) & "/" & CStr(s = FormatDateTime(v, 0))

    ' module Function that returns Variant
    Debug.Print "B433-08=" & CStr(FormatDateTime(GiveV(), 0) = b)

    ' public Variant FIELD of an in-project class
    Set c = New B433Holder
    c.Modified = d
    Debug.Print "B433-09=" & CStr(FormatDateTime(c.Modified, 1) = FormatDateTime(d, 1))

    ' late-bound COM member value - the rs("X").value family in the corpus
    On Error Resume Next
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", d
    Debug.Print "B433-10=" & CStr(Err.Number = 0) & "/" & _
        CStr(FormatDateTime(o.Item("k"), 0) = b)
    On Error GoTo 0

    ' a slot carrying a Date with a time part must still reach the same outlet as the literal
    v = DateSerial(2026, 1, 5) + TimeSerial(14, 30, 45)
    Debug.Print "B433-11=" & CStr(VarType(FormatDateTime(v, 2)) = 8) & "/" & _
        CStr(FormatDateTime(v, 3) = FormatDateTime(CDate(CDbl(v)), 3)) & "/" & _
        CStr(FormatDateTime(v, 0) <> b)

    Debug.Print "B433-DONE"
End Sub
