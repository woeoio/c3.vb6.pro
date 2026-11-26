Option Explicit

' §11.184 (knife 201): a Date array element keeps its identity whatever its rank.
' Oracle: VB6Mini /make
'   .temp/k201p/P242.bas -> p242.txt  (S1..S7: the assignment station)
'   .temp/k201p/P243.bas -> p243.txt  (T1..T4: Variant boxing of an element)
' Readings: 2-D/3-D static elements and a ReDim'd 2-D element all convert
' "2021-03-04"/"2022-05-06"/"2023-07-08" (44259/44687/45115, err=0); bad text
' raises 13 and keeps the old value (45000 / 0). VarType answers 7 ("Date")
' for scalar, 1-D dynamic, 1-D static and 2-D static alike.
' Before this knife the multi-dim forms emitted a raw BSTR into a double slot
' (cl C2440, no exe) and boxed the element as VT_R8 (vt=5, tn=Double).
Public Sub Main()
    Dim d As Date
    Dim s As String
    Dim d2(0 To 1, 0 To 1) As Date
    Dim d3(0 To 1, 0 To 1, 0 To 1) As Date
    Dim q2() As Date
    Dim a1(0 To 1) As Date
    Dim q1() As Date

    On Error Resume Next

    Err.Clear
    d2(0, 1) = "2021-03-04"
    Debug.Print "S1 d=" & CStr(CDbl(d2(0, 1))) & " err=" & CStr(Err.Number)

    s = "2021-03-04"
    Err.Clear
    d2(1, 0) = s
    Debug.Print "S2 d=" & CStr(CDbl(d2(1, 0))) & " err=" & CStr(Err.Number)

    ReDim q2(0 To 1, 0 To 1)
    Err.Clear
    q2(1, 1) = "2022-05-06"
    Debug.Print "S3 d=" & CStr(CDbl(q2(1, 1))) & " err=" & CStr(Err.Number)

    Err.Clear
    d3(1, 1, 1) = "2023-07-08"
    Debug.Print "S4 d=" & CStr(CDbl(d3(1, 1, 1))) & " err=" & CStr(Err.Number)

    d2(0, 0) = 45000
    Err.Clear
    d2(0, 0) = "abc"
    Debug.Print "S5 d=" & CStr(CDbl(d2(0, 0))) & " err=" & CStr(Err.Number)

    Err.Clear
    q2(0, 0) = "abc"
    Debug.Print "S6 d=" & CStr(CDbl(q2(0, 0))) & " err=" & CStr(Err.Number)

    d = "2021-03-04"
    ReDim q1(0 To 1)
    q1(0) = d
    a1(1) = d

    Debug.Print "T1 scalar vt=" & CStr(VarType(d)) & " tn=" & TypeName(d)
    Debug.Print "T2 1d-dyn vt=" & CStr(VarType(q1(0))) & " tn=" & TypeName(q1(0))
    Debug.Print "T3 1d-sta vt=" & CStr(VarType(a1(1))) & " tn=" & TypeName(a1(1))
    Debug.Print "T4 2d-sta vt=" & CStr(VarType(d2(0, 1))) & " tn=" & TypeName(d2(0, 1))
End Sub
