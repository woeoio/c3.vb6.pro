Attribute VB_Name = "NullOpsProbe"
Option Explicit
' ============================================================================
' NullOpsProbe - manual VB6 IDE probe for Null propagation cells
'
' HOW TO USE
'   1. Open VB6, create a new Standard EXE project.
'   2. Project > Add Module > Existing... and pick this file.
'   3. Press Ctrl+G (Immediate window), type:   ProbeNull      then Enter.
'   4. Copy the whole output block back (it starts with "== NullOpsProbe =="
'      and ends with "== DONE ==").
'
' FORMAT
'   <id>  <label>  | isnull=<True/False> val=<value or <Null>> tn=<TypeName> err=<Err.Number>
'   - "err" is the Err.Number *of that single statement* (Err is cleared first).
'   - "<CStr err N>" in val means the value was not printable and CStr raised N.
'
' NOTE  Every statement uses a Variant variable holding Null (nv), never the
'       literal Null, so the file compiles identically in any VB6 build.
' ============================================================================

Private Function S(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Then
        S = "<Null>"
    ElseIf IsEmpty(v) Then
        S = "<Empty>"
    Else
        S = CStr(v)
        If Err.Number <> 0 Then S = "<CStr err " & Err.Number & ">"
    End If
    On Error GoTo 0
End Function

Private Function L(ByVal s As String) As String
    If Len(s) >= 30 Then
        L = s
    Else
        L = s & Space$(30 - Len(s))
    End If
End Function

Private Sub PF(ByVal id As String, ByVal label As String, ByRef v As Variant, ByVal e As Long)
    Debug.Print id & "  " & L(label) & " | isnull=" & IsNull(v) & " val=" & S(v) _
                & " tn=" & TypeName(v) & " err=" & e
End Sub

Public Sub ProbeNull()
    Dim nv As Variant
    Dim r As Variant
    Dim s As String
    Dim d As Double
    Dim i As Long
    Dim cnt As Long
    Dim junk As Variant
    Dim ok As Boolean

    nv = Null
    On Error Resume Next

    Debug.Print "== NullOpsProbe =="
    Debug.Print "-- A. Xor --"
    Err.Clear: r = nv Xor True:  PF "X1", "Null Xor True", r, Err.Number
    Err.Clear: r = nv Xor False: PF "X2", "Null Xor False", r, Err.Number
    Err.Clear: r = nv Xor 0:     PF "X3", "Null Xor 0", r, Err.Number
    Err.Clear: r = nv Xor 5:     PF "X4", "Null Xor 5", r, Err.Number
    Err.Clear: r = nv Xor nv:    PF "X5", "Null Xor Null", r, Err.Number
    Err.Clear: r = True Xor nv:  PF "X6", "True Xor Null", r, Err.Number
    Err.Clear: r = 5 Xor nv:     PF "X7", "5 Xor Null", r, Err.Number

    Debug.Print "-- B. Eqv --"
    Err.Clear: r = nv Eqv True:  PF "E1", "Null Eqv True", r, Err.Number
    Err.Clear: r = nv Eqv False: PF "E2", "Null Eqv False", r, Err.Number
    Err.Clear: r = nv Eqv 0:     PF "E3", "Null Eqv 0", r, Err.Number
    Err.Clear: r = nv Eqv 5:     PF "E4", "Null Eqv 5", r, Err.Number
    Err.Clear: r = nv Eqv nv:    PF "E5", "Null Eqv Null", r, Err.Number
    Err.Clear: r = True Eqv nv:  PF "E6", "True Eqv Null", r, Err.Number
    Err.Clear: r = False Eqv nv: PF "E7", "False Eqv Null", r, Err.Number
    Err.Clear: r = 5 Eqv nv:     PF "E8", "5 Eqv Null", r, Err.Number

    Debug.Print "-- C. Imp --"
    Err.Clear: r = nv Imp True:  PF "I1", "Null Imp True", r, Err.Number
    Err.Clear: r = nv Imp False: PF "I2", "Null Imp False", r, Err.Number
    Err.Clear: r = nv Imp 0:     PF "I3", "Null Imp 0", r, Err.Number
    Err.Clear: r = nv Imp 5:     PF "I4", "Null Imp 5", r, Err.Number
    Err.Clear: r = nv Imp nv:    PF "I5", "Null Imp Null", r, Err.Number
    Err.Clear: r = True Imp nv:  PF "I6", "True Imp Null", r, Err.Number
    Err.Clear: r = False Imp nv: PF "I7", "False Imp Null", r, Err.Number
    Err.Clear: r = 5 Imp nv:     PF "I8", "5 Imp Null", r, Err.Number

    Debug.Print "-- D. Or / And numeric side (re-confirm) --"
    Err.Clear: r = nv Or 5:      PF "D1", "Null Or 5", r, Err.Number
    Err.Clear: r = nv And 0:     PF "D2", "Null And 0", r, Err.Number
    Err.Clear: r = nv Or 0:      PF "D3", "Null Or 0", r, Err.Number
    Err.Clear: r = nv And 5:     PF "D4", "Null And 5", r, Err.Number
    Err.Clear: r = nv Or -1:     PF "D5", "Null Or -1", r, Err.Number
    Err.Clear: r = nv And -1:    PF "D6", "Null And -1", r, Err.Number

    Debug.Print "-- E. concat chains --"
    Err.Clear: r = nv & nv:          PF "C1", "Null & Null", r, Err.Number
    Err.Clear: r = nv & nv & nv:     PF "C2", "Null & Null & Null", r, Err.Number
    Err.Clear: r = nv & nv & "b":    PF "C3", "Null & Null & ""b""", r, Err.Number
    Err.Clear: r = "a" & nv & nv:    PF "C4", """a"" & Null & Null", r, Err.Number
    Err.Clear: r = nv & "" & nv:     PF "C5", "Null & """" & Null", r, Err.Number
    Err.Clear: r = nv & 1:           PF "C6", "Null & 1", r, Err.Number
    Err.Clear: r = (nv & nv) & "b":  PF "C7", "(Null & Null) & ""b""", r, Err.Number
    Err.Clear: r = nv & (nv & "b"):  PF "C8", "Null & (Null & ""b"")", r, Err.Number

    Debug.Print "-- F. String target --"
    s = "keep"
    Err.Clear: s = nv:        Debug.Print "S1  s = Null                    | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep"
    Err.Clear: s = nv & "x":  Debug.Print "S2  s = Null & ""x""              | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep"
    Err.Clear: s = nv & nv:   Debug.Print "S3  s = Null & Null             | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep"
    Err.Clear: s = CStr(nv):  Debug.Print "S4  s = CStr(Null)              | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep"
    Err.Clear: s = Hex(nv):   Debug.Print "S5  s = Hex(Null)               | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number

    Debug.Print "-- G. Join / UBound / LBound / Split / Replace / Len --"
    Err.Clear: r = Join(nv, ","):      PF "J1", "Join(Null, "","")", r, Err.Number
    Err.Clear: r = UBound(nv):         PF "J2", "UBound(Null)", r, Err.Number
    Err.Clear: r = LBound(nv):         PF "J3", "LBound(Null)", r, Err.Number
    Err.Clear: r = Split(nv, ","):     PF "J4", "Split(Null, "","")", r, Err.Number
    Err.Clear: r = Replace(nv, "a", "b"): PF "J5", "Replace(Null,""a"",""b"")", r, Err.Number
    Err.Clear: r = Len(nv):            PF "J6", "Len(Null)", r, Err.Number
    Err.Clear: r = Trim(nv):           PF "J7", "Trim(Null)", r, Err.Number
    Err.Clear: r = Val(nv):            PF "J8", "Val(Null)", r, Err.Number

    Debug.Print "-- H. numeric target --"
    d = 1.5
    Err.Clear: d = nv:        Debug.Print "N1  d = Null                | d=" & d & " err=" & Err.Number
    d = 1.5
    Err.Clear: d = Abs(nv):   Debug.Print "N2  d = Abs(Null)           | d=" & d & " err=" & Err.Number
    d = 1.5
    Err.Clear: d = nv * 1:    Debug.Print "N3  d = Null * 1            | d=" & d & " err=" & Err.Number
    d = 1.5
    Err.Clear: d = nv + 1:    Debug.Print "N4  d = Null + 1            | d=" & d & " err=" & Err.Number
    i = 7
    Err.Clear: i = nv:        Debug.Print "N5  i = Null                | i=" & i & " err=" & Err.Number
    i = 7
    Err.Clear: i = CInt(nv):  Debug.Print "N6  i = CInt(Null)          | i=" & i & " err=" & Err.Number
    i = 7
    Err.Clear: i = Len(nv):   Debug.Print "N7  i = Len(Null)           | i=" & i & " err=" & Err.Number

    Debug.Print "-- I. condition position --"
    ok = False
    Err.Clear
    If nv Then ok = True
    Debug.Print "F1  If Null Then            | taken=" & ok & " err=" & Err.Number
    ok = False
    Err.Clear
    If nv Or True Then ok = True
    Debug.Print "F2  If Null Or True Then    | taken=" & ok & " err=" & Err.Number
    ok = False
    Err.Clear
    If nv = "" Then ok = True
    Debug.Print "F3  If Null = """" Then       | taken=" & ok & " err=" & Err.Number
    ok = False
    cnt = 0
    Err.Clear
    Do While nv
        cnt = cnt + 1
        If cnt > 1 Then Exit Do
    Loop
    Debug.Print "F4  Do While Null           | iter=" & cnt & " err=" & Err.Number
    ok = False
    Err.Clear
    ok = (nv = "")
    Debug.Print "F5  ok = (Null = """")       | ok=" & ok & " err=" & Err.Number

    Debug.Print "-- J. IIf / Not --"
    Err.Clear: r = IIf(True, nv, 1):  PF "Q1", "IIf(True, Null, 1)", r, Err.Number
    Err.Clear: r = IIf(False, 1, nv): PF "Q2", "IIf(False, 1, Null)", r, Err.Number
    Err.Clear: r = Not nv:            PF "Q3", "Not Null", r, Err.Number
    Err.Clear: r = nv & "":           PF "Q4", "Null & """"", r, Err.Number
    Err.Clear: junk = nv:             PF "Q5", "junk = Null (assign)", junk, Err.Number

    Debug.Print "== DONE =="
End Sub
