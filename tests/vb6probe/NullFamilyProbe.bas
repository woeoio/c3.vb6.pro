Attribute VB_Name = "NullFamilyProbe"
Option Explicit
' ============================================================================
' NullFamilyProbe - second manual probe: adjacency readings needed before
' tightening Len/Trim/Val/UBound + the String/numeric/Boolean target stations.
'
' HOW TO USE
'   1. VB6 > new Standard EXE > Project > Add Module > Existing... > this file
'   2. Ctrl+G, type:  ProbeNullFamily   then Enter
'   3. Paste the whole block back ("== NullFamilyProbe ==" .. "== DONE ==")
'
' SENTINEL DISCIPLINE (this is what batch 1 lacked)
'   Before every cell the result holder is set to the string "<<stale>>".
'   If a statement raises and does not assign, the printout shows val=<<stale>>
'   instead of a leftover value from the previous cell.
'
' FORMAT
'   <id>  <label> | isnull= val= tn= err=
' ============================================================================

' 第七十一刀 后记 (W 段): 本刀的判据只认「整只槽 = 简单变量 / 成员」。下面几形读数未取 ——
' 先量再决定要不要放宽（宁可留缺口，也不拿没量过的格子去改口径）。
Private Type TWV
    V As Variant
End Type

Private Function FNullV() As Variant
    FNullV = Null
End Function

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

' r/sentinel: returns the Variant itself (already reset by the caller)
Private Sub PF(ByVal id As String, ByVal label As String, ByRef v As Variant, ByVal e As Long)
    Debug.Print id & "  " & L(label) & " | isnull=" & IsNull(v) & " val=" & S(v) _
                & " tn=" & TypeName(v) & " err=" & e
End Sub

Public Sub ProbeNullFamily()
    Dim nv As Variant, ev As Variant
    Dim n5 As Variant
    Dim i2 As Long, l2 As Long
    Dim r As Variant
    Dim s As String
    Dim d As Double
    Dim i As Long
    Dim b As Boolean
    Dim a(1 To 3) As Long
    Dim cnt As Long
    Dim tw As TWV
    Dim av(1 To 1) As Variant
    Dim col As Collection

    nv = Null
    ev = Empty
    n5 = 5
    i2 = 5
    l2 = 12345
    a(1) = 10: a(2) = 20: a(3) = 30
    On Error Resume Next

    Debug.Print "== NullFamilyProbe =="

    Debug.Print "-- A. Len --"
    r = "<<stale>>": Err.Clear: r = Len(nv):   PF "A1", "Len(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Len(ev):   PF "A2", "Len(Empty)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Len(i2):   PF "A3", "Len(5) (via Long)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Len(""):   PF "A4", "Len("""")", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Len(l2):   PF "A5", "Len(12345) (via Long)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Len("abc"): PF "A6", "Len(""abc"")", r, Err.Number

    Debug.Print "-- B. Trim / LTrim / RTrim / UCase / LCase --"
    r = "<<stale>>": Err.Clear: r = Trim(nv):  PF "B1", "Trim(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Trim(ev):  PF "B2", "Trim(Empty)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Trim("  a  "): PF "B3", "Trim(""  a  "")", r, Err.Number
    r = "<<stale>>": Err.Clear: r = LTrim(nv): PF "B4", "LTrim(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = RTrim(nv): PF "B5", "RTrim(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = UCase(nv): PF "B6", "UCase(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = LCase(nv): PF "B7", "LCase(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = UCase(ev): PF "B8", "UCase(Empty)", r, Err.Number

    Debug.Print "-- C. Val / Asc / Mid --"
    r = "<<stale>>": Err.Clear: r = Val(nv):   PF "C1", "Val(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Val(ev):   PF "C2", "Val(Empty)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Val("12.5"): PF "C3", "Val(""12.5"")", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Asc(nv):   PF "C4", "Asc(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Asc(ev):   PF "C5", "Asc(Empty)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Mid(nv, 1): PF "C6", "Mid(Null, 1)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Space$(1) & nv: PF "C7", "Space$(1) & Null", r, Err.Number

    Debug.Print "-- D. UBound / LBound --"
    r = "<<stale>>": Err.Clear: r = UBound(nv):  PF "D1", "UBound(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = LBound(nv):  PF "D2", "LBound(Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = UBound(ev):  PF "D3", "UBound(Empty)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = LBound(ev):  PF "D4", "LBound(Empty)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = UBound(n5):  PF "D5", "UBound(5) (via Variant)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = UBound(a):   PF "D6", "UBound(a(1 To 3))", r, Err.Number
    r = "<<stale>>": Err.Clear: r = LBound(a):   PF "D7", "LBound(a(1 To 3))", r, Err.Number

    Debug.Print "-- E. String target --"
    s = "keep": Err.Clear: s = ev:            Debug.Print "E1  s = Empty               | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = 0:             Debug.Print "E2  s = 0                   | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = CStr(ev):      Debug.Print "E3  s = CStr(Empty)         | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = nv & "":       Debug.Print "E4  s = Null & """"           | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = "" & nv:       Debug.Print "E5  s = """" & Null           | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = Trim(nv):      Debug.Print "E6  s = Trim(Null)          | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = Len(nv):       Debug.Print "E7  s = Len(Null)           | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = nv:            Debug.Print "E8  s = Null                | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = CStr(nv):      Debug.Print "E9  s = CStr(Null)          | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number
    s = "keep": Err.Clear: s = Hex(nv):       Debug.Print "E10 s = Hex(Null)           | len=" & Len(s) & " val=[" & s & "] err=" & Err.Number

    Debug.Print "-- W. shapes the 71st cut did not cover --"
    tw.V = Null
    r = "<<stale>>": Err.Clear: r = Len(tw.V):          PF "W1", "Len(udt.Variant member)", r, Err.Number
    With tw
        r = "<<stale>>": Err.Clear: r = Len(.V):        PF "W2", "Len(.V) inside With", r, Err.Number
        r = "<<stale>>": Err.Clear: r = Trim(.V):       PF "W3", "Trim(.V) inside With", r, Err.Number
    End With
    av(1) = Null
    r = "<<stale>>": Err.Clear: r = Len(av(1)):         PF "W4", "Len(Variant array element)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = Len(FNullV()):      PF "W5", "Len(function returning Null)", r, Err.Number
    Set col = New Collection
    col.Add Null
    r = "<<stale>>": Err.Clear: r = Len(col.Item(1)):   PF "W6", "Len(Collection.Item(1))", r, Err.Number

    Debug.Print "-- F. numeric / boolean target --"
    d = 1.5: Err.Clear: d = ev:          Debug.Print "F1  d = Empty               | d=" & d & " err=" & Err.Number
    d = 1.5: Err.Clear: d = nv & "":     Debug.Print "F2  d = Null & """"           | d=" & d & " err=" & Err.Number
    d = 1.5: Err.Clear: d = CStr(nv):    Debug.Print "F3  d = CStr(Null)          | d=" & d & " err=" & Err.Number
    d = 1.5: Err.Clear: d = 0:           Debug.Print "F4  d = 0                   | d=" & d & " err=" & Err.Number
    i = 7:   Err.Clear: i = ev:          Debug.Print "F5  i = Empty               | i=" & i & " err=" & Err.Number
    b = True: Err.Clear: b = nv:         Debug.Print "F6  b = Null                | b=" & b & " err=" & Err.Number
    b = True: Err.Clear: b = (nv = ""):  Debug.Print "F7  b = (Null = """"" & ")       | b=" & b & " err=" & Err.Number

    Debug.Print "-- G. concat chains --"
    r = "<<stale>>": Err.Clear: r = nv & nv & nv & nv:   PF "G1", "Null & Null & Null & Null", r, Err.Number
    r = "<<stale>>": Err.Clear: r = nv & nv & 1:         PF "G2", "Null & Null & 1", r, Err.Number
    r = "<<stale>>": Err.Clear: r = (nv & nv) & (nv & nv): PF "G3", "(Null & Null) & (Null & Null)", r, Err.Number
    r = "<<stale>>": Err.Clear: r = nv & nv & "b" & nv:  PF "G4", "Null & Null & ""b"" & Null", r, Err.Number
    r = "<<stale>>": Err.Clear: r = "" & nv & "" & nv:   PF "G5", """"" & Null & """" & Null", r, Err.Number
    r = "<<stale>>": Err.Clear: r = nv & ev:             PF "G6", "Null & Empty", r, Err.Number

    Debug.Print "-- H. types of the results above --"
    r = "<<stale>>": Err.Clear: r = Len(nv):  Debug.Print "H1  TypeName(Len(Null))   = " & TypeName(r) & "  VarType=" & VarType(r)
    r = "<<stale>>": Err.Clear: r = Len(ev):  Debug.Print "H2  TypeName(Len(Empty))  = " & TypeName(r) & "  VarType=" & VarType(r)
    r = "<<stale>>": Err.Clear: r = Trim(nv): Debug.Print "H3  TypeName(Trim(Null))  = " & TypeName(r) & "  VarType=" & VarType(r)
    r = "<<stale>>": Err.Clear: r = Val(nv):  Debug.Print "H4  TypeName(Val(Null))   = " & TypeName(r) & "  VarType=" & VarType(r)

    Debug.Print "== DONE =="
End Sub
