Attribute VB_Name = "Main"
Option Explicit

' b216_slot (zhang #305 SB216, knife 119): "a project-class instance going into a
' late-bound slot" is ONE question and must have ONE answer.  Nine arms on one stage:
'   A1..A4  Object slot <- in-table class, four RHS shapes (inline New / local class
'           var / module-level As New / function returning the class)  => tn=B216Impl, late=42
'   V1,V2   Variant slot <- same class                          => vt=9, tn=B216Impl, late=42
'           (SB216 ge-2: before this knife that station never asked anybody, the raw
'            instance pointer was promoted by _Generic's pointer arm -> vt looked like
'            VT_DISPATCH while pdispVal was a bare vb6_cls_B216Impl*)
'   N1,N2   Object slot <- the class WITHOUT the MultiUse block (not in the runtime
'           coclass table), both arms                   => tn=Nothing, isNothing=True
'           (SB216 ge-1: before this knife the inline-New arm answered Nothing while the
'            variable arm kept the bare pointer = one question, two answers)
'   N3      Variant slot <- that same out-of-table class        => tn=Nothing
' Witnesses: B216Impl is in the table, B216Plain is not - the two differ ONLY by the
' BEGIN / MultiUse block, so N1..N3 vs A/V arms is the gate's own reading, not a story.

Private mA As New B216Impl
Private mP As New B216Plain

Function MakeImpl() As B216Impl
    Dim t As B216Impl
    Set t = New B216Impl
    Set MakeImpl = t
End Function

Sub Main()
    Dim a As B216Impl
    Set a = New B216Impl

    Dim vo1 As Object
    Set vo1 = New B216Impl
    Debug.Print "A1-inlineNew-tn=" & TypeName(vo1) & " late=" & vo1.Twice(21)

    Dim vo2 As Object
    Set vo2 = a
    Debug.Print "A2-classVar-tn=" & TypeName(vo2) & " late=" & vo2.Twice(21)

    Dim vo3 As Object
    Set vo3 = mA
    Debug.Print "A3-moduleNew-tn=" & TypeName(vo3) & " late=" & vo3.Twice(21)

    Dim vo4 As Object
    Set vo4 = MakeImpl()
    Debug.Print "A4-fnResult-tn=" & TypeName(vo4) & " late=" & vo4.Twice(21)

    Dim v1 As Variant
    Set v1 = New B216Impl
    Debug.Print "V1-variant-inlineNew-vt=" & VarType(v1) & " tn=" & TypeName(v1) & " late=" & v1.Twice(21)

    Dim v2 As Variant
    Set v2 = a
    Debug.Print "V2-variant-classVar-vt=" & VarType(v2) & " tn=" & TypeName(v2) & " late=" & v2.Twice(21)

    Dim po1 As Object
    Set po1 = New B216Plain
    Debug.Print "N1-plain-inlineNew-tn=" & TypeName(po1) & " isNothing=" & (po1 Is Nothing)

    Dim po2 As Object
    Set po2 = mP
    Debug.Print "N2-plain-moduleNew-tn=" & TypeName(po2) & " isNothing=" & (po2 Is Nothing)

    Dim pv As Variant
    Set pv = New B216Plain
    Debug.Print "N3-plain-variant-tn=" & TypeName(pv)

    Debug.Print "B216-DONE"
End Sub
