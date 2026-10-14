Attribute VB_Name = "B171Main"
Option Explicit

' ledger 305 B171 (knife 62): the SECOND owner of 'is this C a whole slot' is the
' text-side predicate. It did not recognize a class-field slot read, so the two
' argument-packing stations stuffed a vb6_VARIANT into a scalar field (C2440, no
' exe on either arch). Three faces pinned here: the colliding name (PV also is an
' intrinsic function name) must answer the same as the clean one (PZ); a numeric
' slot must keep its vt (not be flattened to a string); ByRef write-back must land.

Private Sub TakeSlot(v As Variant)
    Debug.Print "  inner vt=" & VarType(v) & " s=[" & CStr(v) & "]"
End Sub

Private Function EchoBack(v As Variant) As String
    EchoBack = "[" & VarType(v) & "|" & CStr(v) & "]"
End Function

Private Sub Mutate(v As Variant)
    v = "changed"
End Sub

Sub Main()
    Dim h As New B171Holder
    Dim s As String
    Dim n As Long
    Dim w As Variant
    s = "withform"
    n = 7
    h.PV = s
    h.PZ = s
    h.PN = n
    h.PS = "strfield"
    w = s
    Debug.Print "B171-1=" & EchoBack(h.PZ)
    Debug.Print "B171-2=" & EchoBack(h.PV)
    Debug.Print "B171-3=" & EchoBack(w)
    Debug.Print "B171-4=" & EchoBack(h.PN)
    Debug.Print "B171-5=" & EchoBack(h.PS)
    Call TakeSlot(h.PN)
    Debug.Print "B171-6=" & Len(h.PV)
    Debug.Print "B171-7=" & Left(h.PV, 4)
    Debug.Print "B171-8=" & CStr(h.PV = s)
    Debug.Print "B171-9=[" & h.PV & "]"
    With h
        Debug.Print "B171-10=" & EchoBack(.PN)
        Debug.Print "B171-11=" & Len(.PZ)
    End With
    Call Mutate(h.PV)
    Debug.Print "B171-12=" & EchoBack(h.PV)
    Debug.Print "B171-DONE"
End Sub
