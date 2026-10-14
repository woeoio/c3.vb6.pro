Option Explicit

' <vbeclipse> regression fixture (knife 64, ledger 305 B173).
' The builtin/runtime argument-conversion arm used to ask ONLY the text predicate
' "is this C expression a whole vb6_VARIANT slot", so a UDT member of type Variant
' (u.UV) was handed raw to a BSTR parameter: vb6_InStr(1, u.UV, ...) = cl C2440 on
' both arches. Left/Len/Replace/Val were fine because those builtins carry a typed
' parameter table and go through the OTHER arm, which asks the AST first.

Private Type U73
    UV As Variant
    UN As Long
End Type

Private m_gV As Variant
Private m_h As New B173Holder

Private Sub TakeByRef(ByRef bv As Variant)
    Debug.Print "B73-11=" & InStr(bv, "t")
End Sub

Private Sub TakeByVal(ByVal pv As Variant)
    Debug.Print "B73-12=" & InStr(pv, "t")
End Sub

Private Sub TakeClsParam(h As B173Holder)
    Debug.Print "B73-18=" & InStr(h.PV, "t")
End Sub

Private Sub Shapes()
    Dim v As Variant
    v = "withform"
    Debug.Print "B73-1=" & InStr(v, "t")
    m_gV = "withform"
    Debug.Print "B73-2=" & InStr(m_gV, "t")
    Dim u As U73
    u.UV = "withform"
    Debug.Print "B73-3=" & InStr(u.UV, "t")
    Debug.Print "B73-4=" & Len(u.UV)
    Debug.Print "B73-5=" & Replace(u.UV, "o", "0")
    Debug.Print "B73-6=" & CStr(u.UV)
    Debug.Print "B73-7=" & Left(u.UV, 2)
    Debug.Print "B73-8=" & Val(u.UV)
    u.UV = 7
    Debug.Print "B73-9=" & InStr(u.UV, "7")
    Debug.Print "B73-10=" & InStr(u.UV, "z")
    Dim a(1 To 2) As Variant
    a(1) = "withform"
    Debug.Print "B73-13=" & InStr(a(1), "t")
    m_h.PV = "withform"
    Debug.Print "B73-14=" & InStr(m_h.PV, "t")
    Debug.Print "B73-15=" & InStr(m_h.PropV(), "t")
    Dim s As String
    s = "withform"
    Debug.Print "B73-16=" & InStr(s, "t")
    Call TakeByRef(v)
    Call TakeByVal(v)
    Call TakeClsParam(m_h)
End Sub

Public Sub Main()
    Shapes
    Dim k As New B173Holder
    k.PV = "withform"
    Debug.Print "B73-19=" & InStr(k.PV, "t")
    Debug.Print "B73-DONE"
End Sub
