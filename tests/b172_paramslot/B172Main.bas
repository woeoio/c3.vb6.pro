Attribute VB_Name = "B172Main"
Option Explicit

' ledger 305 B172 (knife 63): the THIRD station that emits a class data-field read.
' A field read through a LOCAL object variable was taught to carry the slot tail note in
' knife 62, but the same expression through a class PARAMETER is emitted by another
' station (Fix 088c) and wrote no tail note at all, so "is this a whole slot" answered
' no for it. Result: InStr(h.PV, ...) handed a vb6_VARIANT to a BSTR parameter = C2440,
' no exe on either arch. The param face must now answer the same as the local face.

Private Sub TakeObj(h As B172Holder)
    Debug.Print "B172-1=" & InStr(h.PV, "t")
    Debug.Print "B172-2=" & InStr(h.PZ, "t")
    Debug.Print "B172-3=" & Replace(h.PV, "o", "0")
    Debug.Print "B172-4=" & InStr(h.PN, "7")
    Debug.Print "B172-5=" & InStr(h.PS, "t")
    Debug.Print "B172-6=" & InStr(h.PV, "z")
    Debug.Print "B172-7=" & Len(h.PV)
    With h
        Debug.Print "B172-8=" & InStr(.PZ, "t")
    End With
End Sub

Public Sub Main()
    Dim h As New B172Holder
    h.PV = "withform"
    h.PZ = "withform"
    h.PN = 7
    h.PS = "withform"
    Call TakeObj(h)
    Debug.Print "B172-9=" & InStr(h.PV, "t")
    Debug.Print "B172-10=" & Replace(h.PZ, "o", "0")
    Debug.Print "B172-DONE"
End Sub
