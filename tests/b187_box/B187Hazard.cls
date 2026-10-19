VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Hazard"
Option Explicit

' Same fixture, second arm: what happens when such a slot is handed to a REAL external
' IDispatch. Before the fix the slot held a raw vb6_cls_B187Good*, and the argument cleanup
' on that call is OLE's own VariantClear (not the guarded vb6_ReleaseObject), so the store
' walked into a wild Release: probe exit 0xC0000005. After the fix the slot holds a
' borrowed wrapper and the external collection accepts it, so this answers 999.

Public Function Twice(ByVal n As Long) As Long
    Dim o As New B187Good
    Dim v As Object
    Dim d As Object
    Set v = o
    Set d = CreateObject("Scripting.Dictionary")
    If d Is Nothing Then
        Twice = -777
        Exit Function
    End If
    d.Add "k", v
    If d.Count = 1 Then Twice = 999 Else Twice = 0
End Function
