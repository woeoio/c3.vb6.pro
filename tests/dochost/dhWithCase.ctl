VERSION 5.00
Begin VB.UserControl dhWithCase 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
End
Attribute VB_Name = "dhWithCase"
Option Explicit
' ledger 278 knife 11: the object segment of vb6_<Obj>_<member> comes from the table, so
' a VB6 source that spells the built-in object name in another casing (VB is case-insensitive)
' still lands on the RTL symbol. Before the collapse the backend copied the source spelling
' straight into the C identifier: `With usercontrol` -> vb6_usercontrol_hWnd (nobody declares
' that) -> C2065 in a real build, while the same document written `With UserControl` worked.

Public Sub LowerSelf()
    With usercontrol
        Debug.Print .ScaleWidth
    End With
End Sub

Public Sub LowerQualified()
    Dim n As LongPtr
    n = usercontrol.hDc
End Sub
