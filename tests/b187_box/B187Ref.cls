VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Ref"
Option Explicit

' Ledger 305 B187 fifth leg = the ByRef station. A formal written "p As Object" with no
' ByVal keyword is ByRef (VB6 default), so C hands the callee void** -- the CALLER's own
' slot. Boxing therefore has to happen IN PLACE: a temp would silently drop everything the
' callee writes back. One arm, four shapes:
'   1e6  a late-bound call on the formal answers 42      (BASE 0 / this knife 42)
'   1e4  the callee's "Set p = Nothing" lands in the slot (both 1)
'   1e2  the callee's "Set p = <instance>", early-bound  (both 42)
'   1    the callee's "Set p = CreateObject(...)" stays live (both 1)
' BASE (knife 84, ByVal station only) = 0 + 10000 + 4200 + 1 = 14201.
' This knife = 42000000 + 10000 + 4200 + 1 = 42014201.

Private g_late As Long

Public Function Twice(ByVal n As Long) As Long
    Dim o As B187Good
    Dim cleared As Long
    Dim replaced As Long
    Dim live As Long

    Set o = New B187Good
    g_late = 0
    ReadFormal o
    Set o = New B187Good
    ClearFormal o
    If o Is Nothing Then cleared = 1
    Set o = New B187Good
    GiveFormal o
    On Error Resume Next
    replaced = o.Twice(21)
    On Error GoTo 0
    Set o = New B187Good
    GiveExt o
    If o Is Nothing Then live = 0 Else live = 1

    Twice = (g_late * 1000000) + (cleared * 10000) + (replaced * 100) + live
End Function

Private Sub ReadFormal(p As Object)
    On Error Resume Next
    g_late = p.Twice(21)
End Sub

Private Sub ClearFormal(p As Object)
    Set p = Nothing
End Sub

Private Sub GiveFormal(p As Object)
    Dim g As B187Good
    Set g = New B187Good
    Set p = g
End Sub

Private Sub GiveExt(p As Object)
    Set p = CreateObject("Scripting.Dictionary")
End Sub
