VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B186Held"
Option Explicit

' Ledger 305 B186 (knife 78): witness for the frame ordering. A body that owns its own
' On Error catches first, so Class_Initialize completes and the client sees S_OK plus a live
' object -- the new factory-side frame must not steal it (same shape as the HANDLED arm of
' B185 on the IDispatch side).

Public handled As Long

Private Sub Class_Initialize()
    On Error GoTo H
    Err.Raise 123
    Exit Sub
H:
    handled = Err.Number
End Sub

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
