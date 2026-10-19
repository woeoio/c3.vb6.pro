VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T88Good"
Option Explicit

' Witness: same release path, a terminate that cannot raise. It is also the PRIMARY clsid for the
' probe, so "the clean class still constructs and its member still answers 42" is on the record
' before the boom release is even attempted.

Public seed As Long

Private Sub Class_Initialize()
    seed = 12345
End Sub

Private Sub Class_Terminate()
    seed = 0
End Sub

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
