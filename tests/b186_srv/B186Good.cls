VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B186Good"
Option Explicit

' Ledger 305 B186 (knife 78): the witness class -- Class_Initialize runs and succeeds, so the
' client must still get a usable object in the SAME run that asks for B186Boom. Without this arm
' "the factory now answers an error for everything" would pass.

Public seed As Long

Private Sub Class_Initialize()
    seed = 12345
End Sub

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
