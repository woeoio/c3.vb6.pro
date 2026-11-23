Option Explicit

Public Tag As Long

Private Sub Class_Initialize()
    gMade = gMade + 1
End Sub

Private Sub Class_Terminate()
    gTerms = gTerms + 1
End Sub
