Option Explicit

Public Tag As Long

Private Sub Class_Initialize()
    gItMade = gItMade + 1
End Sub

Private Sub Class_Terminate()
    gItTerms = gItTerms + 1
End Sub
