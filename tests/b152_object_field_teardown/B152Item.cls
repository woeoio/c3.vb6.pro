Option Explicit

Public Tag As Long

Private Sub Class_Initialize()
    gItemMade = gItemMade + 1
End Sub

Private Sub Class_Terminate()
    gItemTerms = gItemTerms + 1
End Sub
