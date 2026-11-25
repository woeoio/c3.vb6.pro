Option Explicit

Public Tag As Long

Private mHold As Object
Private mChild As B151Box

Public Property Get Child() As B151Box
    If mChild Is Nothing Then
        Set mChild = New B151Box
        mChild.Tag = 99
    End If
    Set Child = mChild
End Property

Public Property Get Held() As Object
    Set Held = mHold
End Property

Public Property Set Held(ByVal v As Object)
    Set mHold = v
End Property

Private Sub Class_Initialize()
    gBoxMade = gBoxMade + 1
End Sub

Private Sub Class_Terminate()
    gBoxTerms = gBoxTerms + 1
End Sub
