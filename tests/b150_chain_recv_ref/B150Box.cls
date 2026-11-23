Option Explicit

Public Tag As Long

Private mHold As Object
Private mChild As B150Box

Public Property Get Child() As B150Box
    If mChild Is Nothing Then Set mChild = New B150Box
    Set Child = mChild
End Property

Public Property Set Held(ByVal v As Object)
    Set mHold = v
End Property

Public Property Get Held() As Object
    Set Held = mHold
End Property

Private Sub Class_Terminate()
    gBoxTerms = gBoxTerms + 1
End Sub
