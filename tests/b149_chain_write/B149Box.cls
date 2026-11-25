Option Explicit

Public Tag As Long

Private mHold As Object
Private mChild As B149Box

' 子对象是**另一枚**实例: 写 `o.Child.Held = it` 必须落在它身上, 不是 o
Public Property Get Child() As B149Box
    If mChild Is Nothing Then Set mChild = New B149Box
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
