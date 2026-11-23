Option Explicit

Public Tag As Long

' 返回类型 = 所在类自己（本刀第一处）
Public Property Get Self As B148Box
    Set Self = Me
End Property

' 返回类型 = 别的工程类（对照组）
Public Property Get Fresh As B148Item
    Set Fresh = New B148Item
End Property

' 返回类型 = As Object（本刀第二处）
Public Property Set Held(ByVal v As Object)
    Set gHold = v
End Property

Public Property Get Held() As Object
    Set Held = gHold
End Property

Private Sub Class_Terminate()
    gBoxTerms = gBoxTerms + 1
End Sub
