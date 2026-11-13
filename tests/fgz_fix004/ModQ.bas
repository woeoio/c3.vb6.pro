Attribute VB_Name = "ModQ"
Option Explicit

Public gDuo As Long

Public Sub SetDuo(ByVal n As Long)
    gDuo = n
End Sub

Public Function GetDuo() As Long
    GetDuo = gDuo
End Function
