Option Explicit

' 账 #305 §11.133 (第一百五十二刀): `As Object` 字段出厂那一枚引用。
'
' 腿②(覆写)/腿③(置空) 从第一刀起就发 `vb6_ReleaseObject(&字段)`，只有
' `_Destroy` 这一站缺着 —— 字段里那一枚没人退, 被它握着的实例永不走 Class_Terminate。
' 改前出厂只放「类类型字段」(`As B151Box` 那种), `As Object` 一格整支不出现。

Private mHold As Object
Private mSelf As Object

Public Property Get Held() As Object
    Set Held = mHold
End Property

Public Property Set Held(ByVal v As Object)
    Set mHold = v
End Property

' 裸实例进 As Object 字段: 出厂放手必须被 `vb6_ReleaseObject` 自己的 Fix 191 判据拦下
' (非可分派只清指针, 不越权解引用 lpVtbl) —— 本夹具就是这条判据的钉子。
Public Sub KeepMe()
    Set mSelf = Me
End Sub

Public Property Get SelfIsNothing() As Boolean
    SelfIsNothing = (mSelf Is Nothing)
End Property
