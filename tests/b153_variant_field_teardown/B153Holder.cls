Option Explicit

' 账 #305 §11.134 (第一百五十三刀): `As Variant` 字段出厂那一枚。
'
' Variant 槽与 `As Object` 字段是两条账: 槽里可以是对象 (VT_DISPATCH, VariantCopy
' 会 AddRef)、也可以是 BSTR (自带一枚)。出厂那一站以前整支不出现 —— 槽里那一枚
' 既没人放、也没人清 (漏, 不是崩)。

Private mV As Variant

Public Property Get V() As Variant
    V = mV
End Property

Public Property Let V(ByVal x As Variant)
    mV = x
End Property

Public Property Set V(ByVal x As Variant)
    Set mV = x
End Property

' 槽里放字符串 (BSTR 那一枚归槽自己)
Public Sub PutText(ByVal s As String)
    mV = s
End Sub

Public Property Get TextLen() As Long
    TextLen = Len(CStr(mV))
End Property
