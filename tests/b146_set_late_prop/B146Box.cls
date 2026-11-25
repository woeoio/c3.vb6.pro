Option Explicit

' 每次都新造一枚 (Property Get 交回的是加了引用的实例)
Public Property Get Fresh() As B146Item
    Set Fresh = New B146Item
End Property

' 反面证人: 交回的是模块级**共用**的那一枚
Public Property Get SameObj() As B146Item
    Set SameObj = gKeep
End Property
