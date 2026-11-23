Option Explicit

' 每次都新造一枚 (Property Get 交回的是加了引用的实例)
Public Property Get Fresh() As B144Item
    Set Fresh = New B144Item
End Property

' 反面证人: 交回的是模块级**共用**的那一枚 —— With 临时量退自己那一枚之后,
' 它必须仍然活着 (模块自己那一枚还在)。
Public Property Get SameObj() As B144Item
    Set SameObj = gKeep
End Property
