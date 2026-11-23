Option Explicit

' 一枚"每次都新造"的属性 (Property Get 交回的是加了引用的实例)
Public Property Get Fresh() As B142Item
    Set Fresh = New B142Item
End Property

' 一枚存着的对象字段 —— With 它只是借, 块尾一枚都不许退
Public Bare As B142Item