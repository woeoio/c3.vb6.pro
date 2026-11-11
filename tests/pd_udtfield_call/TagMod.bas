Attribute VB_Name = "TagMod"
Option Explicit

' C2197 §56.29 夹具: UDT 字段声明成工程类 (`As Stack`), With 块内
' `.IDs.GetString(i)` —— 修前推断不出字段的类, 外层调用落裸名表,
' 借到 Ser.GetString 的 3 参 ⇒ 多发缺省实参 C2197。
Public Type TagRec
    IDs As Stack
End Type

Public Sub Main()
    Dim t As TagRec
    Set t.IDs = New Stack
    With t
        Debug.Print .IDs.GetString(0)
    End With
End Sub

' 同一族的数组形 (PD #9 pdSpriteSheet.c `m_ImageCache(i).spriteNames`):
' 接收者是 UDT 数组元素, 类推断也要落到字段自身的 vb6_cls_Stack*.
Public Sub MainArr()
    Dim arr() As TagRec
    ReDim arr(0 To 1)
    Set arr(0).IDs = New Stack
    Debug.Print arr(0).IDs.GetString(0)
End Sub
