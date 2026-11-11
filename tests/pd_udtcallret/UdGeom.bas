Attribute VB_Name = "pdGeom"
Option Explicit

Public Type RectF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Type RGBQuad
    Blue As Byte
    Green As Byte
    Red As Byte
    Reserved As Byte
End Type

' 接收者是 ByRef UDT 形参的模块函数 —— 实参位就是这一刀要折的站。
Public Function TakeRect(ByRef dstRect As RectF, ByRef srcRect As RectF) As Boolean
    dstRect.Left = srcRect.Left
    TakeRect = True
End Function
