Attribute VB_Name = "GDI_Plus"
Option Explicit

Public Type RectF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

Public Function IntersectRectF(ByRef dstRect As RectF, ByRef srcRect1 As RectF, ByRef srcRect2 As RectF) As Boolean
    dstRect.Left = srcRect1.Left
    IntersectRectF = True
End Function
