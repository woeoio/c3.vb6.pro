Attribute VB_Name = "Tools_Pencil"
Option Explicit

' 跨模块零参调用带 Optional 形参 —— PhotoDemon `Tools_Pencil.GetModifiedUpdateRectF()`
' 那一站的最小复现（返回工程 UDT，调用点不写实参）。
Public Function GetRect(Optional ByVal resetRectAfter As Boolean = True) As RectF
    GetRect.Left = 1
    GetRect.Top = 2
    GetRect.Right = 3
    GetRect.Bottom = 4
End Function
