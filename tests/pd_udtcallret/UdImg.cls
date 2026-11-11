VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "pdImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 零参方法返回工程 UDT —— VB6 允许调用点不写括号 (实参位/With 目标都是这个形状)。
Friend Function GetBoundaryRectF() As RectF
    GetBoundaryRectF.Left = 10
    GetBoundaryRectF.Top = 20
    GetBoundaryRectF.Right = 30
    GetBoundaryRectF.Bottom = 40
End Function
