VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' C2440 §56.30 夹具: 类方法返回工程 UDT (pd2DPath.GetPathBoundariesF 的形)。
Public Function GetRect() As URect
    GetRect.a = 1
End Function
