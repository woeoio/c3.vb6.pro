Attribute VB_Name = "ModU"
Option Explicit

' 跨模块裸名读真全局: 修前 C2065 (gShared 被强制 static 且无 extern)。
Public Function GetG() As Long
    GetG = gShared
End Function
