VERSION 5.00
Begin VB.Form FormB 
   Caption         =   "FormB"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
End
Attribute VB_Name = "FormB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 注册在前, 带 1 枚 Optional —— 裸键 lookupModule 先命中它,
' 跨窗体 `FormA.UACT` 若借这张表就多发 (0, 0) ⇒ C2197。
Public Sub UACT(Optional ByVal x As Long = 0)
End Sub
