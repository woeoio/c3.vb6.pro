VERSION 5.00
Begin VB.UserControl dhImp 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   1200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   2400
End
Attribute VB_Name = "dhImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' <vbeclipse> 账 #217 第二刀的另一面: 这里**故意不写** Option Explicit。
' 改前这四枚文档隐式对象名会走 VB6 隐式声明那一路登记成 Variant，发码就在每枚用到它的
' 过程序言真发一枚 `vb6_VARIANT UserControl = ...;` (实测 `VBA` 那枚也一样，还额外带一对
' #pragma push_macro/pop_macro 护着) —— 全是没人读过的死局部。现在 AST 里根本不留名字。

Public Function HostVals2() As String
    HostVals2 = CStr(UserControl.ScaleWidth) & "/" & CStr(UserControl.hWnd)
End Function

Public Function EnvVals2() As String
    EnvVals2 = CStr(Ambient.UserMode) & "/" & VBA.UCase("ab")
End Function
