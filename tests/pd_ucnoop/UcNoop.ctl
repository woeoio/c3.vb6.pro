VERSION 5.00
Begin VB.UserControl UcNoop 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
End
Attribute VB_Name = "UcNoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' 只有 Initialize 一枚过程: 任何别的成员名都进不了 resolveClassMemberCall 的表,
' 于是 `ucTest.SetFocus` 落到"未建模 → 空操作"那条腿。
Private Sub UserControl_Initialize()
End Sub
