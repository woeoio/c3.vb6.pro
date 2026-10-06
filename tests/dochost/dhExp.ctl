VERSION 5.00
Begin VB.UserControl dhExp 
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   1200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   2400
End
Attribute VB_Name = "dhExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> 账 #217 第二刀的判据夹子 (编译面 / 开着 Option Explicit 那一面)。
' 改前每条 `UserControl.x` / `Ambient.x` / `Extender.x` / `VBA.x` 都在语义层落进"未声明的标识符":
' 这一面每形一条 VB3001 (两份真工程合起来 946 条)，另一面 (dhImp.ctl) 每形一枚没人引用的
' vb6_VARIANT 隐式局部。发码那侧一直是对的 (成员与类型都由 kHostPseudoRows 那张表回答，账 #159)，
' 所以钉两头：① 诊断里不许再出现这四个对象名；② 发码符号一个都不能变。

Public Function HostVals() As String
    HostVals = CStr(UserControl.ScaleWidth) & "/" & CStr(UserControl.hWnd)
End Function

Public Function EnvVals() As String
    EnvVals = CStr(Ambient.UserMode) & "/" & VBA.UCase("ab") & "/" & CStr(VBA.Len("abcd"))
End Function

Public Sub PushTag(v As String)
    Extender.Tag = v
End Sub
