VERSION 5.00
Begin VB.UserControl ParentCtl 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1200
End
Attribute VB_Name = "ParentCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' <vbeclipse> 回归夹子 (pd_parentctl): VB6 UserControl 内建 `Parent`。
' 语义 = **宿主容器对象**（窗体 / 容器控件），不是容器窗口句柄；`Parent.Controls`
' 就是父容器的控件集合，与 `UserControl.ParentControls` 是同一件事。
' PhotoDemon pdCommandBar.ctl 有 7 处 `For Each eControl In Parent.Controls`。
' 修前：裸 `Parent` 报假 VB3001，产物把成员访问挂在未声明的符号上。

Public Function ParentControlCount() As Long
    Dim c As Object
    Dim hostObj As Object
    Dim n As Long

    ' 值位：宿主容器对象（不是容器窗口句柄）。
    Set hostObj = Parent

    For Each c In Parent.Controls
        n = n + 1
    Next c

    If hostObj Is Nothing Then n = -1
    ParentControlCount = n
End Function

' 限定写法 `UserControl.Parent.<成员>`（PhotoDemon pdCommandBar*.ctl 6 处）。
' Parent 的静态类型是 Object（晚绑定），真值是宿主窗体；读/写/方法都必须由
' 宿主模型按名答，不能把成员挂在 `vb6_UC_ParentObject()` 这个 void* 上
' （那是 C2224；整编 pdCommandBar.c / pdCommandBarMini.c 6 条）。
Public Function ParentExtenderProbe() As String
    Dim s As String

    s = UserControl.Parent.Caption
    s = UserControl.Parent.Name
    If UserControl.Parent.Visible Then s = s & "!"
    UserControl.Parent.Visible = False
    UserControl.Parent.Hide

    ParentExtenderProbe = s
End Function
