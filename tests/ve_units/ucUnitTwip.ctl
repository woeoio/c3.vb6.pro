VERSION 5.00
Begin VB.UserControl ucUnitTwip 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2400
   ScaleHeight     =   1200
   ScaleMode       =   1  'Twips
   ScaleWidth      =   2400
End
Attribute VB_Name = "ucUnitTwip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> 回归夹子 (ve_units): 控件自己的坐标系读数。
' TextWidth/ScaleWidth 必须按 .ctl 声明的 ScaleMode 交出 (账 #175/#177)。

Private Sub UserControl_Initialize()
    Debug.Print "I-MODE=" & UserControl.ScaleMode & " I-TW=" & UserControl.TextWidth("MMMM") & " I-SW=" & UserControl.ScaleWidth
End Sub

Public Function Ctx() As String
    ' 账 #179: 这两个值必须**不随调用方变** —— 容器里调也要读到这一枚控件自己的
    ' ScaleMode/ScaleWidth。改前实测容器里调 uTw.Ctx() 得到 mode=3 (另一枚控件留下的
    ' 残值), 控件自己调得到 mode=1。hWnd 那一族不在这里钉 (那是台账 #159)。
    Ctx = UserControl.ScaleMode & "/" & UserControl.ScaleWidth
End Function
Public Function TW(ByVal s As String) As Long
    TW = UserControl.TextWidth(s)
End Function

Public Function TH(ByVal s As String) As Long
    TH = UserControl.TextHeight(s)
End Function

Public Function SW() As Long
    SW = UserControl.ScaleWidth
End Function

Public Function HwOk() As String
    ' 账 #159: 这条比较改前发成 vb6_VarCmpLongNe(&vb6_UserControl_hWnd, 0) —— 拿
    ' 8 字节 void* 全局的**地址**当 vb6_VARIANT* 传 (RTL 签名第一形参是
    ' vb6_VARIANT*) ⇒ 读到的是越界垃圾, 恒不勾。现在类型由 kHostPseudoRows 答
    ' LongPtr ⇒ 走 C 直比 (-(vb6_UserControl_hWnd != 0))。
    HwOk = CStr(UserControl.hWnd <> 0)
End Function
Public Function HwVs() As String
    ' 第二头证人: 同一条读法再问一次「hWnd 与 hDC 是不是两个不同的值」。
    ' 只钉 HwOk 一条时, 「两个成员都被读成同一个垃圾」也能蒙过去。
    HwVs = CStr(UserControl.hWnd <> UserControl.hDC)
End Function

Public Function CntOk() As String
    ' 账 #180 (B19): ContainerHwnd 以前住在 int32_t 槽里 —— x64 上写进来就截断, 而
    ' 语料里 VBFlexGrid.ctl 把它当 HWND 直接交给 MapWindowPoints / GetWindowLongW。
    ' 两头判据: ① 容器句柄非零 (没人填值 / 填错槽都会红); ② 容器不是本控件自己
    ' (读错成 hWnd 会红)。宽度这一头由 scripts/check_host_pseudo_table.ps1 钉
    ' (表答 LongPtr ⇒ RTL 声明必须是指针宽度, 不对上就红)。
    CntOk = CStr(UserControl.ContainerHwnd <> 0 And UserControl.ContainerHwnd <> UserControl.hWnd)
End Function
Public Function CntStr() As String
    ' 第三条读数: 同一个值走 CStr 那一路必须打出**数字**而不是空串 (账 #159 那族:
    ' void* 落进 _Generic 的 default: vb6_VariantObject 就是空串)。
    CntStr = CStr(UserControl.ContainerHwnd)
End Function

Public Sub MoveParent(ByVal l As Long, ByVal t As Long, ByVal w As Long, ByVal h As Long)
    ' account 251: the four arguments of UserControl.Parent.Move are VB-side numbers,
    ' i.e. TWIPS here (the same unit the host model hands back for Parent.Left /
    ' .Width, and the same unit a .frm records). czUI's fullscreen arm saves what it
    ' read out of the host model and moves the parent with it, so a raw MoveWindow
    ' here puts twips into a pixel slot.
    UserControl.Parent.Move l, t, w, h
End Sub
