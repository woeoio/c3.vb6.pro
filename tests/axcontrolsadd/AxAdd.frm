VERSION 5.00
Begin VB.Form AxAdd 
   Caption         =   "AxAdd"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   7200
   LinkTopic       =   "AxAdd"
   ScaleHeight     =   3600
   ScaleWidth      =   7200
   Begin VB.Timer tChk 
      Enabled         =   -1   'True
      Interval        =   200
      Left            =   6600
      Top             =   120
   End
End
Attribute VB_Name = "AxAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Fix <vbeclipse> 2026-10-06: 运行期 Me.Controls.Add 的 ActiveX 宿主回归。
' 旧路径: SetClientSite + 写死 400x300 rect / 错算的 HIMETRIC extent, 不 DoVerb
' 不登记转发/绘制 —— windowless OCX 一片空白。现在与设计期 vb6_OcxHost_Create
' 同一套完整宿主流程 (windowless 三件套 + DoVerb INPLACEACTIVATE + 客户区
' SetObjectRects + IViewObject 绘制登记)。
' 断言: ADD=OK (实例化+激活不崩); 采样由 C3_OCX_TRACE=1 的 stderr 配合人工核验。

Private addOk As Boolean
Private tick As Long

Private Sub Form_Load()
    Dim o As Object
    Set o = Me.Controls.Add("NewTabCtl.NewTab", "nt1")
    addOk = Not (o Is Nothing)
    Debug.Print "ADD=" & IIf(addOk, "OK", "FAIL")
End Sub

Private Sub tChk_Timer()
    tick = tick + 1
    If tick >= 15 Then Unload Me   ' ~3s 后收尾
End Sub
