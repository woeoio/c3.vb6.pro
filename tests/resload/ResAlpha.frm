VERSION 5.00
Begin VB.Form ResAlpha 
   Caption         =   "ResAlpha"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   4800
   LinkTopic       =   "ResAlpha"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   Begin VB.Timer tChk 
      Enabled         =   -1   'True
      Interval        =   100
      Left            =   4200
      Top             =   120
   End
   Begin VB.Image imgA 
      Height          =   960
      Left            =   480
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1920
   End
End
Attribute VB_Name = "ResAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Fix <vbeclipse> 2026-10-06: Picture 现代 RGBA 透明回归。
' imgA BackColor = 蓝, Picture = alpha.png (左半红 / 右半全透明)。
' 透明区必须透出 Image 自己的蓝底 (旧路径白底+Render 会顶成白色)。
' tChk 在首次绘制后 GetPixel 采样右半 (透明区) 与左半 (红) 的实际屏色。

Private done As Boolean
Private tick As Long

Private Sub Form_Load()
    Me.BackColor = vbBlue
    imgA.BackColor = vbBlue
    Set imgA.Picture = LoadResPicture("IMG3", "PNG")
    Debug.Print "PICSET=" & (Not (imgA.Picture Is Nothing))
End Sub

Private Sub tChk_Timer()
    ' FIX 236: the close used to sit BEHIND the done-guard. Once done was set, every
    ' later tick returned before the counter got its increment, so the threshold was
    ' never reached and the exe ran until the harness killed it (CI: run timeout 60s).
    ' The counter now advances first, so every path is bounded.
    tick = tick + 1
    If imgA.Picture Is Nothing Then
        If tick >= 30 Then Unload Me
        Exit Sub
    End If
    If done Then Exit Sub
    done = True
    Debug.Print "PAINTED=" & (Not (imgA.Picture Is Nothing))
    Unload Me
End Sub
