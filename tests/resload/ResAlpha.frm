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
   Begin VB.Image imgN 
      Height          =   960
      Left            =   2400
      Top             =   480
      Width           =   1440
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

'Fix <vbeclipse> 2026-10-10 (zhang 301, 26th cut): the NON-stretch picture face had zero
'  behavior coverage. imgN carries no Stretch line (= 0) and a 96 x 64 px box; the picture is
'  alpha.png (48 x 48, left half red / right half transparent). Without stretching the red run
'  must stop at the picture's natural pixel size (~23), not at the box edge (95).
'  The witness is the RED run, not "any non-background pixel": the first probe read 95 because
'  the fill branch paints the whole box too, so a color-blind scan measures the fill instead of
'  the draw.

Private Declare Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As LongPtr, ByVal X As Long, ByVal Y As Long) As Long

Private done As Boolean
Private tick As Long

Private Sub Form_Load()
    Me.BackColor = vbBlue
    imgA.BackColor = vbBlue
    Set imgA.Picture = LoadResPicture("IMG3", "PNG")
    imgN.BackColor = vbBlue
    Set imgN.Picture = LoadResPicture("IMG3", "PNG")
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
    Dim hDcN As LongPtr
    Dim px As Long, kk As Long
    Dim lastRed As Long, lastNB As Long, nWhite As Long
    lastRed = -1
    lastNB = -1
    nWhite = 0
    hDcN = GetDC(imgN.hWnd)
    For kk = 0 To 119
        px = GetPixel(hDcN, kk, 24)
        If px = 255 Then lastRed = kk
        If px >= 0 And px <> 16711680 Then lastNB = kk
        If px = 16777215 Then nWhite = nWhite + 1
    Next
    ReleaseDC imgN.hWnd, hDcN
    Dim sNs As String
    If lastRed >= 18 And lastRed <= 30 Then sNs = "True" Else sNs = "False"
    Debug.Print "NS01-NATURAL=" & sNs
    Debug.Print "NS01-RAW red=" & CStr(lastRed) & " nonblue=" & CStr(lastNB) & " white=" & CStr(nWhite)

    Unload Me
End Sub
