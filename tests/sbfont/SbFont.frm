VERSION 5.00
Begin VB.Form SbFontForm
   Caption         =   "SbFontForm"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7000
   LinkTopic       = "Form1"
   ScaleHeight     =   2400
   ScaleWidth      =   7000
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.StatusBar StatusBar1
      Align          =   2                     ' AlignBottom
      Height         =   255
      Left           =   0
      Top            =   2145
      Width          =   7000
      _ExtentX       =   12348
      _ExtentY       =   450
      _Anchor        =   10
      FontName        =   "MS Sans Serif"
      FontSize        =   8
      Style          =   0                     ' sbrNormal
      Panels(1)      =   "Ready"
         .Key        =   "fx"
         .Style      =   0                     ' sbrText
         .AutoSize   =   0                     ' sbrFixed
         .Width      =   900
         .MinWidth   =   20
      End
      Panels(2)      =   "WWWWWWWWWW"
         .Key        =   "ct"
         .Style      =   0                     ' sbrText
         .AutoSize   =   2                     ' sbrContents
         .MinWidth   =   20
      End
   End
   Begin MSComctlLib.StatusBar StatusBar2
      Align          =   1                     ' AlignTop
      Height         =   255
      Left           =   0
      Top            =   0
      Width          =   7000
      _ExtentX       =   12348
      _ExtentY       =   450
      _Anchor        =   9
      FontName        =   "MS Sans Serif"
      FontSize        =   20
      Style          =   0                     ' sbrNormal
      Panels(1)      =   "WWWWWWWWWW"
         .Key        =   "ct2"
         .Style      =   0                     ' sbrText
         .AutoSize   =   2                     ' sbrContents
         .MinWidth   =   20
      End
   End
End
Attribute VB_Name = "SbFontForm"
Option Explicit

' 账 #205 的测量面：状态条那两处问字体以前都是**裸问窗口**，而状态条是 RTL 自己注册的
' 窗口类 ⇒ 那一问恒回 NULL，量的（sbrContents 的排版宽）与画的都按 DC 的默认字体。
' 排版结果从 VB 里没有现成的门（Panels(i).Width 读的是**请求值**，不是排版后的宽），
' 所以判据问窗口本人：SB_GETPARTS (WM_USER+6) 交回各格的右边界。
Private Declare Function SbGetParts Lib "user32" Alias "SendMessageW" (ByVal hWin As LongPtr, _
    ByVal msg As Long, ByVal n As Long, ByRef parts As Any) As Long

Function TF(b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function

Private Sub Form_Load()
    Dim pt1(0 To 3) As Long
    Dim pt2(0 To 3) As Long
    Dim wCt8 As Long            ' StatusBar1 那格 sbrContents 的实际宽（8pt）
    Dim wCt20 As Long           ' StatusBar2 同一串文字的实际宽（20pt）
    Dim wFix As Long            ' StatusBar1 那格 sbrFixed 的右边界（不许动）
    Dim wAfter As Long
    Dim h1 As Long
    Dim ok1 As Boolean
    Dim ok2 As Boolean
    Dim ok3 As Boolean

    h1 = 0
    SbGetParts StatusBar1.hWnd, 1030, 2, pt1(0)
    SbGetParts StatusBar2.hWnd, 1030, 1, pt2(0)
    wFix = pt1(0)
    wCt8 = pt1(1) - pt1(0)
    wCt20 = pt2(0)
    Debug.Print "SF00-RAW pf8=" & CStr(StatusBar1.FontPixelHeight) & _
                " pf20=" & CStr(StatusBar2.FontPixelHeight)
    Debug.Print "SF01-RAW fix=" & CStr(wFix) & " ct8=" & CStr(wCt8) & " ct20=" & CStr(wCt20)

    ' ① 同一串文字，20pt 那枚必须比 8pt 那枚宽（改前：两处恒等 = 字体没参与排版）
    ok1 = (wCt8 > 0) And (wCt20 > wCt8)
    Debug.Print "SF02-CONTENTS-TRACKS-FONT=" & TF(ok1)

    ' ② 运行期改字号后重排，那格必须变宽
    StatusBar1.FontSize = 20
    StatusBar1.Panels(2).Text = "WWWWWWWWWW"
    SbGetParts StatusBar1.hWnd, 1030, 2, pt1(0)
    wAfter = pt1(1) - pt1(0)
    h1 = pt1(0)
    Debug.Print "SF03-RAW after=" & CStr(wAfter) & " pfAfter=" & CStr(StatusBar1.FontPixelHeight)
    ok2 = (wAfter > wCt8)
    Debug.Print "SF04-RUNTIME-FONT=" & TF(ok2)

    ' ③ 护栏：显式给过 Width 的那格右边界不许因为换字体而挪
    ok3 = (h1 = wFix)
    Debug.Print "SF05-FIXED-EDGE-UNCHANGED=" & TF(ok3)

    Debug.Print "SBFONT-DONE"
    Unload Me
End Sub
