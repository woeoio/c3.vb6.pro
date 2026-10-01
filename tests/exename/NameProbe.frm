VERSION 5.00
Begin VB.Form NameProbe 
   Caption         =   "NameProbe"
   ClientHeight    =   900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2400
   LinkTopic       =   "NameProbe"
   ScaleHeight     =   900
   ScaleWidth      =   2400
   Begin VB.Timer t1 
      Enabled         =   -1   'True
      Interval        =   40
      Left            =   1900
      Top             =   120
   End
   Begin VB.Label lbl1 
      Caption         =   "name probe"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "NameProbe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 账 #166 的证人：这枚工程的 `ExeName32` 故意与 vbp 文件名**不同**（文件名 NameProbe、
' 产物 RenamedProbe.exe，还带 `.exe` 后缀 —— VB6 两种写法都允许）。它存在的唯一理由就是
' 让"按文件名去找产物"那条口径当场红：改之前 Test-Vbp 找 NameProbe.exe ⇒ FAIL (no exe)，
' 改之后走单一权威（读 .vbp 的 ExeName32）⇒ 正常。
Private gTick As Long

Private Sub t1_Timer()
    gTick = gTick + 1
    If gTick = 1 Then
        Debug.Print "NP-ok=Y/tick=" & CStr(gTick)
        Debug.Print "EXENAME-DONE"
        t1.Enabled = False
        Unload Me
    End If
End Sub
