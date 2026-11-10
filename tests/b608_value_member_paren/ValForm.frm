VERSION 5.00
Begin VB.Form ValForm 
   Caption         =   "ValForm"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   LinkTopic       =   "ValForm"
   ScaleHeight     =   2400
   ScaleWidth      =   4800
   Begin VB.TextBox cboUpdates 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.TextBox cboUpdates 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   350
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   800
   End
End
Attribute VB_Name = "ValForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 探针夹具 (§B94, PhotoDemon C2064 第 3 族): 值成员被写成"带空括号"。
' 这三种形态发码时成员读已经产出完整 C 表达式, 再追 `()` 就是"调用非函数" ⇒ C2064。
'   1) `cboUpdates.UBound()`  → `vb6_CtrlArr_UBound(&vb6_arr_cboUpdates)  /* ctrl array UBound */`
'   2) `cmdOK.FontSize()`     → `vb6_GetControlFontSize(vb6_hwnd_cmdOK)  /* ctrl prop read */`
'   3) `PDImages.GetActiveImage.Width()` → `...->Width  /* class var .Width field */`
' 负控: 同一批成员的**普通读法**(无括号)与真方法/带参读法必须逐字不变。

Private Sub Form_Load()
    Dim n As Long
    n = cboUpdates.UBound()
    n = cboUpdates.LBound()
    n = cmdOK.FontSize()
    cmdOK.FontSize = 12
    n = cboUpdates.Count
    n = PDImages.GetActiveImage.Width()
    n = PDImages.GetActiveImage.Height
End Sub
