VERSION 5.00
Begin VB.Form FrmUcNoop 
   Caption         =   "FrmUcNoop"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Begin PdUcNoop.UcNoop ucTest 
      Height          =   375
      Left            =   120
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FrmUcNoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' <vbeclipse> C2065 族① 判据夹子: 与 PhotoDemon `Forms\File_Export_Layers.frm:214`
' (`cmdExportSettings.SetFocus`, pdButton UC) 同形。UC 上没建模的方法发成空操作,
' 但"保留接收者引用"那一格以前直接把裸控件名当 C 标识符交出去 ⇒ C2065。
Public Sub FocusIt()
    ucTest.SetFocus
End Sub
