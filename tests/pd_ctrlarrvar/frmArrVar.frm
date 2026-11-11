VERSION 5.00
Begin VB.Form frmArrVar 
   Caption         =   "frmArrVar"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Begin PdArrVar.pdButtonToolbox cmdFlyoutLock 
      Height          =   390
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   390
   End
   Begin PdArrVar.pdButtonToolbox cmdFlyoutLock 
      Height          =   390
      Index           =   1
      Left            =   600
      Top             =   120
      Width           =   390
   End
End
Attribute VB_Name = "frmArrVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ArrVar.ThemeFlyoutControls cmdFlyoutLock
End Sub
