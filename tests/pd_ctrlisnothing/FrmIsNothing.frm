VERSION 5.00
Begin VB.Form FrmIsNothing 
   Caption         =   "FrmIsNothing"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Begin VB.TextBox txtDummy 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmIsNothing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    If Not txtDummy Is Nothing Then txtDummy.Text = "x"
End Sub
