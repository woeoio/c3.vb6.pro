VERSION 5.00
Begin VB.MDIForm MaxMdi 
   BackColor       =   &H8000000C&
   Caption         =   "MAX-MDI"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MaxMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Dim c As MaxChild
    Set c = New MaxChild
    c.Show
End Sub
