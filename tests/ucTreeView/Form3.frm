VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   LinkTopic       =   "Form3"
   ScaleHeight     =   7485
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucTreeViewPlus ucTreeViewPlus1 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _extentx        =   11456
      _extenty        =   12726
      font            =   "Form3.frx":0000
      textcolor       =   0
      selectioncolor  =   14610398
      fullrowselect   =   0   'False
      subtextcolor    =   9812594
      subtextfont     =   "Form3.frx":0028
      plusminuscolor  =   0
      plusminushotcolor=   16737123
      nodespacing     =   34
      marginbottom    =   10
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With ucTreeViewPlus1
        .TextColor = &H68635F
        .PlusMinusColor = &H68635F
        .ShowTreeLines = True
        .ExpandAllNodes
        .TreeLinesColor = &H55AA22
        .BorderColor = &H55AA22
        .BorderWidth = 4
        .LoadFromXML App.Path & "\JerarquiaEmpresa.xml"
        .ExpandAllNodes
        .Update
    End With
End Sub



