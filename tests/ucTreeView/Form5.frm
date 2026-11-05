VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   7215
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.ucTreeViewPlus ucTreeViewPlus1 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _extentx        =   5953
      _extenty        =   9340
      font            =   "Form5.frx":0000
      textcolor       =   0
      subtextfont     =   "Form5.frx":0028
      plusminuscolor  =   0
      plusminushotcolor=   16737123
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SwitchSize As Long
Private Sub Form_Load()
   
    With ucTreeViewPlus1
    
        SwitchSize = 32 * .GetWindowsDPI
    
        .TextColor = &H68635F
        .PlusMinusColor = &H68635F
        .CheckBoxes = True
        '.ExpandAllNodes
        .TreeLinesColor = &H55AA22
        .BorderColor = &H77DD77
        .BorderWidth = 4
        .NodeSpacing = 30
        .ToggleIconSize = 16
        .MinNodeHeight = 40
        .PlusMinusAlignRight = True
        .AlternateNodeColors = True
        '.FullRowSelect = False
        '.PlusMinusHotColor = vbBlue
        .LoadFromXML App.Path & "\menu.xml"
    End With

End Sub
Private Sub Form_Resize()
    ucTreeViewPlus1.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
End Sub

Private Sub ucTreeViewPlus1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lLeft As Long
    Dim Child As Node, Node As Node
    If Button <> vbLeftButton Then Exit Sub
    Set Node = ucTreeViewPlus1.HitTest(x, Y)
    If Not Node Is Nothing Then
        If Node.Text = "Extensions" Then
    
            lLeft = ucTreeViewPlus1.ClientWidth - SwitchSize * 2.5
        
            If x > lLeft And x < lLeft + SwitchSize Then
                Node.Checked = Not Node.Checked
                'n.Disabled = Not n.Checked
                ucTreeViewPlus1.Refresh
                For Each Child In Node.Childs
                    Child.Disabled = Not Node.Checked
                Next
            End If
        End If
    End If
End Sub

Private Sub ucTreeViewPlus1_NodePostPaint(Node As Node, ByVal hdc As Long, ByVal hGraphics As Long)
    If Node.Text = "Extensions" Then
        ucTreeViewPlus1.DrawSwitch hGraphics, IIf(Node.Checked, &HFF6363, &HCCCCCC), ucTreeViewPlus1.ClientWidth - SwitchSize * 2.5, Node.Top + (Node.Bottom - Node.Top - SwitchSize) / 2, SwitchSize, SwitchSize, Node.Checked
    End If
End Sub

