VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Tree Explorer"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
   LinkTopic       =   "Form2"
   ScaleHeight     =   7620
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   840
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Cursors folder"
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin Proyecto1.ucTreeViewPlus TVP 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12091
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      TextColor       =   0
      FullRowSelect   =   0   'False
      BeginProperty SubTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PlusMinusColor  =   2171169
      PlusMinusHotColor=   16737123
      CheckBoxColor   =   16771276
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cTE As ClsTreeExplorer
Attribute cTE.VB_VarHelpID = -1

Private Sub cTE_SelectFile(ByVal FilePath As String)
    Text1.Text = FilePath
End Sub

Private Sub cTE_SelectFolder(ByVal FolderPath As String)
    Debug.Print FolderPath
End Sub

Private Sub Form_Load()
    Set cTE = New ClsTreeExplorer
    cTE.Init TVP, "*.*" 'all files
    
    'cTE.Init TVP, "*.jpg;*.bmp;*.png;*.gif" 'only file image
    
    'cTE.Init TVP,"" 'only folder
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    TVP.Height = Me.ScaleHeight - TVP.Top * 2
End Sub

Private Sub Command1_Click()
   cTE.OpenPath "C:\Windows\Cursors"
   'Debug.Print TVP.SelNode.GetFullPath
   MsgBox cTE.GetNodePath(TVP.SelNode)
End Sub



