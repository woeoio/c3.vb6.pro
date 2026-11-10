VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HostForm"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Direct"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "DirectLabel"
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin PdParentSem.UCProbe ucP1 
      Height          =   600
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1058
   End
   Begin VB.Frame Frame1 
      Caption         =   "FrameOne"
      Height          =   1200
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   2400
      Begin VB.CommandButton Command1 
         Caption         =   "InFrame"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "FrameTwo"
      Height          =   1200
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   2400
      Begin PdParentSem.UCProbe ucP2 
         Height          =   600
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   1058
      End
   End
   Begin VB.PictureBox Pic1 
      Height          =   1200
      Left            =   6000
      ScaleHeight     =   1140
      ScaleWidth      =   2340
      TabIndex        =   5
      Top             =   1680
      Width           =   2400
      Begin VB.CommandButton Command3 
         Caption         =   "InPicture"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' PD-Parent 语义刀 —— 宿主侧那一层的 oracle 断言。
' 真 VB6 6.00.9782 实测: 控件的 Parent 一律是**顶层窗体** (Frame/PictureBox 里的也一样),
' 直接父容器那一问叫 Container。
Private Sub Form_Load()
    Debug.Print "PS-CMD2-PARENT-NAME=" & Command2.Parent.Name
    Debug.Print "PS-CMD2-PARENT-CAPTION=" & Command2.Parent.Caption
    Debug.Print "PS-CMD2-PARENT-TYPENAME=" & TypeName(Command2.Parent)
    Debug.Print "PS-LABEL2-PARENT-NAME=" & Label2.Parent.Name
    Debug.Print "PS-CMD1-PARENT-NAME=" & Command1.Parent.Name
    Debug.Print "PS-CMD1-PARENT-TYPENAME=" & TypeName(Command1.Parent)
    Debug.Print "PS-CMD1-CONTAINER-CAPTION=" & Command1.Container.Caption
    Debug.Print "PS-CMD3-PARENT-NAME=" & Command3.Parent.Name
    Debug.Print "PS-UCP1-PARENT-NAME=" & ucP1.Parent.Name
    Debug.Print "PS-UCP1-NAME=" & ucP1.Name
    Debug.Print "PS-UCP2-PARENT-NAME=" & ucP2.Parent.Name
    Debug.Print "PS-UCP2-CONTAINER-CAPTION=" & ucP2.Container.Caption
    ucP1.Dump "P1"
    ucP2.Dump "P2"
    Debug.Print "PS-DONE"
    Unload Me
End Sub
