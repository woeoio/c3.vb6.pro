VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6915
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   11295
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "AllowEdit"
      Height          =   255
      Index           =   7
      Left            =   3120
      TabIndex        =   74
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Contact  List"
      Height          =   615
      Index           =   4
      Left            =   7320
      TabIndex        =   73
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Menu / Tools"
      Height          =   615
      Index           =   3
      Left            =   6240
      TabIndex        =   72
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Drag and Drop"
      Height          =   615
      Index           =   2
      Left            =   5160
      TabIndex        =   71
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Corporate Hierarchy"
      Height          =   615
      Index           =   1
      Left            =   4080
      TabIndex        =   70
      Top             =   6120
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tree Explorer"
      Height          =   615
      Index           =   0
      Left            =   3000
      TabIndex        =   69
      Top             =   6120
      Width           =   1035
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   5
      Left            =   3120
      ScaleHeight     =   405
      ScaleWidth      =   2895
      TabIndex        =   65
      Top             =   5520
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Index           =   5
         Left            =   2250
         Max             =   20
         TabIndex        =   66
         Top             =   10
         Value           =   20
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   400
         Index           =   5
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   67
         Text            =   "0"
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "MarginBottom"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   68
         Top             =   80
         Width           =   1455
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "PlusMinusAlignRight"
      Height          =   255
      Index           =   6
      Left            =   3120
      TabIndex        =   64
      Top             =   2280
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   60
      Top             =   4080
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFE8CC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   62
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   2160
         TabIndex        =   61
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "SelectionColor"
         Height          =   255
         Index           =   8
         Left            =   480
         TabIndex        =   63
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "PlusMinusStyle = 1 - TVP_PlusMinus"
      Height          =   255
      Index           =   5
      Left            =   3120
      TabIndex        =   59
      Top             =   1920
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   4
      Left            =   3120
      ScaleHeight     =   405
      ScaleWidth      =   2895
      TabIndex        =   55
      Top             =   5040
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Index           =   4
         Left            =   2250
         Max             =   20
         Min             =   4
         TabIndex        =   56
         Top             =   10
         Value           =   14
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   400
         Index           =   4
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "10"
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "ToggleIconSize"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   58
         Top             =   80
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   3
      Left            =   3120
      ScaleHeight     =   405
      ScaleWidth      =   2895
      TabIndex        =   51
      Top             =   4560
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Index           =   3
         Left            =   2265
         Max             =   50
         TabIndex        =   52
         Top             =   10
         Value           =   34
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   400
         Index           =   3
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "16"
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "NodeSpacing"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   54
         Top             =   80
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   2
      Left            =   3120
      ScaleHeight     =   405
      ScaleWidth      =   2895
      TabIndex        =   47
      Top             =   4080
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Index           =   2
         Left            =   2250
         Max             =   20
         TabIndex        =   48
         Top             =   10
         Value           =   15
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   400
         Index           =   2
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "5"
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Margin"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   50
         Top             =   80
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   1
      Left            =   3120
      ScaleHeight     =   405
      ScaleWidth      =   2895
      TabIndex        =   43
      Top             =   3600
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Index           =   1
         Left            =   2265
         Max             =   5
         TabIndex        =   44
         Top             =   10
         Value           =   4
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   400
         Index           =   1
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "1"
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "BorderWidth"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   46
         Top             =   80
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   0
      Left            =   3120
      ScaleHeight     =   405
      ScaleWidth      =   2895
      TabIndex        =   39
      Top             =   3120
      Width           =   2895
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Index           =   0
         Left            =   2250
         Max             =   15
         TabIndex        =   41
         Top             =   10
         Value           =   10
         Width           =   390
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   0
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "5"
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "BorderRadius"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   42
         Top             =   80
         Width           =   1455
      End
   End
   Begin Proyecto1.ucTreeViewPlus ucTreeViewPlus1 
      Height          =   5415
      Left            =   120
      TabIndex        =   38
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8493
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   0
      CheckBoxes      =   -1  'True
      ShowTreeLines   =   -1  'True
      BeginProperty SubTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PlusMinusColor  =   0
      PlusMinusHotColor=   16737123
      MinNodeHeight   =   24
      MousePointerHands=   -1  'True
   End
   Begin VB.CheckBox Check1 
      Caption         =   "FullRowSelect"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   37
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   33
      Top             =   3600
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00CCCCCC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   35
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   2160
         TabIndex        =   34
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "TreeLinesColor"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   36
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   29
      Top             =   3120
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   2160
         TabIndex        =   31
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   30
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "TextColor"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   32
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   25
      Top             =   2640
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00CCCCCC&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   27
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2160
         TabIndex        =   26
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "SubTextColor"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   28
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   21
      Top             =   2160
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   23
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF6363&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   22
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "PlusMinusHotColor"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   24
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   17
      Top             =   1680
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   19
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   18
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "PlusMinusColor"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   20
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   13
      Top             =   1200
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   15
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FF6363&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "CheckBoxColor"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   16
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   9
      Top             =   720
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000A&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   11
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "BorderColor"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   6720
      ScaleHeight     =   375
      ScaleWidth      =   2775
      TabIndex        =   5
      Top             =   240
      Width           =   2775
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   0
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "BackColor"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   60
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load from custom XML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ShowPlusMinus"
      Height          =   255
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "ShowTreeLines"
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CheckBoxes"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "AlternateNodeColors"
      Height          =   255
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type


Private Function ShowColorDialog(lColor As Long) As Boolean
    Dim cc As CHOOSECOLOR
    Dim lReturn As Long
    Dim aCustomColors(15) As Long
    
    If (lColor And &H80000000) Then lColor = GetSysColor(lColor And &HFF&)
    
    ' Set the size of the structure
    cc.lStructSize = Len(cc)
    cc.hwndOwner = Me.hwnd
    cc.flags = &H1 Or &H2 Or &H8
    cc.lpCustColors = VarPtr(aCustomColors(0))
    cc.rgbResult = lColor ' Initial color
    
    ' Call the ChooseColor function
    lReturn = CHOOSECOLOR(cc)
    
    ' Check if the user clicked OK
    If lReturn <> 0 Then
        ' Get the selected color
        ShowColorDialog = True
        lColor = cc.rgbResult
    End If
End Function

Private Sub Check1_Click(index As Integer)
    Select Case index
        Case 0
            ucTreeViewPlus1.AlternateNodeColors = Check1(index).Value
        Case 1
            ucTreeViewPlus1.CheckBoxes = Check1(index).Value
        Case 2
            ucTreeViewPlus1.ShowTreeLines = Check1(index).Value
        Case 3
            ucTreeViewPlus1.PlusMinusVisible = Check1(index).Value
        Case 4
            ucTreeViewPlus1.FullRowSelect = Check1(index).Value
        Case 5
            ucTreeViewPlus1.PlusMinusStyle = IIf(Check1(index).Value, TVP_PlusMinus, TVP_Arrow)
        Case 6
            ucTreeViewPlus1.PlusMinusAlignRight = Check1(index).Value
        Case 7
            ucTreeViewPlus1.AllowEdit = Check1(index).Value
    End Select
End Sub

Private Sub Command1_Click()
    ' Crear un nuevo objeto XMLDocument
    Dim xmlDoc As Object
    
    ' Crear el objeto XMLDocument
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    ' Cargar el archivo XML
    xmlDoc.Load App.Path & "\ejemplo.xml"
    
    ' Verificar si hubo errores al cargar el XML
    If xmlDoc.parseError.errorCode <> 0 Then
        MsgBox "Error al cargar el archivo XML: " & xmlDoc.parseError.reason
        Exit Sub
    End If
    
    ucTreeViewPlus1.Nodes.Clear
    ucTreeViewPlus1.CheckBoxes = False

    ' Convertir el nodo raíz del XML en un nodo de tu clase
    ConvertXMLNodeToNode xmlDoc.documentElement, ucTreeViewPlus1.Nodes

    ucTreeViewPlus1.Update
End Sub

' Función recursiva para convertir un nodo XML en un nodo de la clase Node
Private Sub ConvertXMLNodeToNode(ByVal xmlNode As Object, ByVal parentNode As Node)
    Dim childNode As Object
    Dim newNode As Node
    Dim AttributeValue
    ' Crear un nuevo nodo de tu clase
    Set newNode = parentNode.Add(xmlNode.nodename)
    
    
    AttributeValue = xmlNode.getAttribute("name")
    
    If Not IsNull(AttributeValue) Then
       newNode.SubText = AttributeValue
    End If

    ' Recorrer los nodos hijos del nodo XML
    For Each childNode In xmlNode.childNodes
        If childNode.nodeType = 1 Then ' Solo procesar nodos de tipo Elemento
            ConvertXMLNodeToNode childNode, newNode
        Else
            newNode.SubText = xmlNode.Text
        End If
    Next childNode
End Sub


Private Sub Command2_Click(index As Integer)
    Dim lColor As Long
    
    lColor = Picture2(index).BackColor
    
    If ShowColorDialog(lColor) = False Then
         Exit Sub
    End If
    
    Picture2(index).BackColor = lColor
    
    Select Case index
    
        Case 0
            ucTreeViewPlus1.BackColor = lColor
        Case 1
            ucTreeViewPlus1.BorderColor = lColor
        Case 2
            ucTreeViewPlus1.CheckBoxColor = lColor
        Case 3
            ucTreeViewPlus1.PlusMinusColor = lColor
        Case 4
            ucTreeViewPlus1.PlusMinusHotColor = lColor
        Case 5
            ucTreeViewPlus1.SubTextColor = lColor
        Case 6
            ucTreeViewPlus1.TextColor = lColor
        Case 7
            ucTreeViewPlus1.TreeLinesColor = lColor
        Case 8
            ucTreeViewPlus1.SelectionColor = lColor
    End Select
End Sub



Private Sub Command7_Click()

End Sub

Private Sub Command3_Click(index As Integer)
    Select Case index
        Case 0: Form2.Show
        Case 1: Form3.Show
        Case 2: Form4.Show
        Case 3: Form5.Show
        Case 4: Form6.Show
    End Select
End Sub

Private Sub Form_Load()
    With ucTreeViewPlus1
        .ImageListInitialize 16, 16
        .ImageListAddImage App.Path & "\img\facebook.png"
        .ImageListAddImage App.Path & "\img\instagram.png"
        .ImageListAddImage App.Path & "\img\whatsapp.png"
        .ImageListAddImage App.Path & "\img\youtube.png"
        .ImageListAddImage App.Path & "\img\mensajero.png"
        .ImageListAddImage App.Path & "\img\tik-tok.png"
    End With

    With ucTreeViewPlus1.Nodes.Add("Facebook")
        .SubText = "Meta"
        .Image = 1
        .RadioStyle = True
        .BackColor = RGB(255, 180, 180)
        .TextColor = vbWhite
        '.HideCheck = True
        .Add("Mostrar").Checked = True
        .Add("Ocultar").Image = 1
        .Add("Disernir").Disabled = True
    End With
    
    With ucTreeViewPlus1.Nodes.Add("Instagram")
        .SubText = "Meta"
        .Image = 2
        .RadioStyle = True
        .HideCheck = True
        .Add ("Mostrar")
        .Add("Ocultar").Checked = True
        .Add ("Disernir")
    End With
    
    With ucTreeViewPlus1.Nodes.Add("WhatsApp")
        .SubText = "Meta"
        .Image = 3
        .RadioStyle = True
        .HideCheck = True
        .Add ("Mostrar")
        .Add("Ocultar").Checked = True
        .Add ("Disernir")
    End With

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    ucTreeViewPlus1.Height = Me.ScaleHeight - ucTreeViewPlus1.Top * 2
End Sub

Private Sub Text1_Change(index As Integer)
    Select Case index
    
        Case 0
            ucTreeViewPlus1.BorderRadius = Val(Text1(index).Text)
        Case 1
            ucTreeViewPlus1.BorderWidth = Val(Text1(index).Text)
        Case 2
            ucTreeViewPlus1.Margin = Val(Text1(index).Text)
        Case 3
            ucTreeViewPlus1.NodeSpacing = Val(Text1(index).Text)
        Case 4
            ucTreeViewPlus1.ToggleIconSize = Val(Text1(index).Text)
        Case 5
            ucTreeViewPlus1.MarginBottom = Val(Text1(index).Text)
    End Select
End Sub

'Private Sub TreeView1_AfterLabelEdit(Cancel As Integer, NewString As String)
'Cancel = 1
'End Sub

Private Sub ucTreeViewPlus1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 And ucTreeViewPlus1.AllowEdit Then
        If Not ucTreeViewPlus1.SelNode Is Nothing Then
            ucTreeViewPlus1.StartEditNode ucTreeViewPlus1.SelNode
        End If
    End If
End Sub

Private Sub ucTreeViewPlus1_NodeDblClick(Node As Node)
    'MsgBox Node.Text
End Sub

Private Sub VScroll1_Change(index As Integer)
    Text1(index).Text = VScroll1(index).Max - VScroll1(index).Value + VScroll1(index).Min
End Sub

Private Sub ucTreeViewPlus1_NodePostPaint(Node As Node, ByVal hdc As Long, ByVal hGraphics As Long)
     Dim dpif As Single
     dpif = ucTreeViewPlus1.GetWindowsDPI
     'JImageListEx1.Render 0, hGraphics, 0, (ucTreeViewPlus1.Width / 15 - 40) * dpif, Node.Top + ((Node.Bottom - Node.Top) / 2 - 8 * dpif), 16 * dpif, 16 * dpif
End Sub
