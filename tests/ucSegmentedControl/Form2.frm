VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00F6F6F6&
   Caption         =   "Form2"
   ClientHeight    =   6645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10275
   LinkTopic       =   "Form2"
   ScaleHeight     =   6645
   ScaleWidth      =   10275
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   5775
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   9975
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl4 
      Height          =   495
      Left            =   6330
      TabIndex        =   3
      Top             =   120
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   873
      Border          =   -1  'True
      BorderColor     =   14079702
      BorderRadius    =   5
      SelectedBackColor=   5395026
      TabIndicatorColor=   0
      DividingLinesColor=   14079702
      MinItemWidth    =   0
      UseIconVector   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Symbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   2
      AutoSize        =   -1  'True
      Count           =   7
      image_0         =   "Form2.frx":0000
      HasImg_0        =   -1  'True
      image_1         =   "Form2.frx":00BE
      HasImg_1        =   -1  'True
      image_2         =   "Form2.frx":018A
      HasImg_2        =   -1  'True
      image_3         =   "Form2.frx":0255
      HasImg_3        =   -1  'True
      image_4         =   "Form2.frx":0321
      HasImg_4        =   -1  'True
      image_5         =   "Form2.frx":0411
      HasImg_5        =   -1  'True
      image_6         =   "Form2.frx":0598
      HasImg_6        =   -1  'True
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl3 
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   873
      Border          =   -1  'True
      BorderColor     =   14079702
      BorderRadius    =   5
      SelectedBackColor=   5395026
      TabIndicatorColor=   0
      DividingLinesColor=   14079702
      MinItemWidth    =   0
      UseIconVector   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Symbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoSize        =   -1  'True
      Count           =   4
      image_0         =   "Form2.frx":06C5
      HasImg_0        =   -1  'True
      image_1         =   "Form2.frx":0774
      HasImg_1        =   -1  'True
      image_2         =   "Form2.frx":0820
      HasImg_2        =   -1  'True
      image_3         =   "Form2.frx":08CF
      HasImg_3        =   -1  'True
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl2 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   873
      Border          =   -1  'True
      BorderColor     =   14079702
      BorderRadius    =   5
      SelectedBackColor=   5395026
      TabIndicatorColor=   0
      DividingLinesColor=   14079702
      MinItemWidth    =   0
      UseIconVector   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   2
      Count           =   1
      Caption_0       =   "Form2.frx":097A
      DropDown_0      =   -1  'True
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   873
      Border          =   -1  'True
      BorderColor     =   14079702
      BorderRadius    =   5
      SelectedBackColor=   5395026
      TabIndicatorColor=   0
      DividingLinesColor=   14079702
      MinItemWidth    =   0
      UseIconVector   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Symbol"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   1
      AutoSize        =   -1  'True
      Count           =   3
      image_0         =   "Form2.frx":09A6
      HasImg_0        =   -1  'True
      image_1         =   "Form2.frx":0AB9
      HasImg_1        =   -1  'True
      image_2         =   "Form2.frx":0BA8
      HasImg_2        =   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

