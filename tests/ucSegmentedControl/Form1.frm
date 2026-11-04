VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "<-- ItemVisible"
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   3360
      Width           =   1095
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl3 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   1560
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   873
      BackColor       =   16777215
      Border          =   -1  'True
      BorderRadius    =   40
      DividingLinesColor=   14079702
      MinItemWidth    =   0
      Padding         =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Count           =   3
      Caption_0       =   "Form1.frx":0000
      Value_0         =   -1  'True
      Caption_1       =   "Form1.frx":0032
      Caption_2       =   "Form1.frx":0066
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl12 
      Height          =   615
      Left            =   3360
      TabIndex        =   13
      Top             =   840
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   1085
      BackColor       =   16777215
      Border          =   -1  'True
      BorderRadius    =   5
      SelectedBackColor=   16777215
      SelectedForeColor=   15663310
      TabIndicatorColor=   15663310
      DividingLinesColor=   14079702
      DesignStyle     =   2
      MinItemWidth    =   0
      Padding         =   5
      UseIconVector   =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Count           =   3
      Caption_0       =   "Form1.frx":0094
      image_1         =   "Form1.frx":00BE
      HasImg_1        =   -1  'True
      Caption_1       =   "Form1.frx":0316
      Value_1         =   -1  'True
      Caption_2       =   "Form1.frx":0340
   End
   Begin SegmentedControl.ucSegmentedControl Pager 
      Height          =   615
      Left            =   1080
      TabIndex        =   12
      Top             =   6240
      Width           =   5430
      _ExtentX        =   9578
      _ExtentY        =   1085
      BackColor       =   16777215
      Border          =   -1  'True
      BorderRadius    =   40
      TabIndicatorColor=   46418
      DividingLinesColor=   14079702
      MinItemWidth    =   40
      Padding         =   5
      UseIconVector   =   -1  'True
      IconsSize       =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   2
      AutoSize        =   -1  'True
      Count           =   9
      image_0         =   "Form1.frx":036A
      HasImg_0        =   -1  'True
      Disabled_0      =   -1  'True
      image_1         =   "Form1.frx":0437
      HasImg_1        =   -1  'True
      Disabled_1      =   -1  'True
      Caption_2       =   "Form1.frx":04EF
      Value_2         =   -1  'True
      Caption_3       =   "Form1.frx":0511
      Caption_4       =   "Form1.frx":0533
      Caption_5       =   "Form1.frx":0555
      Caption_6       =   "Form1.frx":0577
      image_7         =   "Form1.frx":0599
      HasImg_7        =   -1  'True
      image_8         =   "Form1.frx":0643
      HasImg_8        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CtrlType = 1 - CheckBar"
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   4200
      Width           =   2655
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl11 
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   6480
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BackColor       =   16777215
      Border          =   -1  'True
      DividingLinesVisible=   0   'False
      TabIndicatorColor=   0
      DesignStyle     =   1
      MinItemWidth    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Count           =   2
      Caption_0       =   "Form1.frx":0708
      Value_0         =   -1  'True
      Caption_1       =   "Form1.frx":0732
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl10 
      Height          =   495
      Left            =   6960
      TabIndex        =   9
      Top             =   6480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BackColor       =   16777215
      Border          =   -1  'True
      DividingLinesVisible=   0   'False
      TabIndicatorColor=   0
      DesignStyle     =   1
      MinItemWidth    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Count           =   2
      Caption_0       =   "Form1.frx":0768
      Value_0         =   -1  'True
      Caption_1       =   "Form1.frx":0790
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl9 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      BackColor       =   16758271
      Border          =   -1  'True
      BorderColor     =   16739327
      BorderWidth     =   2
      DividingLinesVisible=   0   'False
      SelectedBackColor=   16711935
      TabIndicatorColor=   0
      MinItemWidth    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483634
      Gradient        =   -1  'True
      CtrlType        =   2
      Count           =   1
      Caption_0       =   "Form1.frx":07C0
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl8 
      Height          =   615
      Left            =   3240
      TabIndex        =   7
      Top             =   4800
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1085
      BackColor       =   8222060
      SelectedBackColor=   6840922
      DividingLinesColor=   8222060
      MinItemWidth    =   0
      Padding         =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483643
      CtrlType        =   2
      Count           =   1
      Caption_0       =   "Form1.frx":07F6
      DropDown_0      =   -1  'True
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl7 
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   5520
      Width           =   2895
      _ExtentX        =   4895
      _ExtentY        =   1085
      BackColor       =   8222060
      SelectedBackColor=   6840922
      DividingLinesColor=   8222060
      MinItemWidth    =   0
      Padding         =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483643
      CtrlType        =   2
      Count           =   3
      Caption_0       =   "Form1.frx":081E
      Caption_1       =   "Form1.frx":0846
      Caption_2       =   "Form1.frx":0870
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl6 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      BackColor       =   16777215
      Border          =   -1  'True
      BorderRadius    =   40
      DividingLinesColor=   14079702
      MinItemWidth    =   0
      Padding         =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Symbol"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   1
      Count           =   4
      Caption_0       =   "Form1.frx":089A
      Caption_1       =   "Form1.frx":08BC
      Value_1         =   -1  'True
      Caption_2       =   "Form1.frx":08DE
      Caption_3       =   "Form1.frx":0900
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl5 
      Height          =   3135
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
      BackColor       =   16777215
      Border          =   -1  'True
      Orientation     =   1
      MinItemWidth    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   1
      Count           =   4
      image_0         =   "Form1.frx":0922
      HasImg_0        =   -1  'True
      Caption_0       =   "Form1.frx":0B2C
      image_1         =   "Form1.frx":0B5C
      HasImg_1        =   -1  'True
      Caption_1       =   "Form1.frx":0D18
      Value_1         =   -1  'True
      image_2         =   "Form1.frx":0D48
      HasImg_2        =   -1  'True
      Caption_2       =   "Form1.frx":367A
      image_3         =   "Form1.frx":36AA
      HasImg_3        =   -1  'True
      Caption_3       =   "Form1.frx":C042
      Value_3         =   -1  'True
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl4 
      Height          =   3975
      Left            =   7680
      TabIndex        =   3
      Top             =   2160
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   7011
      BackColor       =   16777215
      Border          =   -1  'True
      DividingLinesVisible=   0   'False
      Orientation     =   1
      DesignStyle     =   1
      MinItemWidth    =   0
      Padding         =   5
      IconsAlignTop   =   -1  'True
      IconsSize       =   32
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Count           =   4
      image_0         =   "Form1.frx":C072
      HasImg_0        =   -1  'True
      Caption_0       =   "Form1.frx":C27C
      image_1         =   "Form1.frx":C2A8
      HasImg_1        =   -1  'True
      Caption_1       =   "Form1.frx":C464
      image_2         =   "Form1.frx":C488
      HasImg_2        =   -1  'True
      Caption_2       =   "Form1.frx":EDBA
      Value_2         =   -1  'True
      image_3         =   "Form1.frx":EDE6
      HasImg_3        =   -1  'True
      Caption_3       =   "Form1.frx":18C91
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl2 
      Height          =   3135
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   5530
      BackColor       =   16777215
      Border          =   -1  'True
      BorderRadius    =   0
      DividingLinesVisible=   0   'False
      SelectedBackColor=   13027014
      TabIndicatorColor=   5944576
      Orientation     =   1
      DesignStyle     =   2
      MinItemWidth    =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   2
      Count           =   4
      image_0         =   "Form1.frx":18CBD
      HasImg_0        =   -1  'True
      Caption_0       =   "Form1.frx":18EC7
      image_1         =   "Form1.frx":18EF7
      HasImg_1        =   -1  'True
      Caption_1       =   "Form1.frx":190B3
      DropDown_1      =   -1  'True
      image_2         =   "Form1.frx":190E3
      HasImg_2        =   -1  'True
      Caption_2       =   "Form1.frx":1BA15
      image_3         =   "Form1.frx":1BA45
      HasImg_3        =   -1  'True
      Caption_3       =   "Form1.frx":258F0
   End
   Begin SegmentedControl.ucSegmentedControl ucSegmentedControl1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9135
      _ExtentX        =   12515
      _ExtentY        =   1296
      TabIndicatorColor=   35840
      DesignStyle     =   1
      MinItemWidth    =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CtrlType        =   2
      Count           =   5
      image_0         =   "Form1.frx":25920
      HasImg_0        =   -1  'True
      Caption_0       =   "Form1.frx":25AE9
      image_1         =   "Form1.frx":25B19
      HasImg_1        =   -1  'True
      Caption_1       =   "Form1.frx":25D6D
      image_2         =   "Form1.frx":25D9D
      HasImg_2        =   -1  'True
      Caption_2       =   "Form1.frx":260AF
      Disabled_2      =   -1  'True
      image_3         =   "Form1.frx":260DF
      HasImg_3        =   -1  'True
      Caption_3       =   "Form1.frx":2624D
      DropDown_3      =   -1  'True
      image_4         =   "Form1.frx":2627B
      HasImg_4        =   -1  'True
      Caption_4       =   "Form1.frx":26650
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Boolean




Private Sub Command1_Click()
ucSegmentedControl4.ItemVisible(1) = Not ucSegmentedControl4.ItemVisible(1)
End Sub


Private Sub Form_Resize()
    ucSegmentedControl1.Move 0, 0, Me.ScaleWidth
End Sub

Private Sub Pager_ItemClick(Index As Long)
    Const Min = 1
    Const Max = 50
    Dim OldPage As Long
    Dim NewPage As Long
    Dim i As Long
    Dim lStar As Long
    Dim lEnd As Long
    
    For i = 2 To 6
        If Pager.ItemValue(i) Then
            OldPage = Pager.ItemCaption(i)
        End If
        Pager.ItemValue(i) = False
    Next
    
   Select Case Index
        Case 0: NewPage = Min
        Case 1: NewPage = OldPage - 1
        Case 7: NewPage = OldPage + 1
        Case 8: NewPage = Max
        Case Else: NewPage = Pager.ItemCaption(Index)
   End Select

   If NewPage < Min Then NewPage = Min
   If NewPage > Max Then NewPage = Max
      
   Pager.ItemDisabled(0) = NewPage = Min
   Pager.ItemDisabled(1) = NewPage = Min
   Pager.ItemDisabled(7) = NewPage = Max
   Pager.ItemDisabled(8) = NewPage = Max
   
   If NewPage - 2 < Min Then
        lStar = Min
        lEnd = Min + 4
    ElseIf NewPage + 2 > Max Then
        lStar = Max - 4
        lEnd = Max
    Else
        If Index > 2 And Index < 7 Then
            NewPage = Pager.ItemCaption(Index)
        End If
        lStar = NewPage - 2
        lEnd = NewPage + 2
    End If
    
    Dim n As Long
    n = 2
    For i = lStar To lEnd
        Pager.ItemCaption(n) = i
        If NewPage = i Then
            Pager.ItemValue(n) = True
        End If
        n = n + 1
    Next
    
    Debug.Print NewPage
End Sub

Private Sub ucSegmentedControl10_ItemClick(Index As Long)
    Dim Ctrl As Object
    
    For Each Ctrl In Controls
        If TypeName(Ctrl) = "ucSegmentedControl" Then
            Ctrl.Gradient = Index = 1
        End If
    Next
End Sub

Private Sub ucSegmentedControl11_ItemClick(Index As Long)
    Dim Ctrl As Object
    If Index = 1 Then
        'Me.Picture = LoadPicture(App.Path & "\image.jpg")
        Me.AutoRedraw = True
        Me.PaintPicture LoadPicture(App.Path & "\image.jpg"), 0, 0, Me.ScaleWidth, Me.ScaleHeight
        Me.Picture = Me.Image
    Else
        Me.Picture = Nothing
    End If
    
    For Each Ctrl In Controls
        If TypeName(Ctrl) = "ucSegmentedControl" Then
            Ctrl.BackColorOpacity = IIF(Index = 0, 100, 60)
        End If
    Next
End Sub

Private Sub ucSegmentedControl2_ItemClick(Index As Long)
    If Index = 1 Then
        ucSegmentedControl5.Visible = Not ucSegmentedControl5.Visible
        ucSegmentedControl2.ItemValue(Index) = Not ucSegmentedControl2.ItemValue(Index)
    Else
        ucSegmentedControl5.Visible = False
        ucSegmentedControl2.ItemValue(1) = False
    End If
End Sub

Private Sub ucSegmentedControl5_MouseLeave()
        ucSegmentedControl5.Visible = False
        ucSegmentedControl2.ItemValue(1) = False
End Sub

'Private Sub ucSegmentedControl2_ItemMouseEnter(Index As Long)
'  If Index = 1 Then
'     ucSegmentedControl5.Visible = True
'  End If
'End Sub


'Private Sub ucSegmentedControl2_ItemMouseLeave(Index As Long)
'  If Index = 1 Then
'
'   If n = False Then ucSegmentedControl5.Visible = False
'  End If
'End Sub

'Private Sub ucSegmentedControl5_MouseEnter()
'  n = True
'End Sub
'
'Private Sub ucSegmentedControl5_MouseLeave()
'   ucSegmentedControl5.Visible = False
'End Sub
Private Sub ucSegmentedControl7_ItemClick(Index As Long)

End Sub

Private Sub ucSegmentedControl9_ItemClick(Index As Long)
    Form2.Show , Me
End Sub
