VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form Form1 
   Caption         =   "Chart appearance"
   ClientHeight    =   7500
   ClientLeft      =   1530
   ClientTop       =   570
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7500
   ScaleWidth      =   7470
   Begin TeeChart.TChart TChart1 
      Height          =   3855
      Left            =   5
      OleObjectBlob   =   "DisplayForm.frx":0000
      TabIndex        =   0
      Top             =   60
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Panel properties"
      Height          =   3495
      Left            =   0
      TabIndex        =   24
      Top             =   3960
      Width           =   6015
      Begin VB.CommandButton Command6 
         Caption         =   "Panel colour"
         Height          =   375
         Left            =   2880
         TabIndex        =   59
         Top             =   420
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   4320
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   420
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Back colour"
         Height          =   375
         Left            =   2880
         TabIndex        =   57
         Top             =   900
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   4320
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   900
         Width           =   375
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Border"
         Height          =   255
         Left            =   4800
         TabIndex        =   55
         Top             =   480
         Width           =   855
      End
      Begin VB.CheckBox Check4 
         Caption         =   "No Back colour"
         Height          =   375
         Left            =   4800
         TabIndex        =   54
         Top             =   900
         Width           =   1095
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   315
         Left            =   1440
         Max             =   4
         TabIndex        =   53
         Top             =   1440
         Width           =   315
      End
      Begin VB.VScrollBar VScroll2 
         Height          =   315
         Left            =   120
         Max             =   4
         Min             =   1
         TabIndex        =   52
         Top             =   1440
         Value           =   1
         Width           =   315
      End
      Begin VB.Frame Frame11 
         Caption         =   "Bevel inner"
         Height          =   1035
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1095
         Begin VB.OptionButton Option10 
            Caption         =   "Raised"
            Height          =   255
            Left            =   90
            TabIndex        =   51
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Lowered"
            Height          =   255
            Left            =   90
            TabIndex        =   50
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option8 
            Caption         =   "None"
            Height          =   255
            Left            =   90
            TabIndex        =   49
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bevel outer"
         Height          =   1035
         Left            =   1440
         TabIndex        =   44
         Top             =   360
         Width           =   1095
         Begin VB.OptionButton Option13 
            Caption         =   "Raised"
            Height          =   255
            Left            =   90
            TabIndex        =   47
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Lowered"
            Height          =   255
            Left            =   90
            TabIndex        =   46
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Option11 
            Caption         =   "None"
            Height          =   255
            Left            =   90
            TabIndex        =   45
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Back Image"
         Height          =   1455
         Left            =   120
         TabIndex        =   36
         Top             =   1980
         Width           =   2655
         Begin VB.Frame Frame9 
            Caption         =   "Style"
            Height          =   1215
            Left            =   1440
            TabIndex        =   40
            Top             =   180
            Width           =   1095
            Begin VB.OptionButton Option5 
               Caption         =   "Stretch"
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   855
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Tile"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   480
               Width           =   855
            End
            Begin VB.OptionButton Option7 
               Caption         =   "Centre"
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   720
               Width           =   855
            End
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Inside"
            Height          =   195
            Left            =   180
            TabIndex        =   39
            Top             =   1140
            Width           =   975
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Clear"
            Height          =   255
            Left            =   180
            TabIndex        =   38
            Top             =   780
            Width           =   975
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Browse..."
            Height          =   255
            Left            =   180
            TabIndex        =   37
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Gradient"
         Height          =   2055
         Left            =   2880
         TabIndex        =   25
         Top             =   1380
         Width           =   3015
         Begin VB.Frame Frame10 
            Caption         =   "Direction"
            Height          =   855
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   2775
            Begin VB.OptionButton Option1 
               Caption         =   "Right Left"
               Height          =   255
               Left            =   1440
               TabIndex        =   35
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Left Right"
               Height          =   255
               Left            =   1440
               TabIndex        =   34
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Bottom Top"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               Top             =   480
               Width           =   1275
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Top Bottom"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Visible"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox Text3 
            BackColor       =   &H0000FFFF&
            Height          =   375
            Left            =   1200
            TabIndex        =   29
            Top             =   700
            Width           =   375
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1200
            TabIndex        =   28
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command9 
            Caption         =   "End colour"
            Height          =   375
            Left            =   1680
            TabIndex        =   27
            Top             =   700
            Width           =   1215
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Start colour"
            Height          =   375
            Left            =   1680
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         Height          =   270
         Left            =   540
         TabIndex        =   61
         Top             =   1500
         Width           =   465
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Offset"
         Height          =   255
         Left            =   1860
         TabIndex        =   60
         Top             =   1500
         Width           =   555
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Margins"
      Height          =   2055
      Left            =   6120
      TabIndex        =   18
      Top             =   5400
      Width           =   1335
      Begin VB.HScrollBar HScroll3 
         Height          =   375
         Left            =   120
         Max             =   35
         Min             =   3
         TabIndex        =   23
         Top             =   660
         Value           =   3
         Width           =   495
      End
      Begin VB.HScrollBar HScroll4 
         Height          =   375
         Left            =   720
         Max             =   35
         Min             =   3
         TabIndex        =   22
         Top             =   660
         Value           =   5
         Width           =   495
      End
      Begin VB.VScrollBar VScroll3 
         Height          =   375
         Left            =   480
         Max             =   35
         Min             =   3
         TabIndex        =   21
         Top             =   240
         Value           =   3
         Width           =   375
      End
      Begin VB.VScrollBar VScroll4 
         Height          =   375
         Left            =   480
         Max             =   35
         Min             =   3
         TabIndex        =   20
         Top             =   1080
         Value           =   3
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Reset"
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   1635
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "3D"
      Height          =   1335
      Left            =   6120
      TabIndex        =   13
      Top             =   3960
      Width           =   1335
      Begin VB.CheckBox Check5 
         Caption         =   "3D Enable"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   285
         Width           =   1095
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   120
         Max             =   100
         Min             =   1
         TabIndex        =   14
         Top             =   645
         Value           =   1
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "3D %:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   1005
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Walls"
      Height          =   1455
      Left            =   6120
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
      Begin VB.CheckBox Check11 
         Caption         =   "Visible"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   285
         Width           =   855
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         Left            =   120
         Max             =   10
         TabIndex        =   9
         Top             =   765
         Value           =   1
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Size"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label Label11 
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   1125
         Width           =   375
      End
   End
   Begin VB.Frame AxesFrame 
      Caption         =   "Chart axes"
      Height          =   2295
      Left            =   6120
      TabIndex        =   1
      Top             =   0
      Width           =   1335
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1095
         Begin VB.CheckBox Check10 
            Caption         =   "Right"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Bottom"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Left"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Top"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Frame"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   1725
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog3 
      Left            =   6885
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   6885
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6405
      Top             =   4140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
' enable / disable gradient options...
  With TChart1.Panel.Gradient
    .Visible = Check1.Value
    Option1.Enabled = .Visible
    Option2.Enabled = .Visible
    Option3.Enabled = .Visible
    Option4.Enabled = .Visible
    Option3.Value = .Visible
  End With
End Sub

Private Sub Check10_Click()
' show / hide Right axis...
  TChart1.Axis.Right.Visible = Check10.Value
End Sub

Private Sub Check11_Click()
' show / hide Walls...
  TChart1.Walls.Visible = Check11.Value
End Sub

Private Sub Check2_Click()
' show BackImage inside axis or inside whole Chart background...
  TChart1.Panel.BackImageInside = Check2.Value
End Sub


Private Sub Check3_Click()
' show / hide Chart border...
  With TChart1.Panel
    If Check3.Value = 1 Then
      .BorderStyle = bsSingle
    Else
      .BorderStyle = bsNone
    End If
  End With
End Sub


Private Sub Check4_Click()
' change Chart back wall color...
  With TChart1.Walls
    If Check4.Value = 1 Then
      .Back.Transparent = True
    Else
      .Back.Transparent = False
      .Back.Color = Text4.BackColor
    End If
  End With
End Sub


Private Sub Check5_Click()
' switch 2D / 3D ...
  TChart1.Aspect.View3D = Check5.Value
  HScroll1.Enabled = Check5.Value
End Sub

Private Sub Command1_Click()
    HScroll3.Value = 3
    VScroll3.Value = 3
    HScroll4.Value = 32
    VScroll4.Value = 32
End Sub


Private Sub Check6_Click()
  
  'On top of the Frame are the TChart Axes
  'Therefore you won't see any change unless
  'Axis visible is 'false' on one of the Axes
  
  TChart1.Frame.Visible = Check6.Value
End Sub


Private Sub Check7_Click()
' show / hide Top axis...
 TChart1.Axis.Top.Visible = Check7.Value
End Sub

Private Sub Check8_Click()
' show / hide Left axis...
 TChart1.Axis.Left.Visible = Check8.Value
End Sub


Private Sub Check9_Click()
' show / hide Bottom axis...
 TChart1.Axis.Bottom.Visible = Check9.Value
End Sub

Private Sub Command10_Click()
' clear background bitmap image...
  TChart1.Panel.BackImageClear
End Sub

Private Sub Command11_Click()
' change Chart Walls back color...
 CommonDialog1.ShowColor
 If CommonDialog1.Color > 0 Then
   TChart1.Walls.BackColor = CommonDialog1.Color
   Text4.BackColor = CommonDialog1.Color
 End If
End Sub


Private Sub Command2_Click()
' show / hide Top axis...
  With TChart1.Axis.Top
    .Visible = Not .Visible
  End With
End Sub


Private Sub Command3_Click()
' show / hide Left axis...
  With TChart1.Axis.Left
    .Visible = Not .Visible
  End With
End Sub


Private Sub Command4_Click()
' show / hide Left wall...
  With TChart1.LeftWall
    .Visible = Not .Visible
  End With
End Sub


Private Sub Command5_Click()
' show / hide Bottom wall...
  With TChart1.BottomWall
    .Visible = Not .Visible
  End With
End Sub


Private Sub Command6_Click()
' change Chart background color...
 CommonDialog1.ShowColor
 If CommonDialog1.Color > 0 Then
   TChart1.Panel.Color = CommonDialog1.Color
   Text1.BackColor = CommonDialog1.Color
 End If
End Sub


Private Sub Command7_Click()
' load an image to fill Chart background...
  CommonDialog1.Filter = "Pictures(*.bmp,*.wmf,*.jpg)|*.bmp;*.wmf;*.jpg"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
    TChart1.Panel.BackImageLoad (CommonDialog1.FileName)
    Option5.Enabled = True
    Option6.Enabled = True
    Option7.Enabled = True
    Option5.Value = True
    Check2.Enabled = True
  End If
End Sub


Private Sub Command8_Click()
' change background Gradient start color...
 CommonDialog2.ShowColor
 If CommonDialog2.Color > 0 Then
   TChart1.Panel.Gradient.StartColor = CommonDialog2.Color
   Text2.BackColor = CommonDialog2.Color
 End If
End Sub


Private Sub Command9_Click()
' change background Gradient end color...
 CommonDialog3.ShowColor
 If CommonDialog3.Color > 0 Then
   TChart1.Panel.Gradient.EndColor = CommonDialog3.Color
   Text3.BackColor = CommonDialog3.Color
 End If
End Sub


Private Sub Form_Load()
    Option1.Enabled = False
    Option2.Enabled = False
    Option3.Enabled = False
    Option4.Enabled = False
    Option5.Enabled = False
    Option6.Enabled = False
    Option7.Enabled = False
    'BackImage Inside
    Check2.Value = False
    Check2.Enabled = False
    Check5.Value = 1
    'Frame on
    Check6.Value = 1
    HScroll1.Value = 15
    'right and bottom margin
    HScroll4.Value = 32
    VScroll4.Value = 32
    'wall size and visible
    TChart1.Walls.Back.Transparent = False
    TChart1.Walls.Back.Brush.Color = vbGreen
    HScroll2.Value = 0
    Check11.Value = 1
    'Axes visible
    Check7.Value = 1
    Check8.Value = 1
    Check9.Value = 1
    Check10.Value = 1
    Option8.Value = True
    Option13.Value = True
    TChart1.Series(0).FillSampleValues 10
    TChart1.Series(1).FillSampleValues 10
    'BevelWidth 5 - 4
    VScroll2.Value = 4
    'No back colour
    Check4.Value = 1
End Sub

Private Sub HScroll1_Change()
' change 3D percent effect...
  TChart1.Aspect.Chart3DPercent = HScroll1.Value
  Label8.Caption = Str$(HScroll1.Value)
End Sub

Private Sub HScroll2_Change()
' change Walls size...
  TChart1.Walls.Left.Size = HScroll2.Value
  TChart1.Walls.Bottom.Size = HScroll2.Value
  TChart1.Walls.Back.Size = HScroll2.Value
  Label11.Caption = Str$(HScroll2.Value)
End Sub


Private Sub Hscroll3_Change()
  TChart1.Panel.MarginLeft = HScroll3.Value
End Sub

Private Sub HScroll4_Change()
  TChart1.Panel.MarginRight = 38 - HScroll4.Value
End Sub

Private Sub Option1_Click()
  With TChart1.Panel.Gradient
    If Option1.Value = True Then
      .Direction = gdLeftRight
    End If
  End With
  Option2.Value = False
  Option3.Value = False
  Option4.Value = False
End Sub

Private Sub Option10_Click()
  With TChart1.Panel
    If Option10.Value = True Then
      .BevelInner = bvRaised
    End If
  End With
  Option8.Value = False
  Option9.Value = False
End Sub

Private Sub Option11_Click()
  With TChart1.Panel
    If Option11.Value = True Then
      .BevelOuter = bvNone
    End If
  End With
  Option12.Value = False
  Option13.Value = False
End Sub

Private Sub Option12_Click()
  With TChart1.Panel
    If Option12.Value = True Then
      .BevelOuter = bvLowered
    End If
  End With
  Option11.Value = False
  Option13.Value = False
End Sub

Private Sub Option13_Click()
  With TChart1.Panel
    If Option13.Value = True Then
      .BevelOuter = bvRaised
    End If
  End With
  Option11.Value = False
  Option12.Value = False
End Sub

Private Sub Option2_Click()
  With TChart1.Panel.Gradient
    If Option2.Value = True Then
      .Direction = gdRightLeft
    End If
  End With
  Option1.Value = False
  Option3.Value = False
  Option4.Value = False
End Sub


Private Sub Option3_Click()
  With TChart1.Panel.Gradient
    If Option3.Value = True Then
      .Direction = gdTopBottom
    End If
  End With
  Option1.Value = False
  Option2.Value = False
  Option4.Value = False
End Sub


Private Sub Option4_Click()
  With TChart1.Panel.Gradient
    If Option4.Value = True Then
      .Direction = gdBottomTop
    End If
  End With
  Option1.Value = False
  Option2.Value = False
  Option3.Value = False
End Sub


Private Sub Option5_Click()
  With TChart1.Panel
    If Option5.Value = True Then
      .BackImageMode = pbmStretch
    End If
  End With
  Option6.Value = False
  Option7.Value = False
End Sub

Private Sub Option6_Click()
  With TChart1.Panel
    If Option6.Value = True Then
      .BackImageMode = pbmTile
    End If
  End With
  Option5.Value = False
  Option7.Value = False
End Sub


Private Sub Option7_Click()
  With TChart1.Panel
    If Option7.Value = True Then
      .BackImageMode = pbmCenter
    End If
  End With
  Option5.Value = False
  Option6.Value = False
End Sub


Private Sub Option8_Click()
  With TChart1.Panel
    If Option8.Value = True Then
      .BevelInner = bvNone
    End If
  End With
  Option9.Value = False
  Option10.Value = False
End Sub


Private Sub Option9_Click()
  With TChart1.Panel
    If Option9.Value = True Then
      .BevelInner = bvLowered
    End If
  End With
  Option8.Value = False
  Option10.Value = False
End Sub

Private Sub VScroll1_Change()
  TChart1.Panel.BevelOffset = VScroll1.Value
End Sub


Private Sub VScroll2_Change()
  TChart1.Panel.BevelWidth = VScroll2.Value
End Sub


Private Sub VScroll3_Change()
  TChart1.Panel.MarginTop = VScroll3.Value
End Sub


Private Sub VScroll4_Change()
  TChart1.Panel.MarginBottom = 38 - VScroll4.Value
End Sub





