VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   ScaleHeight     =   431
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   816
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00393939&
      Height          =   2415
      Left            =   9000
      ScaleHeight     =   2355
      ScaleWidth      =   2595
      TabIndex        =   8
      Top             =   840
      Width           =   2655
      Begin Proyecto1.ucSlider ucSlider1 
         Height          =   2415
         Index           =   5
         Left            =   480
         TabIndex        =   9
         Top             =   0
         Width           =   1695
         _extentx        =   2990
         _extenty        =   4260
         font            =   "Form1.frx":0000
         tickssize       =   15
         max             =   10
         ticklinescolor  =   16777215
         trackbackcolor  =   0
         trackforecolorpositive=   0
         gradientstyle   =   0   'False
         knobwidth       =   65
         knobheight      =   30
         mousepointer    =   0
         arrimgknob      =   "Form1.frx":002C
      End
   End
   Begin Proyecto1.ucSlider ucSlider1 
      Height          =   615
      Index           =   3
      Left            =   9120
      TabIndex        =   7
      Top             =   120
      Width           =   2895
      _extentx        =   5106
      _extenty        =   1085
      font            =   "Form1.frx":0CF4
      tickssize       =   5
      orientationhorizontal=   -1  'True
      trackforecolorpositive=   48384
      ticktext        =   3
      ticklines       =   3
      biselsize       =   4
      knobwidth       =   30
      knobheight      =   30
      mousepointer    =   0
      arrimgknob      =   "Form1.frx":0D20
      value           =   50
   End
   Begin Proyecto1.LabelPlus LP1 
      Height          =   495
      Left            =   9960
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   855
      _extentx        =   1508
      _extenty        =   873
      backcolor       =   0
      backcoloropacity=   70
      backshadow      =   0   'False
      border          =   -1  'True
      bordercolor     =   16777215
      bordercornerlefttop=   7
      bordercornerrighttop=   7
      bordercornerbottomright=   7
      bordercornerbottomleft=   7
      borderwidth     =   1
      captionalignmenth=   1
      captionalignmentv=   1
      caption         =   "Form1.frx":0D38
      font            =   "Form1.frx":0D5A
      forecolor       =   16777215
      shadowcoloropacity=   0
      calloutposicion =   3
      calloutlen      =   7
      callout         =   -1  'True
      mousepointer    =   0
      iconfont        =   "Form1.frx":0D82
      iconforecolor   =   0
      iconopacity     =   0
   End
   Begin Proyecto1.ucSlider ucSlider1 
      Height          =   1095
      Index           =   4
      Left            =   9120
      TabIndex        =   5
      Top             =   4560
      Width           =   2655
      _extentx        =   4683
      _extenty        =   1931
      font            =   "Form1.frx":0DAE
      tickssize       =   5
      orientationhorizontal=   -1  'True
      trackforecolorpositive=   29695
      ticktext        =   3
      ticklines       =   3
      knobstyle       =   1
      biselsize       =   4
      knobwidth       =   30
      knobheight      =   30
      mousepointer    =   0
      arrimgknob      =   "Form1.frx":0DDA
      value           =   50
   End
   Begin Proyecto1.ucSlider ucSlider6 
      Height          =   975
      Left            =   9000
      TabIndex        =   4
      Top             =   3840
      Width           =   2895
      _extentx        =   5106
      _extenty        =   1720
      font            =   "Form1.frx":0DF2
      tickslongfrequency=   25
      tickssmallhiden =   -1  'True
      tickssize       =   5
      min             =   -3
      max             =   3
      orientationhorizontal=   -1  'True
      tickdistance    =   3
      tracksize       =   5
      trackbackcolor  =   3750201
      ticktext        =   3
      ticklines       =   3
      gradientstyle   =   0   'False
      knobwidth       =   50
      knobheight      =   30
      mousepointer    =   0
      arrimgknob      =   "Form1.frx":0E1E
   End
   Begin Proyecto1.ucSlider ucSlider4 
      Height          =   4935
      Left            =   6840
      TabIndex        =   3
      Top             =   480
      Width           =   1695
      _extentx        =   2990
      _extenty        =   8705
      font            =   "Form1.frx":141C
      tickssize       =   5
      trackbackcolor  =   0
      trackforecolorpositive=   35897
      ticktext        =   3
      gradientstyle   =   0   'False
      knobwidth       =   30
      knobheight      =   50
      mousepointer    =   0
      arrimgknob      =   "Form1.frx":1448
   End
   Begin Proyecto1.ucSlider ucSlider1 
      Height          =   4695
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1695
      _extentx        =   2990
      _extenty        =   8281
      font            =   "Form1.frx":1A46
      tickssize       =   5
      knobwidth       =   30
      knobheight      =   30
      mousepointer    =   0
      mousepointerhands=   -1
      arrimgknob      =   "Form1.frx":1A72
   End
   Begin Proyecto1.ucSlider ucSlider1 
      Height          =   4935
      Index           =   2
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   1695
      _extentx        =   2990
      _extenty        =   8705
      font            =   "Form1.frx":1A8A
      tickssize       =   5
      trackbackcolor  =   0
      gradientstyle   =   0   'False
      knobwidth       =   30
      knobheight      =   50
      mousepointer    =   0
      arrimgknob      =   "Form1.frx":1AB6
   End
   Begin Proyecto1.ucSlider ucSlider1 
      Height          =   4935
      Index           =   1
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   1695
      _extentx        =   2990
      _extenty        =   8705
      font            =   "Form1.frx":1D32
      tickssize       =   5
      trackbackcolor  =   9737364
      trackforecolorpositive=   16756058
      knobcolor       =   16756058
      biselcolor      =   2179584
      gradientstyle   =   0   'False
      knobwidth       =   30
      knobheight      =   30
      mousepointer    =   0
      arrimgknob      =   "Form1.frx":1D5E
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cSubClass As clsSubClass
Attribute cSubClass.VB_VarHelpID = -1
Private Const WM_MOUSEWHEEL As Long = &H20A
Private mCtlSlider As ucSlider


Private Sub cSubClass_WindowProc(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
    Dim lStep As Long
    If uMsg = WM_MOUSEWHEEL Then
        If Not mCtlSlider Is Nothing Then
            If mCtlSlider.Max >= 100 Then lStep = 10 Else lStep = 1
            If wParam > 0 Then
                If mCtlSlider.Value < mCtlSlider.Max Then
                    mCtlSlider.Value = mCtlSlider.Value + lStep
                End If
            Else
                If mCtlSlider.Value > mCtlSlider.Min Then
                    mCtlSlider.Value = mCtlSlider.Value - lStep
                End If
            End If
        End If
    End If
End Sub


Private Sub Form_Load()
    Set cSubClass = New clsSubClass
    cSubClass.ssc_Subclass Me.hwnd
    cSubClass.ssc_AddMsg Me.hwnd, WM_MOUSEWHEEL, MSG_AFTER
    cSubClass.ssc_Subclass Picture1.hwnd
    cSubClass.ssc_AddMsg Picture1.hwnd, WM_MOUSEWHEEL, MSG_AFTER
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set mCtlSlider = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cSubClass = Nothing
End Sub

Private Sub ucSlider1_MouseEnter(INDEX As Integer)
    Set mCtlSlider = ucSlider1(INDEX)
End Sub

Private Sub ucSlider4_Change()
    DoEvents
    LP1.Left = ucSlider4.Left + ucSlider4.KnobLeft - LP1.Width - 10 * LP1.GetWindowsDPI
    LP1.Top = ucSlider4.Top + ucSlider4.KnobTop + ucSlider4.KnobHeight * LP1.GetWindowsDPI / 2 - LP1.Height / 2
    LP1.Caption = ucSlider4.Value
End Sub

Private Sub ucSlider4_MouseEnter()
    LP1.CallOutPosicion = coRight
    LP1.Visible = True
    ucSlider4_Change
    Set mCtlSlider = ucSlider4
End Sub

Private Sub ucSlider4_MouseLeave()
    LP1.Visible = False
End Sub


Private Sub ucSlider6_Change()
    DoEvents
    LP1.Left = ucSlider6.Left + ucSlider6.KnobLeft + ucSlider6.KnobWidth * LP1.GetWindowsDPI / 2 - LP1.Width / 2
    LP1.Top = ucSlider6.Top + ucSlider6.KnobTop - LP1.Height - 4 * LP1.GetWindowsDPI
    LP1.Caption = ucSlider6.Value
End Sub

Private Sub ucSlider6_MouseEnter()
    LP1.CallOutPosicion = coBottom
    LP1.Visible = True
    ucSlider6_Change
    Set mCtlSlider = ucSlider6
End Sub

Private Sub ucSlider6_MouseLeave()
    LP1.Visible = False
End Sub

