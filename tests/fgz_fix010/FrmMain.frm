VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "F10"
   ClientHeight    =   3000
   ClientWidth     =   4000
   Begin VB.PictureBox picTop 
      BackColor       =   &H00F8CCB1&
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   1215
   End
   Begin VB.Image imgTop 
      Height          =   495
      Left            =   1500
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picBox 
      Height          =   1215
      Left            =   120
      Top             =   800
      Width           =   2400
      Begin VB.Label labIn 
         BackColor       =   &H000000FF&
         Caption         =   "in"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         Top             =   120
         Width           =   1215
      End
      Begin VB.Image imgIn 
         Height          =   255
         Left            =   1500
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   120
   End
   Begin VB.Timer Watch 
      Enabled         =   -1   'True
      Interval        =   2000
      Left            =   3000
      Top             =   240
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' FIX-010 (upstream FGZ). Bare `Timer1.Enabled = True` takes the direct
' vb6_SetTimerEnabled path, so use an Object variable to force the late-bound
' vb6_ComSetProp(t, L"Enabled") / L"Interval" path into vb6_Host_SetProp.
' The design-time Interval is 1000 ms; the runtime write should shorten it to
' 50 ms. If either branch is missing, Timer1 cannot deliver 3 beats before the
' 2000 ms watchdog, which prints F10-TIMEOUT instead (and the needle goes red).
Private gTicks As Long

Private Sub Form_Load()
    Dim t As Object
    Set t = Timer1
    t.Interval = 50
    t.Enabled = True
End Sub

Private Sub Timer1_Timer()
    gTicks = gTicks + 1
    If gTicks < 3 Then Exit Sub
    Debug.Print "F10-TICKS=3"
    Unload Me
End Sub

Private Sub Watch_Timer()
    Debug.Print "F10-TIMEOUT"
    Unload Me
End Sub
