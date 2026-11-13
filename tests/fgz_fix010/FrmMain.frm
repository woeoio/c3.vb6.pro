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
      Interval        =   3000
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
' GA judgment is the READ-BACK pair F10-IVAL=50 / F10-EN0=False: the late-bound
' writes must land in the window properties that vb6_GetTimerInterval /
' vb6_GetTimerEnabled read. Without the RTL dispatch both read their defaults
' (0 and True), so both needles go red. This does NOT depend on WM_TIMER being
' delivered -- GitHub-hosted runners do not deliver generated Timer events even
' though the same CI-built C3.exe produces F10-TICKS=3 on a local desktop.
' Locally the F10-TICKS=3 / F10-TIMEOUT pair is still the runtime A/B judgment.
' The watchdog line carries its own diagnosis (loaded / ticks) for that A/B.
Private gTicks As Long
Private gLoaded As Long

Private Sub Form_Load()
    Dim t As Object
    gLoaded = 1
    Debug.Print "F10-LOAD"
    Set t = Timer1
    t.Interval = 50
    Debug.Print "F10-IVAL=" & CStr(Timer1.Interval)
    t.Enabled = False
    Debug.Print "F10-EN0=" & CStr(Timer1.Enabled)
    t.Enabled = True
    Debug.Print "F10-EN1=" & CStr(Timer1.Enabled)
End Sub

Private Sub Timer1_Timer()
    gTicks = gTicks + 1
    Debug.Print "F10-TICK=" & CStr(gTicks)
    If gTicks < 3 Then Exit Sub
    Debug.Print "F10-TICKS=3"
    Unload Me
End Sub

Private Sub Watch_Timer()
    Debug.Print "F10-TIMEOUT loaded=" & CStr(gLoaded) & " ticks=" & CStr(gTicks)
    Unload Me
End Sub
