VERSION 5.00
Begin VB.Form QForm 
   Caption         =   "QForm"
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2400
   ScaleHeight     =   1200
   ScaleWidth      =   2400
   Begin VB.Timer t1 
      Enabled         =   -1   'True
      Interval        =   40
      Left            =   120
      Top             =   120
   End
   Begin VB.Timer t2 
      Enabled         =   -1   'True
      Interval        =   20
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "QForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Ledger 259. VB6 ends the process when the last form unloads. The RTL does that with ONE
' WM_QUIT, posted by vb6_Forms_Unregister, and a thread only ever gets one: whoever dequeues
' it owns the exit. An inner pump that takes it (DoEvents, or a modal Show loop) leaves the
' outer GetMessage waiting forever with no window left, so the process just sits there.
'
' Not theoretical: gate #395's only red was combofocus_x86 writing all seven judgment lines
' and then never exiting (60 s, 62 ms CPU, no window).
'
' SHAPE, reshaped for gate #403 (this fixture was that gate's only real red). The unload has
' to land INSIDE an inner pump, and it has to land there every run. It used to be done with
' ONE timer: tick 6 slept 80 ms so tick 7's WM_TIMER was already queued, and DoEvents
' dispatched it, which put the unloading tick inside that pump. That only worked because a
' Timer event process was re-entrant. 90fb8069 made it NOT re-entrant (VB6 semantics:
' vb6_DispatchTimer drops the next beat of the SAME timer while its process has not
' returned), so gate #403 read Q-ORDER=unload-outside. The guard is right; the fixture is
' what went stale. The hole is still there, so the witness moves to the OTHER timer: t1
' (40 ms) still sleeps 80 ms before its tick-6 DoEvents, which guarantees t2's (20 ms)
' WM_TIMER is already in the queue when that pump starts, and t2 owns the unload -- a
' different timer's beat is exactly what the new guard does NOT drop. gLetGo is offered only
' on tick 6, so t2 can never unload early, and the unload can never be credited to a pump it
' did not run in.
'
'   Q-ORDER=unload-inside-doevents   witness that the unload really happened inside the pump
'                                    (without it a build that never reaches the hole would
'                                     still pass)
'   Q-ORDER=unload-outside           the pump ran but t2 never took the offer: the shape is
'                                    gone and the quit path is not being tested -- red by the
'                                    needle above, by design
'   Q-DONE                           the judgments were all written
'   the process exiting on its own   is the actual judgment -- run_tests.ps1 fails a case on
'                                    a run timeout, so a swallowed quit is red by itself
' ASCII only.

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private gTick As Long
Private gLetGo As Long
Private gUnloaded As Long

Private Sub t1_Timer()
    Dim my As Long
    my = gTick + 1
    gTick = my
    If my < 6 Then
        Sleep 80
        DoEvents
        Exit Sub
    End If
    If my > 6 Then Exit Sub
    gLetGo = 1
    Sleep 80
    DoEvents
    If gUnloaded = 1 Then
        Debug.Print "Q-ORDER=unload-inside-doevents"
    Else
        Debug.Print "Q-ORDER=unload-outside"
    End If
    gLetGo = 0
    Debug.Print "Q-DONE"
End Sub

Private Sub t2_Timer()
    If gLetGo = 0 Then Exit Sub
    If gUnloaded = 1 Then Exit Sub
    gUnloaded = 1
    gLetGo = 0
    t1.Enabled = False
    t2.Enabled = False
    Unload Me
End Sub
