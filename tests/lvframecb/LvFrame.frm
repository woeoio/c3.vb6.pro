VERSION 5.00
Begin VB.Form LvFrame 
   Caption         =   "LvFrame"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   6600
   Begin VB.Timer Timer1 
      Enabled         =   -1  'True
      Interval        =   40
      Left            =   6000
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COUT"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   120
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "F"
      Height          =   900
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4200
      Begin VB.CommandButton Command1 
         Caption         =   "CIN"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1000
      End
      Begin VB.CheckBox Check1 
         Caption         =   "CB0"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   1200
      End
   End
End
Attribute VB_Name = "LvFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' No ActiveX control in this fixture on purpose: CI has no VB6 OCX installed, and
' C3 falls back to a bare identifier when the typelib is missing (VB4001 -> C2065),
' so anything that declares ComctlLib.ListView can only build on a dev box.
' The list-view side of the enum bug is covered by tests/extlist (dev box only);
' what is gated here is the enum evaluation itself plus control events through a Frame.

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const BM_CLICK_MSG As Long = 245

' Intra-module: enum members referencing module-level Consts must be evaluated,
' not fall back to the declaration ordinal.
Private Const LSX_FULL As Long = &H20
Private Const LSX_TRACK As Long = &H8
Private Enum LvFlag
    FlagFull = LSX_FULL
    FlagTrack = LSX_TRACK
    FlagBoth = FlagFull Or FlagTrack
End Enum

Private Sub Form_Load()
    ' Expect 32/8/40; ordinal fallback would print 0/1/1
    Debug.Print "ENUM=" & LvFlag.FlagFull & "/" & LvFlag.FlagTrack & "/" & LvFlag.FlagBoth
    ' Cross-module: defined in LvMod.bas, referenced here
    Debug.Print "ENUM.X=" & LvFlag2.Flag2Full & "/" & LvFlag2.Flag2Track & "/" & LvFlag2.Flag2Both
End Sub

Private Sub Check1_Click(Index As Integer)
    Debug.Print "CB=" & Index
End Sub

Private Sub Command1_Click(Index As Integer)
    Debug.Print "CMDIN"
End Sub

Private Sub Command2_Click()
    Debug.Print "CMDOUT"
End Sub

Private Sub Timer1_Timer()
    Static step As Long
    step = step + 1
    Select Case step
    Case 1
        ' Native path: this is what a real mouse click ends up sending to the parent
        SendMessage Check1(0).hWnd, BM_CLICK_MSG, 0&, ByVal 0&
        Debug.Print "AFTER_BMCLICK chk=" & Check1(0).Value
        Check1(0).Value = 1
        Debug.Print "AFTER_PROG chk=" & Check1(0).Value
        SendMessage Command1(0).hWnd, BM_CLICK_MSG, 0&, ByVal 0&
        SendMessage Command2.hWnd, BM_CLICK_MSG, 0&, ByVal 0&
    Case 4
        Debug.Print "LVFRAME-DONE"
        Unload Me
    End Select
End Sub
