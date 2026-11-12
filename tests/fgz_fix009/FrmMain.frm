VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "F9"
   ClientHeight    =   3000
   ClientWidth     =   4000
   Begin VB.PictureBox picHost 
      BackColor       =   &H0056402F&
      Height          =   1800
      Left            =   120
      Top             =   120
      Width           =   3000
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "B"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lblItem 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1500
      End
   End
   Begin VB.Label lblTop 
      BackStyle       =   0  'Transparent
      Caption         =   "T0"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label lblTop 
      BackStyle       =   0  'Transparent
      Caption         =   "T1"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   4
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label labPlain 
      Caption         =   "P"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2040
      Width           =   900
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' FIX-009 (upstream FGZ). Control-array events must carry the RUNTIME element
' index, not the compile-time info.arrayIndex; and MouseMove must be emitted
' with 5 args when the handler has 5 params (Index first).
'   lblItem  : Label array INSIDE a PictureBox container (the failing shape;
'             only one element used to get a subclass installed + WM_DESTROY
'             referenced an undeclared vb6_hwnd_lblItem => C2065)
'   lblTop   : top-level Label array (already had the P0-4 loop, but no index)
'   labPlain : non-array guard - must stay 4-arg MouseMove / 0-arg Click
Private Sub Form_Load()
    Debug.Print "F9-LOAD"
End Sub

Private Sub lblItem_Click(Index As Integer)
    Debug.Print "F9-CLICK=" & CStr(Index)
End Sub

Private Sub lblItem_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "F9-MOVE=" & CStr(Index)
End Sub

Private Sub lblTop_Click(Index As Integer)
    Debug.Print "F9-TOPCLICK=" & CStr(Index)
End Sub

Private Sub labPlain_Click()
    Debug.Print "F9-PLAINCLICK"
End Sub

Private Sub labPlain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "F9-PLAINMOVE"
End Sub
