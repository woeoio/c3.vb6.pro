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

' FIX-009 section-7 (upstream FGZ). The six remaining dispatch arms need the
' RUNTIME array index as first arg when the handler declares Index As Integer.
'   lblItem  : container-hosted array -> all six handlers declare Index
'   labPlain : non-array guard -> same six handlers WITHOUT Index; their
'              emitted calls must stay byte-identical to the pre-s7 form.
' DblClick is the tricky one: a lone parameter is the array Index here,
' but a Cancel out-param on non-array UC/OCX handlers.
Private Sub lblItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "F9S7-ARR-MOUSEDOWN=" & CStr(Index)
End Sub

Private Sub lblItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "F9S7-ARR-MOUSEUP=" & CStr(Index)
End Sub

Private Sub lblItem_KeyPress(Index As Integer, KeyAscii As Integer)
    Debug.Print "F9S7-ARR-KEYPRESS=" & CStr(Index)
End Sub

Private Sub lblItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Debug.Print "F9S7-ARR-KEYDOWN=" & CStr(Index)
End Sub

Private Sub lblItem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Debug.Print "F9S7-ARR-KEYUP=" & CStr(Index)
End Sub

Private Sub lblItem_DblClick(Index As Integer)
    Debug.Print "F9S7-ARR-DBLCLICK=" & CStr(Index)
End Sub

Private Sub labPlain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "F9S7-PLAIN-MOUSEDOWN"
End Sub

Private Sub labPlain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "F9S7-PLAIN-MOUSEUP"
End Sub

Private Sub labPlain_KeyPress(KeyAscii As Integer)
    Debug.Print "F9S7-PLAIN-KEYPRESS"
End Sub

Private Sub labPlain_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print "F9S7-PLAIN-KEYDOWN"
End Sub

Private Sub labPlain_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "F9S7-PLAIN-KEYUP"
End Sub

Private Sub labPlain_DblClick()
    Debug.Print "F9S7-PLAIN-DBLCLICK"
End Sub

