VERSION 5.00
Begin VB.Form ZeroForm 
   Caption         =   "ZeroForm"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   4680
   Begin VB.ListBox List1 
      Height          =   600
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1200
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   1560
      TabIndex        =   3
      Text            =   "x"
      Top             =   120
      Width           =   1200
   End
   Begin VB.PictureBox Pic1 
      Height          =   600
      Left            =   1560
      ScaleHeight     =   540
      ScaleWidth      =   1140
      TabIndex        =   4
      Top             =   600
      Width           =   1200
   End
   Begin MSWinsockLib.Winsock wsZ 
      Left            =   3000
      Top             =   120
   End
   Begin MSComctlLib.ImageList ImageZ 
      _ExtentX        =   449
      _ExtentY        =   449
      _Version        =   49152
      ImageWidth      =   16
      ImageHeight     =   16
   End
   Begin MSComctlLib.StatusBar StatusBarZ 
      Align           =   2
      Height          =   255
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   450
      Style           =   0
   End
   Begin MSComDlg.CommonDialog dlZ 
      Left            =   3000
      Top             =   1200
      CancelError     =   -1
      DialogTitle     =   "z"
      Filter          =   "All|*.*"
      Flags           =   0
      InitDir         =   "."
      DefaultExt      =   "txt"
   End
End
Attribute VB_Name = "ZeroForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ledger 278 knife 12 (B72): the control-method tables answer BOTH questions - which C exit
' and how many arguments that exit takes. These four call shapes are the ones the statement /
' expression / With docks emit today; the sentinel compares the emitted call arity with the
' table and with the RTL prototype, so a future param added to an RTL exit cannot leave the
' codegen shipping the old count (locally that is only warning C4020, on the newer cl it is
' error C2197 -- the shape that already cost gate #371 once).
Private Sub DoZero()
    List1.Clear
    Combo1.Clear
    Text1.SetFocus
End Sub
Private Function Sizes() As Long
    Sizes = Pic1.TextHeight("ab") + Pic1.TextWidth("cd")
End Function
Private Function Conv() As Long
    Conv = ScaleX(1440, 1, 3) + ScaleY(1440, 1, 3)
End Function
' ledger 278 knife 14 (B72): two collection Clears (different receiver expressions: the ImageList
' slot is vb6_com_<name>, the StatusBar one is vb6_hwnd_<name>) plus the six CommonDialog Show*
' exits. The Show* names used to be built by string surgery at both docks.
Private Sub DoColl()
    ImageZ.ListImages.Clear
    StatusBarZ.Panels.Clear
End Sub
Private Sub DoDlg()
    dlZ.ShowOpen
    dlZ.ShowSave
    dlZ.ShowColor
    dlZ.ShowFont
    dlZ.ShowPrinter
    dlZ.ShowAbout
End Sub
' ledger 278 knife 13 (B72): the Winsock family table answers the same two questions. The
' three argument-less shapes end at the statement dock (that dock may only take rows whose
' count is 1); the five shapes with arguments end at the expression dock. The sentinel
' compares each emitted call's argument count with the table and with the RTL prototype.
Private Sub DoWs()
    wsZ.Close
    wsZ.Listen
    wsZ.Connect
End Sub
Private Sub SetWs()
    wsZ.Accept 5
    wsZ.Bind 8001, "127.0.0.1"
    wsZ.SendData "hi"
End Sub
Private Sub ReadWs()
    Dim sZ As String
    wsZ.GetData sZ, 8, 100
    wsZ.PeekData sZ, 8, 100
End Sub
