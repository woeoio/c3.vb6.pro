VERSION 5.00
Begin VB.Form frmSet 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "seed"
      Top             =   120
      Width           =   2400
   End
   Begin VB.CheckBox Check1 
      Caption         =   "chk"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "cmd"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1500
   End
End
Attribute VB_Name = "frmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cmdW As CommandButton
' ledger 258 (B87). Comments stay ASCII: C3 reads sources as ANSI/GBK (ledger 222).
' The right-hand side of a Set statement is an OBJECT-REFERENCE context: VB6 hands over
' the control itself, the default property only unfolds in a VALUE context - so the line
' "s = Text1" is Text1.Text while "Set o = Text1" is the control.
Private Sub Form_Load()
    Dim o As Object
    Dim v As Variant
    Dim s As String
    Dim got As String
    Dim e1 As Long
    On Error Resume Next
    Text1.Text = "alpha"
    ' SO1 existence: a write through the slot lands on the window AND reads back through
    ' the slot. The counter-witness is built in: Text1 started as "alpha", so own=beta
    ' proves the write moved and read=beta proves the slot is not blank.
    Err.Clear
    Set o = Text1
    o.Text = "beta"
    e1 = Err.Number
    got = o.Text
    Debug.Print "SO1-RAW err=" & e1 & " tn=" & TypeName(o) & " own=" & Text1.Text & " read=" & got
    Debug.Print "SO1=" & CStr(e1 = 0 And TypeName(o) = "TextBox" And Text1.Text = "beta" And got = "beta")
    ' SO2 identity: the slot is a live reference, not a snapshot of the text.
    Text1.Text = "gamma"
    got = o.Text
    Debug.Print "SO2-RAW got=" & got & " own=" & Text1.Text
    Debug.Print "SO2=" & CStr(got = "gamma" And Text1.Text = "gamma")
    ' SO3 existence, second slot kind: a Variant target must hold an OBJECT (9), not the
    ' default-property value (Check1.Value is an Integer, so vt=3 before the fix).
    Err.Clear
    Set v = Check1
    Debug.Print "SO3-RAW err=" & Err.Number & " vt=" & VarType(v) & " tn=" & TypeName(v)
    Debug.Print "SO3=" & CStr(Err.Number = 0 And VarType(v) = 9 And TypeName(v) = "CheckBox")
    ' SO4 counter-witness: the SAME bare name in a VALUE context still unfolds to the
    ' default property. Without this head, "never fold anywhere" also reads green.
    s = Text1
    Debug.Print "SO4-RAW s=" & s
    Debug.Print "SO4=" & CStr(s = "gamma")
    ' SO5 witness: Set into a typed control variable - the shape the old after-the-fact
    ' string surgery used to rescue. It must still work now that the reference context is
    ' decided before the value is emitted.
    Err.Clear
    Set cmdW = Command1
    Debug.Print "SO5-RAW err=" & Err.Number & " cap=" & cmdW.Caption
    Debug.Print "SO5=" & CStr(Err.Number = 0 And cmdW.Caption = "cmd")
    ' SO6 witness: the reference handed to an As Object parameter is the same shape (the
    ' raw control window) and answers a late-bound property write on the real window.
    Debug.Print "SO6-RAW " & LateProbe(Text1)
    Debug.Print "SO6=" & CStr(Text1.Text = "delta")
    Debug.Print "SO-DONE"
    Unload Me
End Sub

Private Function LateProbe(ByVal o As Object) As String
    Dim e2 As Long
    Dim got2 As String
    On Error Resume Next
    Err.Clear
    o.Text = "delta"
    e2 = Err.Number
    got2 = o.Text
    LateProbe = "err=" & e2 & " got=" & got2 & " own=" & Text1.Text
End Function
