VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Controls.ocx"
Begin VB.Form frmMarkupView 
   BorderStyle     =   0  'None
   Caption         =   "Start Page"
   ClientHeight    =   3675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3675
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.Resizer Resizer1 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4590
      _Version        =   786432
      _ExtentX        =   8096
      _ExtentY        =   6482
      _StockProps     =   1
      BackColor       =   -2147483635
      VScrollLargeChange=   1000
      VScrollSmallChange=   200
      HScrollLargeChange=   1000
      HScrollSmallChange=   200
      AutoSize        =   -1  'True
      ControlCount    =   1
      Control(0).Caption=   "lblMarkup"
      Control(0).Width=   100
      Control(0).Height=   100
      Begin XtremeSuiteControls.Label lblMarkup 
         Height          =   3675
         Left            =   -75
         TabIndex        =   1
         Top             =   0
         Width           =   4710
         _Version        =   786432
         _ExtentX        =   8308
         _ExtentY        =   6482
         _StockProps     =   79
         BackColor       =   -2147483640
         EnableMarkup    =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMarkupView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lblMarkup.Caption = LoadXML(App.Path + "\StartPageSample.xaml")
End Sub

Function LoadXML(xmlFileName$) As String
    Dim fileHandle As Integer
    
    If Len(Dir$(xmlFileName)) = 0 Then
        Err.Raise 53
    End If
    
    fileHandle = FreeFile
    Open xmlFileName$ For Input As #fileHandle
    FileText = Input$(LOF(fileHandle), fileHandle)
    Close #fileHandle
    
    LoadXML = FileText
End Function
