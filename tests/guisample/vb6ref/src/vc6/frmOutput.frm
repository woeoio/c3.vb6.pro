VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Controls.ocx"
Begin VB.Form frmOutput 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   333
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.FlatEdit txtOutput 
      Height          =   1335
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3135
      _Version        =   1048577
      _ExtentX        =   5530
      _ExtentY        =   2355
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "FlatEdit1"
      MultiLine       =   -1  'True
      UseVisualStyle  =   0   'False
      ShowBorder      =   0   'False
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
txtOutput.Text = _
    "GUI_VisualStudio10.exe': Loaded 'C:\Codejock.com\cvs\ToolkitPro\Samples\GUI_Samples\GUI_VisualStudio10\Debug\GUI_VisualStudio10.exe', Symbols loaded" & vbCrLf & _
    "GUI_VisualStudio10.exe': Loaded 'D:\WINDOWS\system32\ntdll.dll', No symbols loaded" & vbCrLf & _
    "GUI_VisualStudio10.exe': Loaded 'D:\\WINDOWS\\system32\\kernel32.dll', No symbols loaded." & vbCrLf & _
    "GUI_VisualStudio10.exe': Loaded 'D:\\WINDOWS\\system32\\mfc70d.dll', Symbols loaded." & vbCrLf & _
    "GUI_VisualStudio10.exe': Loaded 'D:\\WINDOWS\\system32\\msvcr70d.dll', Symbols loaded." & vbCrLf & _
    "GUI_VisualStudio10.exe': Loaded 'D:\\WINDOWS\\system32\\gdi32.dll', No symbols loaded."
    

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Select Case txtOutput.Appearance
        Case xtpAppearanceVisualStudio2012Dark, xtpAppearanceVisualStudio2012Light
            txtOutput.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
        Case Else
            txtOutput.Move Me.ScaleLeft + 1, Me.ScaleTop + 1, Me.ScaleWidth - 2, Me.ScaleHeight - 2
    End Select

End Sub

Public Sub SetTheme(nTheme As AppearanceStyle)
    
    txtOutput.Appearance = nTheme
    

End Sub

