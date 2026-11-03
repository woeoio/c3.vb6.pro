VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Controls.ocx"
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#20.3#0"; "SYNTAX~1.OCX"
Begin VB.Form frmDocument 
   BorderStyle     =   0  'None
   Caption         =   "frmDocument"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   Begin XtremeSyntaxEdit.SyntaxEdit wndSyntaxEdit 
      Height          =   2895
      Left            =   2760
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      _Version        =   1048577
      _ExtentX        =   7435
      _ExtentY        =   5106
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   -1  'True
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
   End
   Begin XtremeSuiteControls.ScrollBar ScrollBarVertical 
      Height          =   2415
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   255
      _Version        =   1048577
      _ExtentX        =   450
      _ExtentY        =   -1
      _StockProps     =   64
      Orientation     =   1
      UseVisualStyle  =   0   'False
      Appearance      =   14
   End
   Begin XtremeSuiteControls.ScrollBar ScrollBarHorizontal 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   3480
      Width           =   2895
      _Version        =   1048577
      _ExtentX        =   5106
      _ExtentY        =   -1
      _StockProps     =   64
      UseVisualStyle  =   0   'False
      Appearance      =   14
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fMain As frmMain
Public bLoadConfigOnce As Boolean
Dim CurrentTheme As Long
Dim sIni As String

Private Sub Form_Activate()
    RefreshColors
End Sub

Private Sub Form_Load()
    lDocumentCount = lDocumentCount + 1
    Caption = "Document " & lDocumentCount
    
    ScrollBarVertical.ZOrder 0
    ScrollBarHorizontal.ZOrder 0
    wndSyntaxEdit.SetScrollBars ScrollBarHorizontal.hwnd, ScrollBarVertical.hwnd
    
    wndSyntaxEdit.FileMonitorMode = True
    wndSyntaxEdit.Text = LoadSampleFile(App.Path + "\SampleFrm.txt")
    
    bLoadConfigOnce = True
    CurrentTheme = -1
    sIni = ""

    RefreshColors
    
    Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    wndSyntaxEdit.Move 0, 0, Me.ScaleWidth - 0, Me.ScaleHeight - 0
    wndSyntaxEdit.RightMargin = wndSyntaxEdit.Width - 400
    
    ScrollBarHorizontal.Move 0, wndSyntaxEdit.Height - ScrollBarHorizontal.Height, wndSyntaxEdit.Width - ScrollBarVertical.Width, ScrollBarHorizontal.Height
    ScrollBarVertical.Move Me.ScaleWidth - ScrollBarVertical.Width, 0, ScrollBarVertical.Width, wndSyntaxEdit.Height '- ScrollBarHorizontal.Height
End Sub

Function LoadSampleFile(FileName$) As String
    Dim fileHandle As Integer
    Dim FileText As String
    
    If Len(Dir$(FileName$)) = 0 Then
        Err.Raise 53
    End If
    
    fileHandle = FreeFile
    Open FileName$ For Input As #fileHandle
    FileText = Input$(LOF(fileHandle), fileHandle)
    Close #fileHandle
    
    LoadSampleFile = FileText
End Function

Public Sub RefreshColors()

    If (CurrentTheme = AppTheme) Then
        Exit Sub
    End If
    
    ScrollBarHorizontal.Appearance = ControlTheme
    ScrollBarVertical.Appearance = ControlTheme
   
    Select Case AppTheme
        Case ID_THEME_VC11_LIGHT, ID_THEME_VC14_LIGHT
            sIni = "\Config\SyntaxEditLight.ini"
        Case ID_THEME_VC11_DARK, ID_THEME_VC14_DARK
            sIni = "\Config\SyntaxEditDark.ini"
        Case Else
            sIni = "\Config\SyntaxEdit.ini"
    End Select

    Dim clrBackColor As Long
    Dim clrLineNumberBackColor As Long
    Dim clrLineNumberTextColor As Long
    
    clrBackColor = RGB(255, 255, 255)
    clrLineNumberBackColor = RGB(255, 255, 255)
    clrLineNumberTextColor = RGB(43, 145, 175)

    Dim bSelMargin As Boolean
    bSelMargin = True
    
    Dim sSchema As String
    sSchema = "\Config\Schemas\_cpp.schclass"

    Select Case AppTheme
        Case ID_THEME_VC11_LIGHT, ID_THEME_VC14_LIGHT
            clrBackColor = RGB(246, 246, 246)
            clrLineNumberBackColor = RGB(246, 246, 246)
            clrLineNumberTextColor = RGB(43, 145, 175)

            bSelMargin = True
            sSchema = "\Config\Schemas\_cppLight.schclass"
        Case ID_THEME_VC11_DARK, ID_THEME_VC14_DARK
            clrBackColor = RGB(28, 28, 28)
            clrLineNumberBackColor = RGB(28, 28, 28)
            clrLineNumberTextColor = RGB(86, 156, 214)

            bSelMargin = False
            sSchema = "\Config\Schemas\_cppDark.schclass"
        Case Else
            clrBackColor = RGB(255, 255, 255)
            clrLineNumberBackColor = RGB(255, 255, 255)
            clrLineNumberTextColor = RGB(43, 145, 175)

            bSelMargin = True
            sSchema = "\Config\Schemas\_cpp.schclass"
    End Select

    CurrentTheme = AppTheme

    wndSyntaxEdit.PaintManager.BackColor = clrBackColor
    wndSyntaxEdit.PaintManager.LineNumberBackColor = clrLineNumberBackColor
    wndSyntaxEdit.PaintManager.LineNumberTextColor = clrLineNumberTextColor
    wndSyntaxEdit.ShowSelectionMargin = bSelMargin

    wndSyntaxEdit.DataManager.FileExt = ".cpp"

    If (bLoadConfigOnce) Then
        bLoadConfigOnce = False

        wndSyntaxEdit.ConfigFile = App.Path + sIni
    Else
        If (strSchema = "") Then
            
            Dim strINI As String
            Dim strNewSchema As String
            
            strINI = LoadSampleFile(App.Path + sIni)
            strNewSchema = LoadSampleFile(App.Path + sSchema)

            strSchema = strNewSchema
            strColor = strINI

            wndSyntaxEdit.SetSyntaxAndColorScheme strSchema, strColor, False, False
        Else
            wndSyntaxEdit.SetSyntaxAndColorScheme strSchema, strColor, False, False
        End If
    End If
End Sub
