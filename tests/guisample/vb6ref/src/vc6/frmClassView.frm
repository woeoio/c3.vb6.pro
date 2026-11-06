VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#20.3#0"; "CommandBars.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Controls.ocx"
Begin VB.Form frmClassView 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   454
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TreeView treeClassView 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      _Version        =   1048577
      _ExtentX        =   8281
      _ExtentY        =   7646
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ShowLines       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      IconSize        =   16
      Indentation     =   20
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   5640
      Top             =   1320
      _Version        =   1048577
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   9
   End
End
Attribute VB_Name = "frmClassView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandBars_Resize()
   On Error Resume Next
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.GetClientRect Left, Top, Right, Bottom
    treeClassView.Move Left, Top, Right - Left, Bottom - Top
End Sub

Private Sub Form_Load()

    Dim Root As TreeViewNode
    Dim Node As TreeViewNode
    
    Set treeClassView.Icons = SuiteControlsGlobalSettings.Icons
           
    Set Root = treeClassView.Nodes.Add(, , "Root", "XtremeDemo classes", ID_CLASS_TREE_1)
    Root.Bold = True
        
    Set Node = treeClassView.Nodes.Add("Root", xtpTreeViewChild, , "CAboutDlg", ID_CLASS_TREE_4)
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "CAboutDlg()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "DoDataExchange(CDataExchange *pDX)", ID_CLASS_TREE_6
    
    Set Node = treeClassView.Nodes.Add("Root", xtpTreeViewChild, , "CChildFrame", ID_CLASS_TREE_4)
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "AssertValid()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "CChildFrame()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "~CChildFrame()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "Dump(CDumpContext &dc)", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "PreCreateWindow(CREATESTRUCT& cs)", ID_CLASS_TREE_5
    
    Set Node = treeClassView.Nodes.Add("Root", xtpTreeViewChild, , "CMainFrame", ID_CLASS_TREE_4)
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "AssertValid()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "CMainFrame()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "~CMainFrame()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "Dump(CDumpContext &dc)", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "OnCreate(LPCREATESTRUCT lpCreateStruct)", ID_CLASS_TREE_6
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "PreCreateWindow(CREATESTRUCT& cs)", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "m_wndStatusBar", ID_CLASS_TREE_8
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "m_wndToolBar", ID_CLASS_TREE_8
    
    Set Node = treeClassView.Nodes.Add("Root", xtpTreeViewChild, , "CXtremeDemoApp", ID_CLASS_TREE_4)
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "CXtremeDemoApp()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "InitInstance()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "OnAppAbout()", ID_CLASS_TREE_5
    
    Set Node = treeClassView.Nodes.Add("Root", xtpTreeViewChild, , "CXtremeDemoDoc", ID_CLASS_TREE_4)
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "AssertValid()", ID_CLASS_TREE_6
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "CXtremeDemoDoc()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "~CXtremeDemoDoc()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "Dump(CDumpContext &dc)", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "OnNewDocument()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "Serialize(CArchive& ar)", ID_CLASS_TREE_5
    Node.Expanded = True
    
    Set Node = treeClassView.Nodes.Add("Root", xtpTreeViewChild, , "CXtremeDemoView", ID_CLASS_TREE_4)
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "AssertValid()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "CXtremeDemoView()", ID_CLASS_TREE_6
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "~CXtremeDemoView()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "Dump(CDumpContext &dc)", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "GetDocument()", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "OnBeginPrinting(CDC* pDC, CPrintInfo *pInfo)", ID_CLASS_TREE_5
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "OnDraw(CDC *pDC)", ID_CLASS_TREE_5
    
    Set Node = treeClassView.Nodes.Add("Root", xtpTreeViewChild, , "Globals", ID_CLASS_TREE_2)
    treeClassView.Nodes.Add Node, xtpTreeViewChild, , "theApp", ID_CLASS_TREE_7
    Node.Expanded = True
    
    Root.Expanded = True
    
    Dim ControlAdd As CommandBarPopup
    Dim ControlOpen As CommandBarButton
   
    With CommandBars.ActiveMenuBar.Controls
        Set ControlAdd = .Add(xtpControlSplitButtonPopup, ID_PROJECT_ADD_CLASS, "")
        ControlAdd.CommandBar.Controls.Add xtpControlButton, ID_PROJECT_ADD_CLASS, "Add Class..."
        ControlAdd.CommandBar.Controls.Add xtpControlButton, 0, "Add Resource..."
        
        Set ControlOpen = .Add(xtpControlButton, ID_FILE_OPEN, "Open File...")
        ControlOpen.BeginGroup = True
    End With
    
    CommandBars.ActiveMenuBar.EnableDocking xtpFlagStretched
    CommandBars.ActiveMenuBar.ModifyStyle XTP_CBRS_GRIPPER, 0
End Sub

Private Sub Form_Activate()
    CommandBars.VisualTheme = fMainForm.CommandBars.VisualTheme
    CommandBars.PaintManager.RefreshMetrics
    CommandBars.RecalcLayout
End Sub

Public Sub SetTheme(nTheme As AppearanceStyle)
    SetTreeTheme treeClassView, nTheme
End Sub

