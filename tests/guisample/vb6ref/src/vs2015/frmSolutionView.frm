VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#20.3#0"; "CommandBars.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Controls.ocx"
Begin VB.Form frmSolutionView 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4590
   ForeColor       =   &H80000010&
   LinkTopic       =   "Form1"
   ScaleHeight     =   196
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TreeView treeSolutionView 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      _Version        =   1048577
      _ExtentX        =   7223
      _ExtentY        =   3413
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ShowLines       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      IconSize        =   16
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   960
      Top             =   2400
      _Version        =   1048577
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   9
   End
End
Attribute VB_Name = "frmSolutionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    
    Dim frmMain As frmMain
    Set frmMain = Forms(0)
    
    Select Case Control.Id
        Case ID_VIEW_PROPERTIESWINDOW: frmMain.DockingPaneManager.ShowPane ID_PANE_PROPERTIES
    End Select
End Sub

Private Sub CommandBars_Resize()
   On Error Resume Next
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.GetClientRect Left, Top, Right, Bottom
    treeSolutionView.Move Left, Top, Right - Left, Bottom - Top
End Sub

Private Sub Form_Activate()
    CommandBars.VisualTheme = fMainForm.CommandBars.VisualTheme
    CommandBars.PaintManager.RefreshMetrics
    CommandBars.RecalcLayout
End Sub

Private Sub Form_Load()
    Dim Root As TreeViewNode, Node As TreeViewNode, Solution As TreeViewNode
    
    Set treeSolutionView.Icons = SuiteControlsGlobalSettings.Icons
      
    Set Solution = treeSolutionView.Nodes.Add(, , , "Solution 'XtremeDemo': 1 project(s)", ID_SOLUTION_TREE_1)
    Solution.Expanded = True
    
    Set Root = treeSolutionView.Nodes.Add(Solution, xtpTreeViewChild, , "XtremeDemo Files", ID_SOLUTION_TREE_2)
    Root.Bold = True
    
    Set Node = treeSolutionView.Nodes.Add(Root, xtpTreeViewChild, , "Source Files", ID_SOLUTION_TREE_3)
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "ChildFrm.cpp", ID_SOLUTION_TREE_5
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "MainFrm.cpp", ID_SOLUTION_TREE_5
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "StdAfx.cpp", ID_SOLUTION_TREE_5
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemo.cpp", ID_SOLUTION_TREE_5
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemo.rc", ID_SOLUTION_TREE_5
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemoDoc.cpp", ID_SOLUTION_TREE_5
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemoView.cpp", ID_SOLUTION_TREE_5
    Node.Expanded = True

    Set Node = treeSolutionView.Nodes.Add(Root, xtpTreeViewChild, , "Header Files", ID_SOLUTION_TREE_3)
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "ChildFrm.h", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "MainFrm.h", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "StdAfx.h", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemo.h", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemoDoc.h", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemoView.h", ID_SOLUTION_TREE_6
    
    Set Node = treeSolutionView.Nodes.Add(Root, xtpTreeViewChild, , "Resource Files", ID_SOLUTION_TREE_3)
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "Toolbar.bmp", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemo.ico", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemo.rc2", ID_SOLUTION_TREE_6
    treeSolutionView.Nodes.Add Node, xtpTreeViewChild, , "XtremeDemoDoc.ico", ID_SOLUTION_TREE_6
    
    Set Node = treeSolutionView.Nodes.Add(Root, xtpTreeViewChild, , "ReadMe.txt", ID_SOLUTION_TREE_6)
    Set Node = treeSolutionView.Nodes.Add(Root, xtpTreeViewChild, , "External Dependencies", ID_SOLUTION_TREE_3)
    
    Root.Expanded = True
    Solution.Expanded = True
    
    Dim ControlOpen As CommandBarButton

    With CommandBars.ActiveMenuBar.Controls
        Set ControlOpen = .Add(xtpControlButton, ID_VIEW_PROPERTIESWINDOW, "Properties")
    End With
    
    CommandBars.ActiveMenuBar.EnableDocking xtpFlagStretched
    CommandBars.ActiveMenuBar.ModifyStyle XTP_CBRS_GRIPPER, 0

End Sub

Public Sub SetTheme(nTheme As AppearanceStyle)
    SetTreeTheme treeSolutionView, nTheme
End Sub
