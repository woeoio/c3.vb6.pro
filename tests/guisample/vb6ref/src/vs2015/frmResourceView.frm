VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Controls.ocx"
Begin VB.Form frmResourceView 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   FillColor       =   &H80000010&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TreeView treeResourceView 
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
      _Version        =   1048577
      _ExtentX        =   7646
      _ExtentY        =   3201
      _StockProps     =   77
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ShowLines       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      IconSize        =   16
   End
End
Attribute VB_Name = "frmResourceView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Root As TreeViewNode
    Dim Node As TreeViewNode
    
    Set treeResourceView.Icons = SuiteControlsGlobalSettings.Icons
  
    Set Root = treeResourceView.Nodes.Add(, , , "XtremeDemo resources", ID_RESOURCE_TREE_1)
    
    Root.Bold = True
    
    Set Node = treeResourceView.Nodes.Add(Root, xtpTreeViewChild, , "Accelerator", ID_RESOURCE_TREE_1)
    
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "IDR_MAINFRAME", ID_RESOURCE_TREE_3
    Node.Expanded = True
    
    Set Node = treeResourceView.Nodes.Add(Root, xtpTreeViewChild, , "Dialog", ID_RESOURCE_TREE_1)
    
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "IDD_ABOUTBOX", ID_RESOURCE_TREE_4
    Node.Expanded = True
    
    Set Node = treeResourceView.Nodes.Add(Root, xtpTreeViewChild, , "Icon", ID_RESOURCE_TREE_1)
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "IDR_MAINFRAME", ID_RESOURCE_TREE_5
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "IDR_XTREMETYPE", ID_RESOURCE_TREE_5
    Node.Expanded = True

    Set Node = treeResourceView.Nodes.Add(Root, xtpTreeViewChild, , "Menu", ID_RESOURCE_TREE_1)
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "IDR_MAINFRAME", ID_RESOURCE_TREE_6
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "IDR_XTREMETYPE", ID_RESOURCE_TREE_6

    Set Node = treeResourceView.Nodes.Add(Root, xtpTreeViewChild, , "String Table", ID_RESOURCE_TREE_1)
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "String Table", ID_RESOURCE_TREE_7

    Set Node = treeResourceView.Nodes.Add(Root, xtpTreeViewChild, , "Toolbar", ID_RESOURCE_TREE_1)
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "IDR_MAINFRAME", ID_RESOURCE_TREE_8

    Set Node = treeResourceView.Nodes.Add(Root, xtpTreeViewChild, , "Version", ID_RESOURCE_TREE_1)
    treeResourceView.Nodes.Add Node, xtpTreeViewChild, , "VS_VERSION_INFO", ID_RESOURCE_TREE_9
    
    Root.Expanded = True

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    treeResourceView.Move Me.ScaleLeft, Me.ScaleTop, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Sub SetTheme(nTheme As AppearanceStyle)
    SetTreeTheme treeResourceView, nTheme
End Sub
