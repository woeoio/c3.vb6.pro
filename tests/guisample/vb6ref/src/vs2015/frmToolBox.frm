VERSION 5.00
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "TaskPanel.ocx"
Begin VB.Form frmToolBox 
   BorderStyle     =   0  'None
   ClientHeight    =   7170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeTaskPanel.TaskPanel wndToolBox 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _Version        =   1048577
      _ExtentX        =   7858
      _ExtentY        =   12303
      _StockProps     =   64
      VisualTheme     =   14
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Set wndToolBox.Icons = TaskPanelGlobalSettings.Icons
    CreateToolBox
End Sub

Function CreateToolboxGroup(Caption As String) As TaskPanelGroup
    Dim Folder As TaskPanelGroup, Pointer As TaskPanelGroupItem
    
    Set Folder = wndToolBox.Groups.Add(0, Caption)
    Set Pointer = Folder.Items.Add(0, ID_TOOLBOXITEM_POINTER, xtpTaskItemTypeLink, 0)
    Pointer.AllowDrag = False
    Pointer.AllowDrop = False
    Pointer.Selected = True
    Folder.IconIndex = 100
    
    Set CreateToolboxGroup = Folder
End Function

Function CreateToolboxItem(Group As TaskPanelGroup, Caption As String, IconIndex As Long) As TaskPanelGroup
    Group.Items.Add IconIndex + 1, Caption, xtpTaskItemTypeLink, IconIndex
End Function

Sub CreateToolBox()
    
    Dim FolderData As TaskPanelGroup, FolderPropertyPanes As TaskPanelGroup, _
        FolderComponents As TaskPanelGroup, FolderWindowsForms As TaskPanelGroup

    Set FolderPropertyPanes = CreateToolboxGroup(ID_TOOLBOXFOLDER_PROPERTYPANES)

    Set FolderData = CreateToolboxGroup(ID_TOOLBOXFOLDER_DATA)
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_DATASET, 26
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_OLEDBDATAADAPTER, 27
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_OLEDBCONNECTION, 28
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_OLEDBCOMMAND, 29
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_SQLDATAADAPTER, 30
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_SQLCONNECTION, 31
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_SQLCOMMAND, 32
    CreateToolboxItem FolderData, ID_TOOLBOXITEM_DATAVIEW, 33

    Set FolderComponents = CreateToolboxGroup(ID_TOOLBOXFOLDER_COMPONENTS)
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_FILESYSTEMWATCHER, 34
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_EVENTLOG, 35
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_DIRECTORYENTRY, 36
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_DIRECTORYSEARCHER, 37
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_MESSAGEQUEUE, 38
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_PERFORMANCECOUNTER, 39
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_PROCESS, 40
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_SERVICECONTROLLER, 41
    CreateToolboxItem FolderComponents, ID_TOOLBOXITEM_TIMER, 42
    
    Set FolderWindowsForms = CreateToolboxGroup(ID_TOOLBOXFOLDER_WINDOWSFORMS)
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_BUTTON, 1
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_CHECK_BOX, 2
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_EDIT_CONTROL, 3
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_COMBO_BOX, 4
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_LIST_BOX, 5
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_GROUP_BOX, 6
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_RADIO_BUTTON, 7
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_STATIC_TEXT, 8
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_PICTURE_CONTROL, 9
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_HORIZONTAL_SCROLL_BAR, 10
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_VERTICAL_SCROLL_BAR, 11
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_SLIDER_CONTROL, 12
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_SPIN_CONTROL, 13
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_PROGRESS_CONTROL, 14
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_HOT_KEY, 15
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_LIST_CONTROL, 16
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_TREE_CONTROL, 17
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_TAB_CONTROL, 18
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_ANIMATION_CONTROL, 19
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_RICH_EDIT_CONTROL, 20
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_DATE_TIME_PICKER, 21
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_MONTH_CALENDAR_CONTROL, 22
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_IP_ADDRESS_CONTROL, 23
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_EXTENDED_COMBO_BOX, 24
    CreateToolboxItem FolderWindowsForms, ID_TOOLBOXITEM_CUSTOM_CONTROL, 25

    CreateToolboxGroup (ID_TOOLBOXFOLDER_CLIPBOARDRING)
    CreateToolboxGroup (ID_TOOLBOXFOLDER_GENERAL)
         
    wndToolBox.Behaviour = xtpTaskPanelBehaviourExplorer
    
    wndToolBox.AllowDrag = True
    wndToolBox.SelectItemOnFocus = True
    wndToolBox.HotTrackStyle = xtpTaskPanelHighlightItem
    wndToolBox.SingleSelection = True
    FolderPropertyPanes.Expanded = True
    
    FolderComponents.Expanded = True
    
    wndToolBox.Reposition
        
End Sub

Private Sub Form_Resize()
    wndToolBox.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

