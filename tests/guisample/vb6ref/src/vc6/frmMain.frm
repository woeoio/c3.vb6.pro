VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#20.3#0"; "Codejock.CommandBars.v20.3.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#20.3#0"; "Codejock.DockingPane.v20.3.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Visual Studio Sample"
   ClientHeight    =   8790
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   14220
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   1080
      Top             =   1320
      _Version        =   1245187
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   14
   End
   Begin XtremeDockingPane.DockingPane DockingPaneManager 
      Left            =   1080
      Top             =   840
      _Version        =   1245187
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   16
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents Workspace As TabWorkspace
Attribute Workspace.VB_VarHelpID = -1
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Dim frmPropertyGrid As frmPropertyGrid
Dim frmClassView As frmClassView
Dim frmSolutionView As frmSolutionView
Dim frmResourceView As frmResourceView
Dim frmOutput As frmOutput
Dim frmToolBox As frmToolBox

Const ID_INDICATOR_CAPS = 59137
Const ID_INDICATOR_NUM = 59138
Const ID_INDICATOR_SCRL = 59139

Const IMAGEBASE = 10000

Dim MAINFRAME             As String
Dim TOOLBAR_WEB           As String
Dim TOOLBAR_EXT           As String
Dim PANE_PROPERTIES       As String
Dim CLASSTREE             As String
Dim FILETREE              As String
Dim RESTREE               As String
Dim PANE_SOLUTIONEXPLORER As String
Dim TOOLBOXICONS          As String
Dim TOOLBAR_THEME         As String
Dim TOOLBAR_STATE         As String
Dim TOOLBAR_FULLSCREEN    As String

Dim AppState As Long

Private Sub MDIForm_Load()

    PropertyGridTheme = xtpGridThemeVisualStudio2015
    TaskpanelTheme = xtpTaskPanelThemeVisualStudio2015
    ControlTheme = xtpAppearanceVisualStudio2015
    AppTheme = ID_THEME_VC14_LIGHT
    AppState = ID_STATE_DEFAULT
    strSchema = ""
    strColor = ""
            
    CommandBarsGlobalSettings.App = App

    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    Dim Control As CommandBarControl
    Dim Combo As CommandBarComboBox
    Dim ControlFile As CommandBarPopup, ControlEdit As CommandBarPopup, ControlView As CommandBarPopup
    Dim ControlNew As CommandBarPopup, ControlOpen As CommandBarPopup, ControlFind As CommandBarPopup, ControlBrowser As CommandBarPopup, ControlToolbars  As CommandBarPopup
    Dim ControlProject As CommandBarPopup, ControlBuild As CommandBarPopup, ControlDebug As CommandBarPopup, ControlWindows As CommandBarPopup, ControlTools As CommandBarPopup
    Dim ControlWindow As CommandBarPopup, ControlHelp As CommandBarPopup
    
    Set ControlFile = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&File")
    With ControlFile.CommandBar.Controls
        
        Set ControlNew = .Add(xtpControlPopup, 0, "New")
        ControlNew.IconId = ID_FILE_NEW_PROJECT
        ControlNew.CommandBar.Controls.Add xtpControlButton, ID_FILE_NEW_PROJECT, "New Project"
        ControlNew.CommandBar.Controls.Add xtpControlButton, ID_FILE_NEW, "New File..."
        ControlNew.CommandBar.Controls.Add xtpControlButton, ID_FILE_NEW_BLANK, "Blank Solution"
        
        Set ControlOpen = .Add(xtpControlPopup, 0, "Open")
        ControlOpen.IconId = ID_PROJECT_ADDNEWITEM
        ControlOpen.CommandBar.Controls.Add xtpControlButton, ID_FILE_OPEN_PROJECT, "Open Project"
        Set Control = ControlOpen.CommandBar.Controls.Add(xtpControlButton, ID_FILE_OPEN, "Open File...")
        
        .Add xtpControlButton, ID_FILE_CLOSE, "&Close"
        .Add xtpControlButton, ID_FILE_SAVE, "&Save"
        .Add xtpControlButton, ID_FILE_SAVE_AS, "Save &As"
        
        Set Control = .Add(xtpControlButton, ID_FILE_OPEN_SOLUTION, "Op&en Solution...")
        Control.BeginGroup = True
        .Add xtpControlButton, ID_FILE_CLOSE_SOLUTION, "Close Solu&tion..."
        
        Set Control = .Add(xtpControlButton, ID_FILE_PRINT, "Print...")
        Control.BeginGroup = True
        .Add xtpControlButton, ID_FILE_PRINT_SETUP, "Print Set&up..."
        
        Set Control = .Add(xtpControlButton, ID_FILE_EXIT, "&Exit")
        Control.BeginGroup = True
    End With
    
    Set ControlEdit = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Edit")
    With ControlEdit.CommandBar.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_UNDO, "&Undo")
        
        Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "Cu&t")
        Control.BeginGroup = True
        
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "&Copy")
        
        Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "&Paste")
        
        .Add xtpControlButton, 0, "Delete"
        
        Set Control = .Add(xtpControlButton, 0, "Select &All")
        Control.BeginGroup = True
        
        Set ControlFind = .Add(xtpControlPopup, 0, "Find And Replace")
        ControlFind.CommandBar.Controls.Add xtpControlButton, 0, "Find..."
        ControlFind.CommandBar.Controls.Add xtpControlButton, 0, "Replace..."
        ControlFind.CommandBar.Controls.Add xtpControlButton, 0, "Find in Files"
    End With

    Set ControlView = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&View")
    With ControlView.CommandBar.Controls
        .Add xtpControlButton, 0, "Open"
        
        Set Control = .Add(xtpControlButton, ID_VIEW_SOLUTIONEXPLORER, "Solution &Explorer")
        Control.BeginGroup = True
        .Add xtpControlButton, ID_VIEW_CLASSVIEW, "Class View"
        .Add xtpControlButton, ID_VIEW_RESOURCEVIEW, "&Resource View"
        .Add xtpControlButton, ID_VIEW_PROPERTIESWINDOW, "&Properties Window"
        .Add xtpControlButton, ID_VIEW_TOOLBOX, "Toolbo&x"
        .Add xtpControlButton, ID_VIEW_WELCOME_PAGE, "Welcome Page"
        
        Set ControlBrowser = .Add(xtpControlPopup, 0, "Web Browser")
        ControlBrowser.CommandBar.Controls.Add xtpControlButton, ID_BROWSER_BACK, "Back"
        ControlBrowser.CommandBar.Controls.Add xtpControlButton, ID_BROWSER_FORWARD, "Forward"
        ControlBrowser.CommandBar.Controls.Add xtpControlButton, ID_BROWSER_HOME, "Home"
        ControlBrowser.CommandBar.Controls.Add xtpControlButton, ID_BROWSER_SEARCH, "Search"
        
        Set ControlToolbars = .Add(xtpControlPopup, 0, "Toolbars")
        ControlToolbars.BeginGroup = True
        ControlToolbars.CommandBar.Controls.Add xtpControlButton, XTP_ID_TOOLBARLIST, "Toolbar List"
        .Add xtpControlButton, ID_VIEW_FULLSCREEN, "Full Screen"
        
    End With
    
    Set ControlProject = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Project")
    With ControlProject.CommandBar.Controls
        .Add xtpControlButton, ID_PROJECT_ADD_CLASS, "Add Class..."
        .Add xtpControlButton, 0, "Add Resource..."
        
        Set Control = .Add(xtpControlButton, ID_PROJECT_ADDNEWITEM, "Add Ne&w Item...")
        Control.BeginGroup = True
        .Add xtpControlButton, ID_PROJECT_ADDEXISTINGITEM, "Add Existin&g Item..."
        
        Set Control = .Add(xtpControlButton, 0, "New Folder")
        Control.BeginGroup = True
    End With
    
    Set ControlBuild = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Build")
    With ControlBuild.CommandBar.Controls
        .Add xtpControlButton, ID_BUILD_SOLUTION, "&Build Solution"
        .Add xtpControlButton, 0, "&Rebuild Solution"
        .Add xtpControlButton, 0, "&Clean Solution"
        
        Set Control = .Add(xtpControlButton, ID_BUILD_COMPILE, "Compile...")
        Control.BeginGroup = True
    End With
    
    Set ControlDebug = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Debug")
    With ControlDebug.CommandBar.Controls
        Set ControlWindows = .Add(xtpControlPopup, 0, "Windows")
        ControlWindows.CommandBar.Controls.Add xtpControlButton, ID_WINDOWS_OUTPUT, "Output"
        .Add xtpControlButton, ID_DEBUG_START, "&Start"
        .Add xtpControlButton, ID_DEBUG_WITHOUT_DEBUGGING, "Start Without Debugging"
        .Add xtpControlButton, 0, "&Processes..."
        .Add xtpControlButton, 0, "E&xceptions..."
        
        Set Control = .Add(xtpControlButton, 0, "Step Into")
        Control.BeginGroup = True
        .Add xtpControlButton, 0, "Step Over"
    End With
    
    Set ControlTools = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Tools")
    With ControlTools.CommandBar.Controls
        .Add xtpControlButton, 0, "Options..."
    End With
    
    Set ControlWindow = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Window")
    With ControlWindow.CommandBar.Controls
        .Add xtpControlButton, ID_WINDOW_NEW, "&New Window"
        Set Control = .Add(xtpControlButton, ID_WINDOW_HIDEALL, "A&uto Hide All")
        Control.BeginGroup = True
        .Add xtpControlButton, 35002, "&Workspace Actions"
        .Add xtpControlButton, 35000, "&Window List"
    End With

    Set ControlHelp = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&Help")
    With ControlHelp.CommandBar.Controls
        .Add xtpControlButton, ID_HELP_DYNAMIC, "&Dynamic Help"
        Set Control = .Add(xtpControlButton, 0, "&Contents...")
        Control.BeginGroup = True
        .Add xtpControlButton, 0, "&Index..."
        .Add xtpControlButton, 0, "&Search..."
        
        Set Control = .Add(xtpControlButton, ID_HELP_ABOUT, "&About Visual Studio Sample")
        Control.BeginGroup = True
    End With
    
    Dim StandardBar As CommandBar, WebBar As CommandBar
    Dim ControlItem As CommandBarPopup
    
    Set StandardBar = CommandBars.Add("Standard", xtpBarTop)
    With StandardBar.Controls
        Set ControlNew = .Add(xtpControlSplitButtonPopup, ID_FILE_NEW_PROJECT, "")
        ControlNew.CommandBar.Controls.Add xtpControlButton, ID_FILE_NEW_PROJECT, "New Project..."
        ControlNew.CommandBar.Controls.Add xtpControlButton, ID_FILE_NEW_BLANK, "New Blank Solution..."
        
        Set ControlItem = .Add(xtpControlSplitButtonPopup, ID_PROJECT_ADDNEWITEM, "")
        ControlItem.CommandBar.Controls.Add xtpControlButton, ID_PROJECT_ADDNEWITEM, "Add New Item"
        ControlItem.CommandBar.Controls.Add xtpControlButton, ID_PROJECT_ADDEXISTINGITEM, "Add Existin&g Item..."
        ControlItem.CommandBar.Controls.Add xtpControlButton, ID_PROJECT_ADD_CLASS, "Add Class..."
        
        .Add xtpControlButton, ID_FILE_OPEN, "Open File..."
        .Add xtpControlButton, ID_FILE_SAVE, "Save"
        
        Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "Cu&t")
        Control.BeginGroup = True
        
        .Add xtpControlButton, ID_EDIT_COPY, "&Copy"
        .Add xtpControlButton, ID_EDIT_PASTE, "&Paste"
        
        Set Control = .Add(xtpControlButton, ID_DEBUG_START, "Start")
        Control.BeginGroup = True
        Set Combo = .Add(xtpControlComboBox, ID_DEBUG_CONFIGURATION, "Configuration")
        Combo.AddItem "Debug"
        Combo.AddItem "Release"
        Combo.AddItem "Configuration Manager"
        Combo.DropDownWidth = 150
        Debug.Print Combo.List(2)
        Debug.Print Combo.ListCount
        Combo.ListIndex = 2
        
        
        
        Set Control = .Add(xtpControlButton, ID_FINDANDREPLACE_FINDINFILES, "Find")
        Control.BeginGroup = True
        Set Combo = .Add(xtpControlComboBox, 0, "Find")
        Combo.DropDownListStyle = True
        Combo.AddItem "www.codejock.com"
        Combo.Width = 180
        
        Set Control = .Add(xtpControlButton, ID_VIEW_SOLUTIONEXPLORER, "Solution &Explorer")
        Control.BeginGroup = True
        
        .Add xtpControlButton, ID_VIEW_PROPERTIESWINDOW, "&Properties Window"
        .Add xtpControlButton, ID_VIEW_TOOLBOX, "Toolbo&x"
        .Add xtpControlButton, ID_VIEW_CLASSVIEW, "Class View"
        
    End With
    
    Set WebBar = CommandBars.Add("Web", xtpBarTop)
    With WebBar.Controls
        .Add xtpControlButton, ID_BROWSER_BACK, "Back"
        .Add xtpControlButton, ID_BROWSER_FORWARD, "Forward"
        .Add xtpControlButton, ID_BROWSER_STOP, "Stop"
        .Add xtpControlButton, ID_BROWSER_REFRESH, "Refresh"
        .Add xtpControlButton, ID_BROWSER_HOME, "Home"
        .Add xtpControlButton, ID_BROWSER_SEARCH, "Search"
        .Add xtpControlButton, ID_BROWSER_FAVORITES, "Favorites"
        .Add xtpControlButton, ID_BROWSER_ADDTOFAVORITE, "Add To Favorite"
        
        Set Combo = .Add(xtpControlComboBox, ID_BROWSER_URL, "URL")
        Combo.BeginGroup = True
        Combo.AddItem "ms-help://MS.VSCC/vsccCommon/html/vsccCommon_Start_Page.htm"
        Combo.Width = 200
        Combo.DropDownWidth = 300
        Combo.DropDownListStyle = True
                
        Set Control = .Add(xtpControlButton, ID_BROWSER_SYNC, "Sync")
        Control.BeginGroup = True
        .Add xtpControlButton, ID_BROWSER_PREV, "Prev"
        .Add xtpControlButton, ID_BROWSER_NEXT, "Next"
        .Add xtpControlButton, ID_BROWSER_FONT, "Fonts"
    End With
       
    CommandBars.Actions.Add ID_THEME_VC6, "Visual Studio 6.0", "Visual Studio 6.0", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC7, "Visual Studio 2002", "Visual Studio 2002", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC8, "Visual Studio 2005", "Visual Studio 2005", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC9, "Visual Studio 2008", "Visual Studio 2008", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC10, "Visual Studio 2010", "Visual Studio 2010", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC11_LIGHT, "Visual Studio 2012 Light", "Visual Studio 2012 Light", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC11_DARK, "Visual Studio 2012 Dark", "Visual Studio 2012 Dark", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC14_LIGHT, "Visual Studio 2015 Light", "Visual Studio 2015 Light", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC14_DARK, "Visual Studio 2015 Dark", "Visual Studio 2015 Dark", "", "Themes"
    CommandBars.Actions.Add ID_THEME_VC14_BLUE, "Visual Studio 2015 Blue", "Visual Studio 2015 Blue", "", "Themes"
    
    Dim ThemesBar As CommandBar
    Set ThemesBar = CommandBars.Add("Themes", xtpBarTop)
            
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC6
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC7
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC8
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC9
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC10
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC11_LIGHT
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC11_DARK
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC14_LIGHT
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC14_DARK
    AddButton ThemesBar.Controls, xtpControlButton, ID_THEME_VC14_BLUE
            
    'CommandBars.Options.ShowExpandButtonAlways = False
    
    CommandBars.KeyBindings.Add FCONTROL, Asc("P"), ID_FILE_PRINT
    CommandBars.KeyBindings.Add FCONTROL, Asc("O"), ID_FILE_OPEN
    CommandBars.KeyBindings.Add FCONTROL, Asc("S"), ID_FILE_SAVE
    CommandBars.KeyBindings.Add FCONTROL, Asc("X"), ID_EDIT_CUT
    CommandBars.KeyBindings.Add FCONTROL, Asc("C"), ID_EDIT_COPY
    CommandBars.KeyBindings.Add FCONTROL, Asc("V"), ID_EDIT_PASTE
    CommandBars.KeyBindings.Add 0, VK_F4, ID_VIEW_PROPERTIESWINDOW
    
    CommandBars.Actions.Add ID_STATE_DEFAULT, "Default", "Default", "", "State"
    CommandBars.Actions.Add ID_STATE_ORANGE, "Orange", "Orange", "", "State"
    CommandBars.Actions.Add ID_STATE_BLUE, "Blue", "Blue", "", "State"
    
    Dim StateBar As CommandBar
    Set StateBar = CommandBars.Add("State", xtpBarTop)
    AddButton StateBar.Controls, xtpControlButton, ID_STATE_DEFAULT
    AddButton StateBar.Controls, xtpControlButton, ID_STATE_ORANGE
    AddButton StateBar.Controls, xtpControlButton, ID_STATE_BLUE
    
    CommandBars.EnableCustomization (True)
    
    Set Workspace = CommandBars.ShowTabWorkspace(True)
    Workspace.EnableGroups
    Workspace.Flags = xtpWorkspaceHideArrowsAlways Or xtpWorkspaceHideClose Or xtpWorkspaceShowActiveFiles Or xtpWorkspaceShowCloseSelectedTab
        
    Dim StatusBar As XtremeCommandBars.IStatusBar
    Set StatusBar = CommandBars.StatusBar
    StatusBar.Visible = True
    
    StatusBar.AddPane 0
    StatusBar.AddPane ID_INDICATOR_CAPS
    StatusBar.AddPane ID_INDICATOR_NUM
    StatusBar.AddPane ID_INDICATOR_SCRL
    
    DockBarRightOf ThemesBar, WebBar
    DockBarRightOf StateBar, ThemesBar
     
    'Important!!!
    DockingPaneManager.SetCommandBars Me.CommandBars
    
    Dim PaneClassView As Pane, PaneResource As Pane, PaneSolution As Pane, PaneOutput As Pane, PaneProperties As Pane, PaneToolBox As Pane
    
    Set PaneOutput = DockingPaneManager.CreatePane(ID_PANE_OUTPUT, 150, 120, DockBottomOf, Nothing)
    PaneOutput.Title = "Output"
    
    Set PaneClassView = DockingPaneManager.CreatePane(ID_PANE_CLASSVIEW, 210, 140, DockRightOf, Nothing)
    PaneClassView.Title = "Class View"
    
    Set PaneSolution = DockingPaneManager.CreatePane(ID_PANE_SOLUTIONVIEW, 210, 140, DockLeftOf, Nothing)
    PaneSolution.Title = "Solution Explorer"
    PaneSolution.AttachTo PaneClassView
    
    Set PaneResource = DockingPaneManager.CreatePane(ID_PANE_RESOURCEVIEW, 210, 140, DockLeftOf, Nothing)
    PaneResource.Title = "Resource View"
    PaneResource.AttachTo PaneClassView
    
    PaneClassView.Select
        
    Set PaneProperties = DockingPaneManager.CreatePane(ID_PANE_PROPERTIES, 210, 140, DockBottomOf, PaneClassView)
    PaneProperties.Title = "Properties"
    
    Set PaneToolBox = DockingPaneManager.CreatePane(ID_PANE_TOOLBOX, 210, 140, DockLeftOf, Nothing)
    PaneToolBox.Title = "ToolBox"
    'PaneToolBox.Hide
    
    DockingPaneManager.Options.ThemedFloatingFrames = True
    DockingPaneManager.Options.ClientMargin = 6
        
    DockingPaneManager.Icons = CommandBarsGlobalSettings.Icons
    
    CommandBars.FindControl(, ID_THEME_VC6, , True).Execute
    
    LoadNewDoc
    WritePaneProbe
End Sub

' 本轮取证专用探针 v2 (仅副本; 厂商目录不动). 两侧同一份代码, 同一输出格式.
' v1 用 Panes 集合: 原生 VB6 直接 438 —— Codejock COM 里没有 Panes, 只有 AddPane/FindPane.
' v2 改用 StatusBar.FindPane(id), 逐 pane 打印 Text/Width/Left, 每行独立免得一行失败吞掉整块.
Private Sub WritePaneProbe()
    On Error Resume Next
    Dim fn As Integer, i As Long
    Dim ids(1 To 4) As Long
    ids(1) = 0
    ids(2) = ID_INDICATOR_CAPS
    ids(3) = ID_INDICATOR_NUM
    ids(4) = ID_INDICATOR_SCRL
    Dim pn As Object
    fn = FreeFile
    Open App.Path & "\panes.txt" For Output As #fn
    Print #fn, "SBVIS=" & CStr(CommandBars.StatusBar.Visible) & " err=" & CStr(Err.Number)
    Err.Clear
    For i = 1 To 4
        Set pn = CommandBars.StatusBar.FindPane(ids(i))
        Print #fn, "P" & CStr(i) & "ID=" & CStr(ids(i)) & " ISNOTHING=" & CStr(pn Is Nothing) & " err=" & CStr(Err.Number)
        Err.Clear
        If Not (pn Is Nothing) Then
            Print #fn, "P" & CStr(i) & " t=[" & pn.Text & "]"
            Print #fn, "P" & CStr(i) & " w=" & CStr(pn.Width) & " l=" & CStr(pn.Left) & " v=" & CStr(pn.Visible)
            Err.Clear
        End If
    Next i
    ' 本轮 A2/B3 取证: Visible 的 set/get 是否真的落地 (逐次读回).
    Dim sb2 As XtremeCommandBars.IStatusBar
    Set sb2 = CommandBars.StatusBar
    Dim sbh As Long
    sbh = 0
    sbh = sb2.hWnd
    Print #fn, "HWND=" & CStr(sbh) & " err=" & CStr(Err.Number)
    Err.Clear
    Print #fn, "VIS_PRE=" & CStr(sb2.Visible) & " err=" & CStr(Err.Number) & " WVIS=" & CStr(IsWindowVisible(sbh))
    Err.Clear
    sb2.Visible = False
    Print #fn, "VIS_SET0=" & CStr(sb2.Visible) & " err=" & CStr(Err.Number) & " WVIS=" & CStr(IsWindowVisible(sbh))
    Err.Clear
    sb2.Visible = True
    Print #fn, "VIS_SET1=" & CStr(sb2.Visible) & " err=" & CStr(Err.Number) & " WVIS=" & CStr(IsWindowVisible(sbh))
    Print #fn, "CBT=" & CStr(CommandBars.VisualTheme)
    Err.Clear
    Close #fn
End Sub

Private Sub DockingPaneManager_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case ID_PANE_PROPERTIES:
            If frmPropertyGrid Is Nothing Then
                Set frmPropertyGrid = New frmPropertyGrid
                frmPropertyGrid.wndPropertyGrid.VisualTheme = PropertyGridTheme
            End If
            Item.Handle = frmPropertyGrid.hwnd
            
        Case ID_PANE_CLASSVIEW:
            If frmClassView Is Nothing Then
                Set frmClassView = New frmClassView
                frmClassView.SetTheme ControlTheme
            End If
            Item.Handle = frmClassView.hwnd
            
        Case ID_PANE_SOLUTIONVIEW:
            If frmSolutionView Is Nothing Then
                Set frmSolutionView = New frmSolutionView
                frmSolutionView.SetTheme ControlTheme
            End If
            Item.Handle = frmSolutionView.hwnd
            
        Case ID_PANE_RESOURCEVIEW:
            If frmResourceView Is Nothing Then
                Set frmResourceView = New frmResourceView
                frmResourceView.SetTheme ControlTheme
            End If
            Item.Handle = frmResourceView.hwnd
            
        Case ID_PANE_OUTPUT
            If frmOutput Is Nothing Then
                Set frmOutput = New frmOutput
                frmOutput.SetTheme ControlTheme
            End If
            Item.Handle = frmOutput.hwnd
            
        Case ID_PANE_TOOLBOX
            If frmToolBox Is Nothing Then
                Set frmToolBox = New frmToolBox
                frmToolBox.wndToolBox.VisualTheme = TaskpanelTheme
            End If
            Item.Handle = frmToolBox.hwnd
                      
    End Select
End Sub

Private Sub Workspace_RClick(ByVal Item As XtremeCommandBars.ITabControlItem)
    If Not Item Is Nothing Then
        Debug.Print Item.Caption
        Item.Selected = True
        Workspace.Refresh
        
        Dim Popup As CommandBar
        Dim Control As CommandBarControl
        
        Set Popup = CommandBars.Add("Popup", xtpBarPopup)
        
        With Popup.Controls
            .Add xtpControlButton, ID_FILE_NEW, "&New", -1, False
            .Add xtpControlButton, ID_FILE_CLOSE, "Close", -1, False
            Set Control = .Add(xtpControlButton, 35002, "&Workspace Actions")
            Control.BeginGroup = True
        End With
    
        Popup.ShowPopup
        Workspace.Refresh
    End If
End Sub

Private Sub CommandBars_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)

    On Error Resume Next
    Select Case Control.Id
        Case ID_EDIT_CUT:
            If ActiveForm Is Nothing Then
                Control.Enabled = False
            ElseIf TypeName(ActiveForm) = "frmDocument" Then
                Control.Enabled = Not ActiveForm.rtfText.SelLength = 0
            Else
                Control.Enabled = False
            End If
        Case ID_EDIT_COPY:
            If ActiveForm Is Nothing Then
                Control.Enabled = False
            ElseIf TypeName(ActiveForm) = "frmDocument" Then
                Control.Enabled = Not ActiveForm.rtfText.SelLength = 0
            Else
                Control.Enabled = False
            End If
        Case ID_EDIT_PASTE:
            Control.Enabled = Not ActiveForm Is Nothing
        Case ID_THEME_VC8:
            Control.Checked = CommandBars.VisualTheme = xtpThemeWhidbey
        Case ID_THEME_VC9:
            Control.Checked = CommandBars.VisualTheme = xtpThemeVisualStudio2008
        Case ID_THEME_VC10:
            Control.Checked = CommandBars.VisualTheme = xtpThemeVisualStudio2010
        Case ID_THEME_VC7:
            Control.Checked = CommandBars.VisualTheme = xtpThemeOfficeXP
        Case ID_THEME_VC6:
            Control.Checked = CommandBars.VisualTheme = xtpThemeVisualStudio6
        Case ID_THEME_VC11_LIGHT:
            Control.Checked = CommandBars.VisualTheme = xtpThemeVisualStudio2012Light
        Case ID_THEME_VC11_DARK:
            Control.Checked = CommandBars.VisualTheme = xtpThemeVisualStudio2012Dark
        Case ID_THEME_VC14_LIGHT:
            Control.Checked = AppTheme = ID_THEME_VC14_LIGHT
        Case ID_THEME_VC14_DARK:
            Control.Checked = AppTheme = ID_THEME_VC14_DARK
        Case ID_THEME_VC14_BLUE:
            Control.Checked = AppTheme = ID_THEME_VC14_BLUE
        Case ID_STATE_DEFAULT:
            Control.Enabled = (AppTheme = ID_THEME_VC14_LIGHT Or AppTheme = ID_THEME_VC14_DARK Or AppTheme = ID_THEME_VC14_BLUE)
            Control.Checked = Control.Enabled And (AppState = ID_STATE_DEFAULT)
        Case ID_STATE_ORANGE:
            Control.Enabled = (AppTheme = ID_THEME_VC14_LIGHT Or AppTheme = ID_THEME_VC14_DARK Or AppTheme = ID_THEME_VC14_BLUE)
            Control.Checked = Control.Enabled And (AppState = ID_STATE_ORANGE)
        Case ID_STATE_BLUE:
            Control.Enabled = (AppTheme = ID_THEME_VC14_LIGHT Or AppTheme = ID_THEME_VC14_DARK Or AppTheme = ID_THEME_VC14_BLUE)
            Control.Checked = Control.Enabled And (AppState = ID_STATE_BLUE)
    End Select
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    On Error Resume Next
    Dim Combo As CommandBarComboBox
    
    If Control.Id >= ID_THEME_VC6 And Control.Id <= ID_THEME_VC14_BLUE Then
        AppTheme = Control.Id
        strColor = ""
        strSchema = ""
    End If
    
    Select Case Control.Id
        Case ID_HELP_ABOUT:    CommandBars.AboutBox
        Case ID_DEBUG_CONFIGURATION:
            Set Combo = Control
            If Not Combo Is Nothing Then
                If Combo.ListIndex > 0 Then MsgBox Combo.List(Combo.ListIndex)
            End If
        Case ID_BROWSER_URL
            Set Combo = Control
            If Not Combo Is Nothing Then
                If Not Combo.Text = "" Then
                    MsgBox Combo.Text
                    If Combo.FindItem(Combo.Text) = 0 Then Combo.AddItem Combo.Text
                End If
            End If
            
        Case ID_VIEW_CLASSVIEW: DockingPaneManager.ShowPane ID_PANE_CLASSVIEW
        Case ID_VIEW_RESOURCEVIEW: DockingPaneManager.ShowPane ID_PANE_RESOURCEVIEW
        Case ID_VIEW_SOLUTIONEXPLORER: DockingPaneManager.ShowPane ID_PANE_SOLUTIONVIEW
        Case ID_VIEW_TOOLBOX: DockingPaneManager.ShowPane ID_PANE_TOOLBOX
        Case ID_VIEW_PROPERTIESWINDOW: DockingPaneManager.ShowPane ID_PANE_PROPERTIES
        Case ID_WINDOWS_OUTPUT: DockingPaneManager.ShowPane ID_PANE_OUTPUT
        
        Case ID_EDIT_CUT:
            Clipboard.SetText ActiveForm.rtfText.SelRTF
            ActiveForm.rtfText.SelText = vbNullString
        Case ID_EDIT_COPY:          Clipboard.SetText ActiveForm.rtfText.SelRTF
        Case ID_EDIT_PASTE:         ActiveForm.rtfText.SelRTF = Clipboard.GetText
        
        Case ID_FILE_NEW:         LoadNewDoc
        Case ID_FILE_NEW_PROJECT:         LoadNewDoc
        Case ID_VIEW_WELCOME_PAGE:         LoadWelcomePage
        
        Case ID_WINDOW_HIDEALL:
            Dim i As Long
            For i = 0 To DockingPaneManager.PanesCount
                If Not DockingPaneManager(i).Closed And Not DockingPaneManager(i).Hidden Then DockingPaneManager(i).Hide
            Next i
        Case ID_FILE_CLOSE:
            If Not Me.ActiveForm Is Nothing Then
                Unload Me.ActiveForm
            End If
            
        Case ID_THEME_VC6:
            SetCommandbarsTheme xtpThemeVisualStudio6, "", ""
            RefreshThemes
            
        Case ID_THEME_VC7:
            SetCommandbarsTheme xtpThemeOfficeXP, "", ""
            RefreshThemes
            
        Case ID_THEME_VC8:
            SetCommandbarsTheme xtpThemeWhidbey, "", ""
            RefreshThemes
            
        Case ID_THEME_VC10:
            SetCommandbarsTheme xtpThemeVisualStudio2010, "", ""
            RefreshThemes
            
        Case ID_THEME_VC9:
            SetCommandbarsTheme xtpThemeVisualStudio2008, "", ""
            RefreshThemes
                                               
        Case ID_THEME_VC11_LIGHT:
            SetCommandbarsTheme xtpThemeVisualStudio2012Light, App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Light.ini"
            RefreshThemes
            
        Case ID_THEME_VC11_DARK:
            SetCommandbarsTheme xtpThemeVisualStudio2012Dark, App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Dark.ini"
            RefreshThemes
            
        Case ID_THEME_VC14_LIGHT:
            SetCommandbarsTheme xtpThemeVisualStudio2015, App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Light.ini"
            RefreshThemes
            
        Case ID_THEME_VC14_DARK:
            SetCommandbarsTheme xtpThemeVisualStudio2015, App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Dark.ini"
            RefreshThemes
            
        Case ID_THEME_VC14_BLUE:
            SetCommandbarsTheme xtpThemeVisualStudio2015, App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Blue.ini"
            RefreshThemes
            
        Case ID_STATE_DEFAULT:
            CommandBarsGlobalSettings.SetAmbientProperty xtpApplicationStateColor, -1
            AppState = ID_STATE_DEFAULT
            
        Case ID_STATE_ORANGE:
            CommandBarsGlobalSettings.SetAmbientProperty xtpApplicationStateColor, RGB(202, 81, 0)
            'OR: CommandBars.Options.VSStateColor = RGB(202, 81, 0)
            AppState = ID_STATE_ORANGE
            
        Case ID_STATE_BLUE:
            CommandBarsGlobalSettings.SetAmbientProperty xtpApplicationStateColor, RGB(0, 122, 204)
            'OR: CommandBars.Options.VSStateColor = RGB(0, 122, 204)
            AppState = ID_STATE_BLUE

        Case Else:
            MsgBox Control.Caption & " clicked", vbOKOnly, "Button Clicked"
    End Select
End Sub

Private Sub RefreshThemes()
    LoadIcons
    
    Select Case AppTheme
    
        Case ID_THEME_VC6:
            PropertyGridTheme = xtpGridThemeDefault
            DockingPaneManager.VisualTheme = ThemeVisualStudio6
            EnableStickers False, False, -1
            TaskpanelTheme = xtpTaskPanelThemeToolbox
            ControlTheme = xtpAppearanceStandard
            
        Case ID_THEME_VC7:
            PropertyGridTheme = xtpGridThemeDefault
            DockingPaneManager.VisualTheme = ThemeDefault
            EnableStickers False, False, -1
            TaskpanelTheme = xtpTaskPanelThemeToolbox
            ControlTheme = xtpAppearanceStandard
            
        Case ID_THEME_VC8:
            PropertyGridTheme = xtpGridThemeWhidbey
            DockingPaneManager.VisualTheme = ThemeWhidbey
            EnableStickers False, False, -1
            TaskpanelTheme = xtpTaskPanelThemeToolboxWhidbey
            ControlTheme = xtpAppearanceVisualStudio2005
            
        Case ID_THEME_VC9:
            PropertyGridTheme = xtpGridThemeWhidbey
            DockingPaneManager.VisualTheme = ThemeVisualStudio2008
            EnableStickers True, True, StickerStyleVisualStudio2008
            TaskpanelTheme = xtpTaskPanelThemeToolboxWhidbey
            ControlTheme = xtpAppearanceVisualStudio2008
            
        Case ID_THEME_VC10:
            PropertyGridTheme = xtpGridThemeVisualStudio2010
            DockingPaneManager.VisualTheme = ThemeVisualStudio2010
            EnableStickers True, True, StickerStyleVisualStudio2010
            TaskpanelTheme = xtpTaskPanelThemeVisualStudio2010
            ControlTheme = xtpAppearanceVisualStudio2010
            
        Case ID_THEME_VC11_LIGHT:
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Light.ini"
            PropertyGridGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Light.ini"
            TaskPanelGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Light.ini"
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Light.ini"
            PropertyGridTheme = xtpGridThemeVisualStudio2012Light
            DockingPaneManager.VisualTheme = ThemeVisualStudio2012
            EnableStickers True, True, StickerStyleVisualStudio2012Light
            TaskpanelTheme = xtpTaskPanelThemeVisualStudio2012Light
            ControlTheme = xtpAppearanceVisualStudio2012Light
            
        Case ID_THEME_VC11_DARK:
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Dark.ini"
            PropertyGridGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Dark.ini"
            TaskPanelGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Dark.ini"
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2012.dll", "VisualStudio2012Dark.ini"
            PropertyGridTheme = xtpGridThemeVisualStudio2012Dark
            DockingPaneManager.VisualTheme = ThemeVisualStudio2012
            EnableStickers True, True, StickerStyleVisualStudio2012Dark
            TaskpanelTheme = xtpTaskPanelThemeVisualStudio2012Dark
            ControlTheme = xtpAppearanceVisualStudio2012Dark
            
        Case ID_THEME_VC14_LIGHT:
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Light.ini"
            PropertyGridGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Light.ini"
            TaskPanelGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Light.ini"
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Light.ini"
            PropertyGridTheme = xtpGridThemeVisualStudio2015
            DockingPaneManager.VisualTheme = ThemeVisualStudio2015
            EnableStickers True, True, StickerStyleVisualStudio2015Light
            TaskpanelTheme = xtpTaskPanelThemeVisualStudio2015
            ControlTheme = xtpAppearanceVisualStudio2015
            
        Case ID_THEME_VC14_DARK:
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Dark.ini"
            PropertyGridGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Dark.ini"
            TaskPanelGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Dark.ini"
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Dark.ini"
            PropertyGridTheme = xtpGridThemeVisualStudio2015
            DockingPaneManager.VisualTheme = ThemeVisualStudio2015
            EnableStickers True, True, StickerStyleVisualStudio2015Dark
            TaskpanelTheme = xtpTaskPanelThemeVisualStudio2015
            ControlTheme = xtpAppearanceVisualStudio2015
            
        Case ID_THEME_VC14_BLUE:
            DockingPaneGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Blue.ini"
            PropertyGridGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Blue.ini"
            TaskPanelGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Blue.ini"
            SuiteControlsGlobalSettings.ResourceImages.LoadFromFile App.Path & "\..\..\..\Styles\VisualStudio2015.dll", "VisualStudio2015Blue.ini"
            PropertyGridTheme = xtpGridThemeVisualStudio2015
            DockingPaneManager.VisualTheme = ThemeVisualStudio2015
            EnableStickers True, True, StickerStyleVisualStudio2015Light
            TaskpanelTheme = xtpTaskPanelThemeVisualStudio2015
            ControlTheme = xtpAppearanceVisualStudio2015
    End Select
    
    If Not frmClassView Is Nothing Then frmClassView.CommandBars.VisualTheme = CommandBars.VisualTheme
    If Not frmSolutionView Is Nothing Then frmSolutionView.CommandBars.VisualTheme = CommandBars.VisualTheme
    
    If Not frmPropertyGrid Is Nothing Then frmPropertyGrid.wndPropertyGrid.VisualTheme = PropertyGridTheme
    If Not frmClassView Is Nothing Then frmClassView.SetTheme ControlTheme
    If Not frmSolutionView Is Nothing Then frmSolutionView.SetTheme ControlTheme
    If Not frmResourceView Is Nothing Then frmResourceView.SetTheme ControlTheme
    If Not frmOutput Is Nothing Then frmOutput.SetTheme ControlTheme
    If Not frmToolBox Is Nothing Then frmToolBox.wndToolBox.VisualTheme = TaskpanelTheme
    
    If (CommandBars.VisualTheme = xtpThemeVisualStudio2010) Then
        DockingPaneManager.Options.ClientMargin = 6
        DockingPaneManager.Options.ThemedFloatingFrames = True
    Else
        DockingPaneManager.Options.ClientMargin = 0
        DockingPaneManager.Options.ThemedFloatingFrames = False
    End If
    
    fMainForm.ActiveForm.RefreshColors
        
End Sub

Public Sub SetCommandbarsTheme(paintTheme As XtremeCommandBars.XTPVisualTheme, hModule As String, lpszINI As String)
    CommandBars.EnableFrameTheme False
    
    Select Case paintTheme
        Case xtpThemeVisualStudio2012Light, xtpThemeVisualStudio2012Dark, xtpThemeVisualStudio2015
            CommandBars.AllowFrameTransparency False
            CommandBars.SetAllCaps True
            CommandBars.StatusBar.SetAllCaps True
            CommandBars.Options.ShowPopupButtonsSelectedToolbarOnly = True
        Case Else
            CommandBars.AllowFrameTransparency True
            CommandBars.SetAllCaps False
            CommandBars.StatusBar.SetAllCaps False
            CommandBars.Options.ShowPopupButtonsSelectedToolbarOnly = False
    End Select
    
    If Not hModule = "" Then
        CommandBarsGlobalSettings.ResourceImages.LoadFromFile hModule, lpszINI
    End If
    
    CommandBars.VisualTheme = paintTheme
    
    Select Case paintTheme
        Case xtpThemeVisualStudio2012Light, xtpThemeVisualStudio2012Dark, xtpThemeVisualStudio2015
            CommandBars.EnableFrameTheme True
        Case Else
            CommandBars.EnableFrameTheme False
    End Select
            
    Dim ToolTipContext As ToolTipContext
    Set ToolTipContext = CommandBars.ToolTipContext
    ToolTipContext.Style = IIf(paintTheme = xtpThemeVisualStudio2015 Or paintTheme = xtpThemeVisualStudio2012Light Or paintTheme = xtpThemeVisualStudio2012Dark, xtpToolTipResource, xtpToolTipStandard)
    ToolTipContext.ShowTitleAndDescription True, xtpToolTipIconNone
    ToolTipContext.ShowImage True, IMAGEBASE
    ToolTipContext.SetMargin 2, 2, 2, 2
    ToolTipContext.MaxTipWidth = 180
    
    CommandBars.RecalcLayout
       
    Select Case paintTheme
        Case xtpThemeVisualStudio2012Light, xtpThemeVisualStudio2012Dark, xtpThemeVisualStudio2015
            CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\res\CjSample.ico", 20, 20
    End Select

End Sub

Public Sub LoadIcons()

    MAINFRAME = App.Path & "\Icons\Toolbar.png"
    TOOLBAR_WEB = App.Path & "\Icons\Web.png"
    TOOLBAR_EXT = App.Path & "\Icons\Menubar.png"
    PANE_PROPERTIES = App.Path & "\Icons\PaneProperties.bmp"
    CLASSTREE = App.Path & "\Icons\img_classtree.bmp"
    FILETREE = App.Path & "\Icons\img_filetree.bmp"
    RESTREE = App.Path & "\Icons\img_restree.bmp"
    PANE_SOLUTIONEXPLORER = App.Path & "\Icons\SolutionExplorerPane.bmp"
    TOOLBOXICONS = App.Path & "\Icons\Toolbox.bmp"
    TOOLBAR_THEME = App.Path & "\Icons\ThemeBar.png"
    TOOLBAR_STATE = App.Path & "\Icons\StateBar.png"
    TOOLBAR_FULLSCREEN = App.Path & "\Icons\FullScreen.bmp"

    Select Case AppTheme
    
        Case ID_THEME_VC11_LIGHT, ID_THEME_VC14_LIGHT:
        
            MAINFRAME = App.Path & "\Icons\ToolbarLight.png"
            TOOLBAR_WEB = App.Path & "\Icons\WebLight.png"
            TOOLBAR_EXT = App.Path & "\Icons\MenubarLight.png"
            PANE_PROPERTIES = App.Path & "\Icons\PanePropertiesLight.bmp"
            CLASSTREE = App.Path & "\Icons\img_classtreeLight.bmp"
            FILETREE = App.Path & "\Icons\img_filetreeLight.bmp"
            RESTREE = App.Path & "\Icons\img_restreeLight.bmp"
            PANE_SOLUTIONEXPLORER = App.Path & "\Icons\SolutionExplorerPaneLight.bmp"
            TOOLBOXICONS = App.Path & "\Icons\ToolboxLight.bmp"
            TOOLBAR_THEME = App.Path & "\Icons\ThemeBarLight.png"
            TOOLBAR_STATE = App.Path & "\Icons\StateBarLight.png"
            TOOLBAR_FULLSCREEN = App.Path & "\Icons\FullScreenLight.bmp"
            
        Case ID_THEME_VC11_DARK, ID_THEME_VC14_DARK:
        
            MAINFRAME = App.Path & "\Icons\ToolbarDark.png"
            TOOLBAR_WEB = App.Path & "\Icons\WebDark.png"
            TOOLBAR_EXT = App.Path & "\Icons\MenubarDark.png"
            PANE_PROPERTIES = App.Path & "\Icons\PanePropertiesDark.bmp"
            CLASSTREE = App.Path & "\Icons\img_classtreeDark.bmp"
            FILETREE = App.Path & "\Icons\img_filetreeDark.bmp"
            RESTREE = App.Path & "\Icons\img_restreeDark.bmp"
            PANE_SOLUTIONEXPLORER = App.Path & "\Icons\SolutionExplorerPaneDark.bmp"
            TOOLBOXICONS = App.Path & "\Icons\ToolboxDark.bmp"
            TOOLBAR_THEME = App.Path & "\Icons\ThemeBarDark.png"
            TOOLBAR_STATE = App.Path & "\Icons\StateBarDark.png"
            TOOLBAR_FULLSCREEN = App.Path & "\Icons\FullScreenDark.bmp"
            
    End Select

    CommandBarsGlobalSettings.Icons.RemoveAll
            
    Dim CommandBarIconArray(0 To 19) As Long
    
    CommandBarIconArray(0) = ID_FILE_NEW_PROJECT
    CommandBarIconArray(1) = ID_PROJECT_ADDNEWITEM
    CommandBarIconArray(2) = ID_FILE_OPEN
    CommandBarIconArray(3) = ID_FILE_SAVE
    CommandBarIconArray(4) = ID_FILE_SAVE_ALL
    CommandBarIconArray(5) = ID_EDIT_CUT
    CommandBarIconArray(6) = ID_EDIT_COPY
    CommandBarIconArray(7) = ID_EDIT_PASTE
    
    CommandBarIconArray(8) = ID_EDIT_UNDO
    CommandBarIconArray(9) = ID_EDIT_REDO
    
    CommandBarIconArray(10) = ID_DEBUG_START
    CommandBarIconArray(11) = ID_EDIT_STATE
    CommandBarIconArray(12) = ID_EDIT_PLATFORM
    
    CommandBarIconArray(13) = ID_FINDANDREPLACE_FINDINFILES
    CommandBarIconArray(14) = ID_EDIT_FIND_EX
    
    CommandBarIconArray(15) = ID_VIEW_SOLUTIONEXPLORER
    CommandBarIconArray(16) = ID_VIEW_PROPERTIESWINDOW
    CommandBarIconArray(17) = ID_VIEW_RESOURCEVIEW
    CommandBarIconArray(18) = ID_VIEW_TOOLBOX
    CommandBarIconArray(19) = ID_VIEW_CLASSVIEW
       
    CommandBarsGlobalSettings.Icons.LoadBitmap MAINFRAME, CommandBarIconArray, xtpImageNormal
        
    CommandBarsGlobalSettings.Icons.LoadBitmap TOOLBAR_THEME, Array(ID_THEME_VC6, ID_THEME_VC7, ID_THEME_VC8, ID_THEME_VC9, ID_THEME_VC10, ID_THEME_VC11_LIGHT, ID_THEME_VC11_DARK, ID_THEME_VC14_LIGHT, ID_THEME_VC14_DARK, ID_THEME_VC14_BLUE, -1), xtpImageNormal
    CommandBarsGlobalSettings.Icons.LoadBitmap TOOLBAR_STATE, Array(ID_STATE_DEFAULT, ID_STATE_ORANGE, ID_STATE_BLUE, -1), xtpImageNormal
    
    Dim CommandBarIconArray2(0 To 13) As Long
    
    CommandBarIconArray2(0) = ID_FILE_NEW
    CommandBarIconArray2(1) = ID_FILE_PRINT
    CommandBarIconArray2(2) = ID_PROJECT_ADDEXISTINGITEM
    CommandBarIconArray2(3) = ID_FILE_NEW_BLANK
    CommandBarIconArray2(4) = ID_FILE_OPEN_SOLUTION
    CommandBarIconArray2(5) = ID_FILE_CLOSE_SOLUTION
    CommandBarIconArray2(6) = ID_PROJECT_ADD_CLASS
    CommandBarIconArray2(7) = ID_BUILD_SOLUTION
    CommandBarIconArray2(8) = ID_BUILD_COMPILE
    CommandBarIconArray2(9) = ID_DEBUG_WITHOUT_DEBUGGING
    CommandBarIconArray2(10) = ID_WINDOW_NEW
    CommandBarIconArray2(11) = ID_HELP_DYNAMIC
    CommandBarIconArray2(12) = ID_WINDOWS_OUTPUT
    
    CommandBarIconArray2(13) = ID_VIEW_CLASSVIEW
           
    CommandBarsGlobalSettings.Icons.LoadBitmap TOOLBAR_EXT, CommandBarIconArray2, xtpImageNormal
    
    Dim CommandBarWebIconArray(0 To 12) As Long
    
    CommandBarWebIconArray(0) = ID_BROWSER_BACK
    CommandBarWebIconArray(1) = ID_BROWSER_FORWARD
    CommandBarWebIconArray(2) = ID_BROWSER_STOP
    CommandBarWebIconArray(3) = ID_BROWSER_REFRESH
    CommandBarWebIconArray(4) = ID_BROWSER_HOME
    CommandBarWebIconArray(5) = ID_BROWSER_SEARCH
    CommandBarWebIconArray(6) = ID_BROWSER_FAVORITES
    CommandBarWebIconArray(7) = ID_BROWSER_ADDTOFAVORITE
    CommandBarWebIconArray(8) = ID_BROWSER_SYNC
    CommandBarWebIconArray(9) = ID_BROWSER_PREV
    CommandBarWebIconArray(10) = ID_BROWSER_NEXT
    CommandBarWebIconArray(11) = ID_BROWSER_FONT
    CommandBarWebIconArray(12) = ID_BROWSER_URL
               
    CommandBarsGlobalSettings.Icons.LoadBitmap TOOLBAR_WEB, CommandBarWebIconArray, xtpImageNormal

    CommandBarsGlobalSettings.Icons.LoadBitmap TOOLBAR_FULLSCREEN, Array(ID_VIEW_FULLSCREEN), xtpImageNormal
       
    SuiteControlsGlobalSettings.Icons.LoadBitmap CLASSTREE, Array(ID_CLASS_TREE_1, ID_CLASS_TREE_2, ID_CLASS_TREE_3, ID_CLASS_TREE_4, ID_CLASS_TREE_5, ID_CLASS_TREE_6, ID_CLASS_TREE_7, ID_CLASS_TREE_8), xtpImageNormal
    SuiteControlsGlobalSettings.Icons.LoadBitmap FILETREE, Array(ID_SOLUTION_TREE_1, ID_SOLUTION_TREE_2, ID_SOLUTION_TREE_3, ID_SOLUTION_TREE_4, ID_SOLUTION_TREE_5, ID_SOLUTION_TREE_6), xtpImageNormal
    SuiteControlsGlobalSettings.Icons.LoadBitmap RESTREE, Array(ID_RESOURCE_TREE_1, ID_RESOURCE_TREE_2, ID_RESOURCE_TREE_3, ID_RESOURCE_TREE_4, ID_RESOURCE_TREE_5, ID_RESOURCE_TREE_6, ID_RESOURCE_TREE_7, ID_RESOURCE_TREE_8, ID_RESOURCE_TREE_9, ID_RESOURCE_TREE_10, ID_RESOURCE_TREE_11), xtpImageNormal
    
    Dim ArrayIcons(42) As Long, i As Long
    For i = 0 To 42
        ArrayIcons(i) = i
    Next i
    TaskPanelGlobalSettings.Icons.LoadBitmap TOOLBOXICONS, ArrayIcons, xtpImageNormal
    
    Select Case AppTheme
    
        Case ID_THEME_VC11_LIGHT, ID_THEME_VC14_LIGHT:
            CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\Icons\black.ico", 20, 20
        Case ID_THEME_VC11_DARK, ID_THEME_VC14_DARK:
            CommandBars.PaintManager.LoadFrameIcon App.hInstance, App.Path + "\Icons\white.ico", 20, 20
            
    End Select
End Sub

Private Sub EnableStickers(bAlphaContext As Boolean, bShow As Boolean, nStyle As XtremeDockingPane.DockingContextStickerStyle)
    
    DockingPaneManager.Options.AlphaDockingContext = bAlphaContext
    DockingPaneManager.Options.ShowDockingContextStickers = bShow
  
    If Not nStyle = -1 Then
        DockingPaneManager.Options.StickerStyle = nStyle
    End If
    
End Sub

Private Sub LoadNewDoc()
    Dim frmD As frmDocument
    Set frmD = New frmDocument
    Set frmD.fMain = Me
    frmD.Show
End Sub

Private Sub LoadWelcomePage()
    Dim frmD As frmMarkupView
    Set frmD = New frmMarkupView
    frmD.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    Dim i As Long
   'close all sub forms
    For i = Forms.Count - 1 To 1 Step -1
        Unload Forms(i)
    Next
    
End Sub

'******************Helper functions******************
Private Sub DockBarRightOf(BarToDock As CommandBar, BarOnLeft As CommandBar)
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    CommandBars.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    
    CommandBars.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position

End Sub

Private Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, Id As Long, Optional BeginGroup As Boolean = False) As CommandBarControl
    Dim Control As CommandBarControl
    Set Control = Controls.Add(ControlType, Id, "")
    
    Control.BeginGroup = BeginGroup
    Set AddButton = Control
End Function
