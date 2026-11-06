Attribute VB_Name = "Definitions"
Public AppTheme As Long
Public strSchema As String
Public strColor As String
Public ControlTheme As AppearanceStyle
Public TaskpanelTheme As XtremeTaskPanel.XTPTaskPanelVisualTheme
Public PropertyGridTheme As XtremePropertyGrid.XTPPropertyGridVisualTheme

Public Const ID_FILE_NEW = 100
Public Const ID_FILE_PRINT = 101
Public Const ID_FILE_NEW_PROJECT = 102
Public Const ID_FILE_NEW_BLANK = 103
Public Const ID_FILE_OPEN_SOLUTION = 104
Public Const ID_FILE_CLOSE_SOLUTION = 105
Public Const ID_FILE_OPEN = 106
Public Const ID_FILE_SAVE = 107
Public Const ID_FILE_SAVE_ALL = 1077
Public Const ID_FILE_CLOSE = 108
Public Const ID_FILE_OPEN_PROJECT = 109
Public Const ID_FILE_SAVE_AS = 110
Public Const ID_FILE_PRINT_SETUP = 111
Public Const ID_FILE_EXIT = 112

Public Const ID_PROJECT_ADDNEWITEM = 120
Public Const ID_PROJECT_ADDEXISTINGITEM = 121
Public Const ID_PROJECT_ADD_CLASS = 122

Public Const ID_BUILD_SOLUTION = 130
Public Const ID_BUILD_COMPILE = 131

Public Const ID_DEBUG_START = 140
Public Const ID_DEBUG_WITHOUT_DEBUGGING = 141
Public Const ID_WINDOWS_OUTPUT = 142
Public Const ID_DEBUG_CONFIGURATION = 143
Public Const ID_WINDOW_HIDEALL = 144

Public Const ID_WINDOW_NEW = 150

Public Const ID_HELP_DYNAMIC = 160

Public Const ID_VIEW_FULLSCREEN = 154
Public Const ID_VIEW_CLASSVIEW = 170
Public Const ID_VIEW_RESOURCEVIEW = 171
Public Const ID_VIEW_SOLUTIONEXPLORER = 172
Public Const ID_VIEW_TOOLBOX = 173
Public Const ID_VIEW_PROPERTIESWINDOW = 174
Public Const ID_VIEW_WELCOME_PAGE = 234

Public Const ID_EDIT_CUT = 180
Public Const ID_EDIT_COPY = 181
Public Const ID_EDIT_PASTE = 182
Public Const ID_FINDANDREPLACE_FINDINFILES = 183
Public Const ID_EDIT_STATE = 1880
Public Const ID_EDIT_PLATFORM = 1881
Public Const ID_EDIT_FIND_EX = 1883
Public Const ID_EDIT_UNDO = 1884
Public Const ID_EDIT_REDO = 1885

Public Const ID_BROWSER_BACK = 190
Public Const ID_BROWSER_FORWARD = 191
Public Const ID_BROWSER_STOP = 192
Public Const ID_BROWSER_REFRESH = 193
Public Const ID_BROWSER_HOME = 194
Public Const ID_BROWSER_SEARCH = 195
Public Const ID_BROWSER_FAVORITES = 196
Public Const ID_BROWSER_ADDTOFAVORITE = 197
Public Const ID_BROWSER_SYNC = 198
Public Const ID_BROWSER_PREV = 199
Public Const ID_BROWSER_NEXT = 200
Public Const ID_BROWSER_FONT = 201
Public Const ID_BROWSER_URL = 202

Public Const ID_PANE_TOOLBOX = ID_VIEW_TOOLBOX
Public Const ID_PANE_CLASSVIEW = ID_VIEW_CLASSVIEW
Public Const ID_PANE_SOLUTIONVIEW = ID_VIEW_SOLUTIONEXPLORER
Public Const ID_PANE_PROPERTIES = ID_VIEW_PROPERTIESWINDOW
Public Const ID_PANE_OUTPUT = ID_WINDOWS_OUTPUT
Public Const ID_PANE_RESOURCEVIEW = ID_VIEW_RESOURCEVIEW

Public Const ID_THEME_OFFICE = 1300
Public Const ID_THEME_DEFAULT = 1310
Public Const ID_THEME_OFFICE2003 = 1320
Public Const ID_THEME_NATIVEXP = 1330
Public Const ID_THEME_SWITCH = 5340
Public Const ID_THEME_OFFICE2007 = 1890
Public Const ID_THEME_RIBBON = 3010

Public Const ID_THEME_VC6 = 32786
Public Const ID_THEME_VC7 = 32787
Public Const ID_THEME_VC8 = 32788
Public Const ID_THEME_VC9 = 32789
Public Const ID_THEME_VC10 = 32790
Public Const ID_THEME_VC11_LIGHT = 32791
Public Const ID_THEME_VC11_DARK = 32792
Public Const ID_THEME_VC14_LIGHT = 32793
Public Const ID_THEME_VC14_DARK = 32794
Public Const ID_THEME_VC14_BLUE = 32795

Public Const ID_STATE_DEFAULT = 32796
Public Const ID_STATE_ORANGE = 32797
Public Const ID_STATE_BLUE = 32798

Public Const ID_HELP_ABOUT = 800

Public Const ITEM_DIRECTORY = 1

Public Const FSHIFT = 4
Public Const FCONTROL = 8
Public Const FALT = 16

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B

Public Const ID_TOOLBOXFOLDER_PROPERTYPANES = "PropertyPanes"
Public Const ID_TOOLBOXFOLDER_DATA = "Data"
Public Const ID_TOOLBOXFOLDER_COMPONENTS = "Components"
Public Const ID_TOOLBOXFOLDER_WINDOWSFORMS = "Dialog Editor"

Public Const ID_TOOLBOXFOLDER_CLIPBOARDRING = "Clipboard Ring"
Public Const ID_TOOLBOXFOLDER_GENERAL = "General"

Public Const ID_TOOLBOXITEM_POINTER = "Pointer"
Public Const ID_TOOLBOXITEM_BUTTON = "Button"
Public Const ID_TOOLBOXITEM_CHECK_BOX = "Check Box"
Public Const ID_TOOLBOXITEM_EDIT_CONTROL = "Edit Control"
Public Const ID_TOOLBOXITEM_COMBO_BOX = "Combo Box"
Public Const ID_TOOLBOXITEM_LIST_BOX = "List Box"
Public Const ID_TOOLBOXITEM_GROUP_BOX = "Group Box"
Public Const ID_TOOLBOXITEM_RADIO_BUTTON = "Radio Button"
Public Const ID_TOOLBOXITEM_STATIC_TEXT = "Static Text"
Public Const ID_TOOLBOXITEM_PICTURE_CONTROL = "Picture Control"
Public Const ID_TOOLBOXITEM_HORIZONTAL_SCROLL_BAR = "Horizontal Scroll Bar"
Public Const ID_TOOLBOXITEM_VERTICAL_SCROLL_BAR = "Vertical Scroll Bar"
Public Const ID_TOOLBOXITEM_SLIDER_CONTROL = "Slider Control"
Public Const ID_TOOLBOXITEM_SPIN_CONTROL = "Spin Control"
Public Const ID_TOOLBOXITEM_PROGRESS_CONTROL = "Progress Control"
Public Const ID_TOOLBOXITEM_HOT_KEY = "Hot Key"

Public Const ID_TOOLBOXITEM_LIST_CONTROL = "List Control"
Public Const ID_TOOLBOXITEM_TREE_CONTROL = "Tree Control"
Public Const ID_TOOLBOXITEM_TAB_CONTROL = "Tab Control"
Public Const ID_TOOLBOXITEM_ANIMATION_CONTROL = "Animation Control"
Public Const ID_TOOLBOXITEM_RICH_EDIT_CONTROL = "Rich Edit 2.0 Control"
Public Const ID_TOOLBOXITEM_DATE_TIME_PICKER = "Date Time Picker"
Public Const ID_TOOLBOXITEM_MONTH_CALENDAR_CONTROL = "Month Calendar Control"
Public Const ID_TOOLBOXITEM_IP_ADDRESS_CONTROL = "IP Address Control"
Public Const ID_TOOLBOXITEM_EXTENDED_COMBO_BOX = "Extended Combo Box"
Public Const ID_TOOLBOXITEM_CUSTOM_CONTROL = "Custom Control"
Public Const ID_TOOLBOXITEM_DATASET = "DataSet"
Public Const ID_TOOLBOXITEM_OLEDBDATAADAPTER = "OleDbDataAdapter"
Public Const ID_TOOLBOXITEM_OLEDBCONNECTION = "OleDbConnection"
Public Const ID_TOOLBOXITEM_OLEDBCOMMAND = "OleDbCommand"
Public Const ID_TOOLBOXITEM_SQLDATAADAPTER = "SqlDataAdapter"
Public Const ID_TOOLBOXITEM_SQLCONNECTION = "SqlConnection"

Public Const ID_TOOLBOXITEM_SQLCOMMAND = "SqlCommand"
Public Const ID_TOOLBOXITEM_DATAVIEW = "DataView"
Public Const ID_TOOLBOXITEM_FILESYSTEMWATCHER = "FileSystemWatcher"
Public Const ID_TOOLBOXITEM_EVENTLOG = "EventLog"
Public Const ID_TOOLBOXITEM_DIRECTORYENTRY = "DirectoryEntry"
Public Const ID_TOOLBOXITEM_DIRECTORYSEARCHER = "DirectorySearcher"
Public Const ID_TOOLBOXITEM_MESSAGEQUEUE = "MessageQueue"
Public Const ID_TOOLBOXITEM_PERFORMANCECOUNTER = "PerformanceCounter"
Public Const ID_TOOLBOXITEM_PROCESS = "Process"
Public Const ID_TOOLBOXITEM_SERVICECONTROLLER = "ServiceController"
Public Const ID_TOOLBOXITEM_TIMER = "Timer"

Public Const ID_CLASS_TREE_0 = 500
Public Const ID_CLASS_TREE_1 = 501
Public Const ID_CLASS_TREE_2 = 502
Public Const ID_CLASS_TREE_3 = 503
Public Const ID_CLASS_TREE_4 = 504
Public Const ID_CLASS_TREE_5 = 505
Public Const ID_CLASS_TREE_6 = 506
Public Const ID_CLASS_TREE_7 = 507
Public Const ID_CLASS_TREE_8 = 508
Public Const ID_CLASS_TREE_9 = 509

Public Const ID_RESOURCE_TREE_0 = 510
Public Const ID_RESOURCE_TREE_1 = 511
Public Const ID_RESOURCE_TREE_2 = 512
Public Const ID_RESOURCE_TREE_3 = 513
Public Const ID_RESOURCE_TREE_4 = 514
Public Const ID_RESOURCE_TREE_5 = 515
Public Const ID_RESOURCE_TREE_6 = 516
Public Const ID_RESOURCE_TREE_7 = 517
Public Const ID_RESOURCE_TREE_8 = 518
Public Const ID_RESOURCE_TREE_9 = 519
Public Const ID_RESOURCE_TREE_10 = 531
Public Const ID_RESOURCE_TREE_11 = 532

Public Const ID_SOLUTION_TREE_0 = 520
Public Const ID_SOLUTION_TREE_1 = 521
Public Const ID_SOLUTION_TREE_2 = 522
Public Const ID_SOLUTION_TREE_3 = 523
Public Const ID_SOLUTION_TREE_4 = 524
Public Const ID_SOLUTION_TREE_5 = 525
Public Const ID_SOLUTION_TREE_6 = 526
Public Const ID_SOLUTION_TREE_7 = 527
Public Const ID_SOLUTION_TREE_8 = 528
Public Const ID_SOLUTION_TREE_9 = 529
Public Const ID_SOLUTION_TREE_10 = 530

Public Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
        (lpversioninformation As OSVERSIONINFO) As Long

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type

Public Function IsWinVistaOrGreater() As Boolean
    
    Dim osVersion As OSVERSIONINFO
    osVersion.dwOSVersionInfoSize = Len(osVersion)
    GetVersionEx osVersion
    IsWinVistaOrGreater = IIf(osVersion.dwMajorVersion >= 6, True, False)
    
End Function

Public Sub SetTreeTheme(wndTreeView As XtremeSuiteControls.TreeView, nTheme As AppearanceStyle)
    wndTreeView.Appearance = nTheme
    
    Dim bVS2010Plus     As Boolean
    bVS2010Plus = ((nTheme = xtpAppearanceVisualStudio2012Dark) Or _
                  (nTheme = xtpAppearanceVisualStudio2012Light) Or _
                  (nTheme = xtpAppearanceVisualStudio2015) Or _
                  (nTheme = xtpAppearanceVisualStudio2010))

    If (bVS2010Plus) Then
        wndTreeView.ShowLines = xtpTreeViewHideLines
    Else
        wndTreeView.ShowLines = xtpTreeViewShowLines
    End If
    
    If IsWinVistaOrGreater Then
        If ((nTheme = xtpAppearanceVisualStudio2015) Or (nTheme = xtpAppearanceVisualStudio2012Light) Or (nTheme = xtpAppearanceVisualStudio2010)) Then
            wndTreeView.EnableExplorerStyle True, False
        Else
            wndTreeView.EnableExplorerStyle False, False
        End If
    End If
    
    Select Case nTheme
    
        Case xtpAppearanceVisualStudio2012Light, xtpAppearanceVisualStudio2012Dark, xtpAppearanceVisualStudio2015
            wndTreeView.ShowBorder = False
            
        Case Else
            wndTreeView.ShowBorder = True
            
    End Select
    
    wndTreeView.Appearance = nTheme
    
End Sub

