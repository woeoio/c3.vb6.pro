VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTreeExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ---------------------------------------------------
' Autor:            Leandro I. Ascierto
' Date:             20/07/2010
' Web:              www.leandroascierto.com
' Requirements:     Windows XP or Later
' History:          20/07/2010 First
'                   07/08/2010 Remove SHSimpleIDListFromPath and CoTaskMemFree < don't work with some directory in CD unit
'                              replaced by an internal Array (ArrIDLIST)
' ---------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long   ' <---
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Const MAX_PATH  As Long = 260

Private Type POINTAPI
    x                   As Long
    Y                   As Long
End Type

Private Type Rect
    Left                As Long
    Top                 As Long
    Right               As Long
    Bottom              As Long
End Type

Private Type TOOLINFO
    cbSize              As Long
    uFlags              As Long
    hwnd                As Long
    uId                 As Long
    Rect                As Rect
    hInst               As Long
    lpszText            As String
    lParam              As Long
End Type

Private Type MENUITEMINFO
    cbSize              As Long
    fMask               As Long
    fType               As Long
    fState              As Long
    wID                 As Long
    hSubMenu            As Long
    hbmpChecked         As Long
    hbmpUnchecked       As Long
    dwItemData          As Long
    dwTypeData          As Long
    cch                 As Long
    hbmpItem            As Long
End Type

Private Type SHFILEINFO
    hIcon               As Long
    iIcon               As Long
    dwAttributes        As Long
    szDisplayName       As String * MAX_PATH
    szTypeName          As String * 80
End Type

Private Type SHITEMID
    cb                  As Long
    abID                As Byte
End Type

Private Type ITEMIDLIST
    mkid                As SHITEMID
End Type

Private Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes    As Long
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * MAX_PATH
    cAlternate          As String * 14
End Type

Private Type ItemData
    FileName As String
    Attribute As Long
    hIcon As Long
    IDList As Long
    HasContent As Boolean
End Type
    
Private Const INVALID_HANDLE_VALUE      As Long = -1

Private Const WM_USER               As Long = &H400

Private Const WS_EX_TOPMOST         As Long = &H8&

Private Const TTM_ADDTOOL           As Long = (WM_USER + 4)
Private Const TTM_UPDATETIPTEXT     As Long = (WM_USER + 12)
Private Const TTM_SETMAXTIPWIDTH    As Long = (WM_USER + 24)
Private Const TTM_DELTOOL           As Long = (WM_USER + 51)
Private Const TTM_TRACKACTIVATE     As Long = (WM_USER + 17)

Private Const TTF_IDISHWND          As Long = &H1
Private Const TTS_ALWAYSTIP         As Long = &H1
Private Const TOOLTIPS_CLASS        As String = "tooltips_class32"

Private Const MIIM_DATA             As Long = &H20


Private Const SHGFI_SMALLICON       As Long = &H1
Private Const SHGFI_ICON            As Long = &H100
Private Const SHGFI_TYPENAME        As Long = &H400
Private Const SHGFI_ATTRIBUTES      As Long = &H800
Private Const SHGFI_DISPLAYNAME     As Long = &H200

Private Const SFGAO_HIDDEN          As Long = &H80000

Private Const HWND_TOPMOST          As Long = -1
Private Const SWP_NOSIZE            As Long = &H1
Private Const SWP_NOMOVE            As Long = &H2
Private Const SWP_NOACTIVATE        As Long = &H10

Public Enum eSpecialFolders
    SF_DESKTOP = &H0
    SF_PROGRAMS = &H2
    SF_MYDOCS = &H5
    SF_FAVORITES = &H6
    SF_STARTUP = &H7
    SF_RECENT = &H8
    SF_SENDTO = &H9
    SF_STARTMENU = &HB
    SF_MYMUSIC = &HD
    SF_MYVIDEOS = &HC
    SF_NETHOOD = &H13
    SF_FONTS = &H14
    SF_SHELLNEW = &H15
    SF_STARTUP2 = &H18
    SF_ALLUSERSDESK = &H19
    SF_APPDATA = &H1A
    SF_PRINTHOOD = &H1B
    SF_APPDATA2 = &H1C
    SF_TEMPINETFILES = &H20
    SF_COOKIES = &H21
    SF_HISTORY = &H22
    SF_ALLUSERSAPPDATA = &H23
    SF_WINDOWS = &H24
    SF_WINSYSTEM = &H25
    SF_PROGFILES = &H26
    SF_MYPICS = &H27
    SF_USERDIR = &H28
    SF_COMMON = &H2B
End Enum

Public Event SelectFile(ByVal FilePath As String)
Public Event SelectFolder(ByVal FolderPath As String)


Private STRING_TYPE                 As String
Private STRING_SIZE                 As String
Private STRING_DATE                 As String
Private STRING_EMPTY                As String

Private MII                         As MENUITEMINFO
Private tSHFI                       As SHFILEINFO
Private TI                          As TOOLINFO
Private ArrFolders()                As ItemData
Private ArrFiles()                  As ItemData
Private cInitialFolders             As Collection
Private m_hwndTT                    As Long
Private j                           As Long
Private m_DisplayHiddenFile         As Boolean
Private mStrFilter                  As String
Private ArrFilter()                 As String
Private m_ShowMouseHourglass        As Boolean
Private ArrIDLIST()                 As String
Private WithEvents cTVP             As ucTreeViewPlus
Attribute cTVP.VB_VarHelpID = -1

Public Function GetNodePath(N As Node)
     GetNodePath = GetPathFromIDList(N.key)
End Function


Public Sub OpenPath(ByVal Path As String)
    Dim i As Long
    Dim PathPart() As String
    Dim j As Long
    Dim sPath As String
    Dim N As Node
    Dim CurPath As String
    PathPart = Split(Path, "\")
    Set N = cTVP.Nodes
    
    
    For j = 0 To UBound(PathPart)
        If j = 0 Then
            sPath = UCase(PathPart(j)) & "\"
        Else
            sPath = sPath & UCase(PathPart(j)) & "\"
        End If
        For i = 1 To N.Childs.Count
            CurPath = UCase(ArrIDLIST(N.Childs(i).key))
            If Right$(CurPath, 1) <> "\" Then CurPath = CurPath & "\"
            
            
            If CurPath = sPath Then
                N.Childs(i).Expanded = True
                cTVP_NodeExpand N.Childs(i)
                
                Set N = N.Childs(i)
                Exit For
            End If
        Next
    Next
    
   
    cTVP.Update
    cTVP.SelNode = N
    'cTVP.EnsureVisible cTVP.Nodes
End Sub




Public Sub Init(TVP As ucTreeViewPlus, Optional Filter As String = "*.*")
    Dim index As Long, i As Long
    Dim sFolder         As String
    Dim FileAttrHidden  As Boolean
    Dim tSHFI           As SHFILEINFO
    Dim IDList As Long
    
    Set cTVP = TVP
    
    
    If m_hwndTT Then
        Call SendMessage(m_hwndTT, TTM_DELTOOL, 0, TI)
        Call DestroyWindow(m_hwndTT)
        m_hwndTT = 0
    End If
    
    m_hwndTT = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, vbNullString, TTS_ALWAYSTIP, 0, 0, 0, 0, TVP.hwnd, 0, App.hInstance, ByVal 0)
    
    MII.cbSize = Len(MII)
    MII.fMask = MIIM_DATA

    If m_hwndTT Then
        With TI
            .cbSize = Len(TI)
            .uFlags = TTF_IDISHWND
            .uId = TVP.hwnd
        End With
        Call SendMessage(m_hwndTT, TTM_SETMAXTIPWIDTH, 0, ByVal 250)
        Call SendMessage(m_hwndTT, TTM_ADDTOOL, 0, TI)
    End If
    

    
    
    
    cTVP.Nodes.Clear
    
    ReDim ArrFiles(1)
    ReDim ArrIDLIST(1)
    
    mStrFilter = Filter
    
    If InStr(mStrFilter, ";") Then
        ArrFilter = Split(mStrFilter, ";")
    Else
        ReDim ArrFilter(0)
        ArrFilter(0) = mStrFilter
    End If
    
    Set cInitialFolders = New Collection
    cTVP.ImageListInitialize
    
    With cInitialFolders
        .Add GetSpecialFolder(SF_DESKTOP)         ' Escritorio
        .Add GetSpecialFolder(SF_PROGFILES)       ' Archivos de programa.
        .Add GetSpecialFolder(SF_MYDOCS)          ' Mis documentos.
        
        For i = 0 To GetDrivesCount
            .Add GetDrive(i)                      ' Todas las unidades disponibles.
        Next

    End With
    
    For index = 1 To cInitialFolders.Count

        sFolder = cInitialFolders(index)
        NormalizePath sFolder
        If PathIsDirectory(sFolder) Then
   
            Call SHGetFileInfo(sFolder, ByVal 0&, tSHFI, Len(tSHFI), SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_ATTRIBUTES Or SHGFI_DISPLAYNAME)
            'AppendMenu hMenu, MF_POPUP, hNewSubMenu, StripNulls(tSHFI.szDisplayName)

            FileAttrHidden = (tSHFI.dwAttributes And SFGAO_HIDDEN) = SFGAO_HIDDEN
            'cImgMenu.AddIconFromHandle tSHFI.hIcon, FileAttrHidden
            IDList = GetIDListFromPath(sFolder)
            cTVP.ImageListAddImage tSHFI.hIcon, CStr(IDList)
            
            cTVP.Nodes.Add StripNulls(tSHFI.szDisplayName), CStr(IDList), CStr(IDList), , True

            DestroyIcon tSHFI.hIcon
            
            'cImgMenu.PutImageToApiMenu cImgMenu.ImageCount, hMenu, i - 1, IDList
            i = i + 1
        End If
     
    Next
End Sub

Private Function GetPathFromIDList(ByVal IDL As Long) As String
    On Error Resume Next
    GetPathFromIDList = ArrIDLIST(IDL)
End Function

Private Function GetIDListFromPath(ByVal sPath As String) As Long
    ArrIDLIST(UBound(ArrIDLIST)) = sPath
    GetIDListFromPath = UBound(ArrIDLIST)
    ReDim Preserve ArrIDLIST(UBound(ArrIDLIST) + 1)
End Function

Private Function StripNulls(sData As String) As String
    StripNulls = Left$(sData, lstrlen(sData))
End Function

Private Sub NormalizePath(sData As String)
    sData = IIf(Right$(sData, 1) = "\", sData, sData & "\")
End Sub

Private Function FileExist(ByVal sFile As String) As Boolean

    Dim tWFD As WIN32_FIND_DATA
    Dim hFindFile As Long
    
    hFindFile = FindFirstFile(sFile, tWFD)
    
    If hFindFile = INVALID_HANDLE_VALUE Then
        FileExist = False
    Else
        FileExist = True
        Call FindClose(hFindFile)
    End If

End Function

Public Function GetSpecialFolder(SpecialFolder As eSpecialFolders) As String
    Dim lRet         As Long
    Dim IDL          As ITEMIDLIST
    Dim sPath        As String * 512
    
    lRet = SHGetSpecialFolderLocation(100, SpecialFolder, IDL)
    
    If lRet = 0 Then
        lRet = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath$)
        GetSpecialFolder = Left$(sPath, InStr(sPath, Chr$(0)) - 1)
    End If
End Function

Public Property Get GetDrivesCount() As Long
    GetDrivesCount = UBound(GetDrive)
End Property


Public Function GetDrive() As Variant
    Dim lngRet As Long
    Dim strDrives As String * 255
    lngRet = GetLogicalDriveStrings(Len(strDrives), strDrives)
    GetDrive = Split(Left(strDrives, lngRet - 1), Chr$(0))
End Function

Private Function GetFileSize(FileName As String) As Currency
    Dim FstValue As Long
    
    FstValue = FileLen(FileName)
    
    If FstValue < 0 Then
        GetFileSize = (2147483648# * 2) + FstValue
    Else
        GetFileSize = FstValue
    End If
End Function

Private Function GetFormatSizeFile(ByVal Bytes As Currency) As String
    If Bytes >= 1024 Then
        If Bytes >= 1073741824 Then
            GetFormatSizeFile = Format((Bytes / 1073741824), "##,###,##0.00") & " GB"
        Else
            If Bytes >= 1048576 Then
                GetFormatSizeFile = Format((Bytes / 1048576), "##,###,##0.00") & " MB"
            Else
                GetFormatSizeFile = Format((Bytes \ 1024), "##,###,##0") & " KB"
            End If
        End If
    Else
        GetFormatSizeFile = Format(Bytes, "##,###,##0") & " Bytes"
    End If
End Function


Private Function GetLnkTarget(ByVal sPath As String) As String 'By Karcrack
    Dim WSH         As Object
    Dim oShellLnk   As Object
 
    Set WSH = CreateObject("WScript.Shell")
    Set oShellLnk = WSH.CreateShortcut(sPath)
 
    GetLnkTarget = oShellLnk.TargetPath
 
    Set WSH = Nothing
    Set oShellLnk = Nothing
End Function

Private Function GetFileTypeName(ByVal sPath As String) As String
    Dim tSHFI As SHFILEINFO
    Call SHGetFileInfo(sPath, ByVal 0&, tSHFI, Len(tSHFI), SHGFI_TYPENAME)
    GetFileTypeName = StripNulls(tSHFI.szTypeName)
End Function

Private Function GetPathName(ByVal sPath As String) As String
    GetPathName = Right$(sPath, Len(sPath) - InStrRev(sPath, "\"))
End Function

Private Sub Class_Initialize()
    Call InitCommonControls
    STRING_TYPE = "Tipo"
    STRING_SIZE = "Tamaño"
    STRING_DATE = "Fecha de Creación"
    STRING_EMPTY = "(Vacio)"
End Sub

Private Sub Class_Terminate()
    If m_hwndTT Then
        Call SendMessage(m_hwndTT, TTM_DELTOOL, 0, TI)
        Call DestroyWindow(m_hwndTT)
        m_hwndTT = 0
    End If
End Sub

Private Sub cTVP_NodeClick(Node As Node)
    Dim sPath As String
    sPath = GetPathFromIDList(Node.key)
    If PathIsDirectory(sPath) Then
        RaiseEvent SelectFolder(sPath)
    Else
        RaiseEvent SelectFile(sPath)
    End If
    
End Sub

Private Sub cTVP_NodeMouseEnter(Node As Node)
    Dim sPath As String
    sPath = GetPathFromIDList(Node.key)
    If PathIsDirectory(sPath) = 0 Then
        ShowToolTips sPath
    End If
End Sub

Private Sub cTVP_NodeMouseLeave(Node As Node)
    If m_hwndTT Then SendMessage m_hwndTT, TTM_TRACKACTIVATE, False, TI
End Sub

Private Sub ShowToolTips(ByVal sPath As String)
    On Local Error Resume Next
        
    TI.lpszText = STRING_TYPE & ": " & GetFileTypeName(sPath) & vbCrLf & _
                  STRING_DATE & ": " & FileDateTime(sPath) & vbCrLf & _
                  STRING_SIZE & ": " & GetFormatSizeFile(GetFileSize(sPath))
                  
    SetWindowPos m_hwndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
                  
    Call SendMessage(m_hwndTT, TTM_UPDATETIPTEXT, 0, TI)
    
    SendMessage m_hwndTT, TTM_TRACKACTIVATE, True, TI

End Sub


Private Sub cTVP_SelNodeChange(Node As Node)
    cTVP_NodeClick Node
End Sub

Private Sub cTVP_NodeDblClick(Node As Node)
    'node.Expanded = Not node.Expanded
    cTVP_NodeExpand Node
    cTVP.Update
End Sub

Private Sub cTVP_NodeExpand(Node As Node)
    Dim sPath As String
    Dim index As Long
    Dim FileAttrHidden As Boolean
    Dim i As Long
    
    sPath = GetPathFromIDList(Node.key)

    Call FindFiles(sPath, m_DisplayHiddenFile)
    

    
    'If UBound(ArrFiles) = 1 And UBound(ArrFolders) = 1 Then
    '    AppendMenu hSubMenu, MF_STRING Or MF_GRAYED, -1, Space$(10) & STRING_EMPTY & Space$(5)
    'Else

        Node.Clear
        For index = 0 To UBound(ArrFolders) - 1
            With ArrFolders(index)
                FileAttrHidden = (.Attribute And vbHidden) = vbHidden
                cTVP.ImageListAddImage .hIcon, CStr(.IDList)
                Node.Add .FileName, CStr(.IDList), CStr(.IDList), , .HasContent
                DestroyIcon .hIcon
                i = i + 1
            End With
        Next
        
        For index = 0 To UBound(ArrFiles) - 1
            With ArrFiles(index)
                FileAttrHidden = (.Attribute And vbHidden) = vbHidden
                cTVP.ImageListAddImage .hIcon, CStr(.IDList)
                Node.Add .FileName, CStr(.IDList), CStr(.IDList), , .HasContent
                DestroyIcon .hIcon
            End With
            i = i + 1
            j = j + 1
        Next
        
    'End If

    If m_ShowMouseHourglass Then Screen.MousePointer = vbArrow
End Sub

Public Function CarpetaContieneArchivosYSubcarpetas(ByVal rutaCarpeta As String) As Boolean
    Dim hFindFile As Long
    Dim findData As WIN32_FIND_DATA


    ' Agregar una barra diagonal invertida al final de la ruta de la carpeta si no la tiene
    If Right(rutaCarpeta, 1) <> "\" Then
        rutaCarpeta = rutaCarpeta & "\"
    End If

    ' Buscar el primer archivo o subcarpeta dentro de la carpeta
    hFindFile = FindFirstFile(rutaCarpeta & "*", findData)

    If hFindFile <> -1 Then
        ' Iterar sobre los archivos y subcarpetas encontrados
        Do
            ' Ignorar las carpetas "." y ".."
            If Left(findData.cFileName, 1) <> "." Then

               CarpetaContieneArchivosYSubcarpetas = True
               Exit Do
            End If
        Loop While FindNextFile(hFindFile, findData) = 1

        ' Cerrar el identificador de búsqueda
        FindClose hFindFile
    End If

    ' Devolver el resultado
    
End Function



Private Function FindFiles(ByVal sDir As String, ByVal bHiddenFile As Boolean) As Boolean
    On Error Resume Next
    
    Dim lRet                As Long
    Dim lhSearch            As Long
    Dim tWFD                As WIN32_FIND_DATA
    Dim sFile               As String
    Dim sFolder             As String
    Dim index               As Long
    Dim tSHFI               As SHFILEINFO
    Dim IDList              As Long
    Dim sPath               As String
    Dim bIsLinkFolder       As Boolean
    Dim sLinkPath           As String
    Dim IsLink              As Boolean
    
    ReDim ArrFolders(0)
    ReDim ArrFiles(0)
    
    If sDir = vbNullString Then Exit Function

    Call NormalizePath(sDir)

    lhSearch = FindFirstFile(sDir & "*", tWFD)
    
    If Not lhSearch = INVALID_HANDLE_VALUE Then
        Do
            If Not ((tWFD.dwFileAttributes And vbHidden) = vbHidden) Or bHiddenFile = True Then
                
                sPath = StripNulls(tWFD.cFileName)
                
                If LCase(Right$(sPath, 4)) = ".lnk" Then
                    IsLink = True
                    sLinkPath = GetLnkTarget(sDir & sPath)
                    bIsLinkFolder = (PathIsDirectory(sLinkPath) = vbDirectory)
                Else
                    IsLink = False
                End If
               
                If Not (tWFD.dwFileAttributes And vbDirectory) = vbDirectory And Not bIsLinkFolder Then

                    If IsLink Then
                        If FileExist(sLinkPath) Then
                            sFile = GetPathName(sLinkPath)
                        Else
                            sFile = vbNullString
                        End If
                    Else
                        sFile = sPath
                    End If
                    
                    If IsInFilter(UCase(sFile)) Then
                    
                        Call SHGetFileInfo(sDir & sPath, ByVal 0&, tSHFI, Len(tSHFI), SHGFI_ICON Or SHGFI_SMALLICON)
                        If IsLink Then
                            IDList = GetIDListFromPath(sLinkPath)
                        Else
                            IDList = GetIDListFromPath(sDir & sPath)
                        End If
                                                
                        index = UBound(ArrFiles)
                        
                        With ArrFiles(index)
                            .Attribute = tWFD.dwFileAttributes
                            .FileName = sFile
                            .IDList = IDList
                            .hIcon = tSHFI.hIcon
                            
                        End With
                        
                        ReDim Preserve ArrFiles(index + 1)
                    End If
                Else
    
                    sFolder = sPath
                    
                    If sFolder <> "." And sFolder <> ".." Then
                    
                        Call SHGetFileInfo(sDir & sFolder, ByVal 0&, tSHFI, Len(tSHFI), SHGFI_ICON Or SHGFI_SMALLICON)
                        
                        If IsLink Then
                            IDList = GetIDListFromPath(sLinkPath)
                            bIsLinkFolder = False
                            sLinkPath = vbNullString
                        Else
                            IDList = GetIDListFromPath(sDir & sFolder)
                        End If

                        index = UBound(ArrFolders)
                        
                        With ArrFolders(index)
                            .Attribute = tWFD.dwFileAttributes
                            .FileName = sFolder
                            .IDList = IDList
                            .hIcon = tSHFI.hIcon
                            .HasContent = CarpetaContieneArchivosYSubcarpetas(sDir & sPath)
                        End With
                        
                        ReDim Preserve ArrFolders(index + 1)
                    End If
                End If
            End If
  
            lRet = FindNextFile(lhSearch, tWFD)

        Loop While lRet
        
        Call FindClose(lhSearch)
        
        FindFiles = True
    End If

End Function

Private Function IsInFilter(ByVal sFile As String) As Boolean
    Dim i As Long
    
    If sFile = vbNullString Then Exit Function
    If mStrFilter = "*.*" Then IsInFilter = True: Exit Function

    For i = 0 To UBound(ArrFilter)
        If LCase(sFile) Like Trim(LCase(ArrFilter(i))) Then
           IsInFilter = True
           Exit Function
        End If
    Next
End Function


