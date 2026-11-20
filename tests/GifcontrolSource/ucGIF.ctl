VERSION 5.00
Begin VB.UserControl ucGIF 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   CanGetFocus     =   0   'False
   ClientHeight    =   990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   ScaleHeight     =   66
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   83
   Begin VB.Timer tmrGIF 
      Enabled         =   0   'False
      Left            =   405
      Top             =   285
   End
End
Attribute VB_Name = "ucGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************
'人人为我，我为人人
'枕善居汉化收藏整理
'发布日期：05/04/08
'描  述：一个Gif 动画播放控件
'网  站：http://www.mndsoft.com/blog/
'e-mail：mnd@mndsoft.com
'OICQ  : 88382850
'****************************************************************************
Option Explicit
' 99% of the research done on this project came from one website:

' Outside Code sources used to help understand/improve Reader & included within:
' DoBitWise, SkipBlock, & most GIF Block Constants
'   borrowed & modified from Vlad Vissoultchev's project @
' PictureFromByteStream borrowed from Kamilche's project @
' Thanx to Jim K, whose GIF Information project, inspired me to complete this
'   project that I've had sitting on the backburner for ages. Wanted a way to
'   reliably include animated GIFs in DCs like menus & buttons.

' Credits provided. Now to the code.

' An excellent project by Vlad, as mentioned above, was the best I've seen to date
' However, that project & just about every other one I've viewed had 3 issues I
' thought could have been correctly accounted for without working around them....
' 1. Other projects did not handle transparent colors correctly. In fact, the
'    authors don't use that color in any of the routines. Vlad's project did this correctly
' 2. All projects either write the individual GIF frames to temporary files or create
'    an array of stdPics to be used when displaying the frames. I wanted a way to
'    only use one stdPic regardless how many frames the GIF included. Thought here
'    was that GIFs containing dozens & dozens of frames unnecessarily wasted resources
' 3. No GIF routines I have reviewed identified and respected the Loop count imposed
'    within the GIF.

' Since this class does no drawing, you are responsible to draw everything. In addition,
' it is your responsibility for timers. The property FrameImage(Index) will pass a
' stdPicture that can be used to apply to a picturebox, imagecontrol, or for blt'ing.

' The usercontrol version of this routine does everything; all you need to do is supply the GIF filename.

' Calling simple properties will let you know the frame's relative position to the
' logical window, its transparency color, delay time, and other properties to include
' the disposal method.  About this method. It is very important that you understand
' the disposal method so you can draw each frame appropriately.

' Disposal method values:  Only 2 values require action:
' Disposal=2 (Erase). Here you are required to erase the DC, but only erase the
'   portion of the DC that the currently displayed frame is occupying. Erasing means
'   to repaint the DC background.
' Disposal=3 (Replace). Here you are required to replace the area occupied by the
'   currently displayed frame with what was there previously. This may include part
'   of the previously displayed frame. This is not the same as Erase & generally
'   requires a separate DC to capture drawing DC before each frame is drawn.

' Because a Disposal value of 3 requires additional work by the coder, it has its own
' property to inform you whether or not you need to set up an offscreen buffer just
' to handle this disposal method: The property name is GifBkBufferRqd

' The downside of most animated GIF routines (including an earlier one by me) was that
' the authors assumed simply drawing the GIF frame was all that was needed to be done.
' But a good routine must handle the disposal of the frame properly too.

Public Event AnimationLoopExpired()
Public Event AnimationProgress(ByVal currentFrame As Long)
Public Event AnimationLoopComplete(ByVal loopsRemaining As Long)

Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_FRAMECHANGED As Long = &H20
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_CLIENTEDGE As Long = &H200&
Private Const WS_BORDER As Long = &H800000
Private Const WS_DLGFRAME As Long = &H400000

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetProp Lib "user32.dll" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long

' Standard Window UDTs
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type Size
    Width As Integer
    Height As Integer
End Type

' Custom UDTs
Private Type GIFcoreProperties
    BkgColor As Long            ' Logical window bkgColor per GIF
    vWindow As Size             ' Logical window width.height
    ScaleCx As Single           '
    ScaleCy As Single           '
    FileNum As Integer          ' File Number for opened GIF
    Loops As Integer            ' Nr loops defined in GIF (can be infinite)
    UsesBkBuffer As Boolean     ' Used Disposal rule #3. Read more above
    HdrOffset As Long           ' End of GIF encoded header
    isCached As Byte            ' GIF file is temp file vs user harddrive file
    gifHdr() As Byte            ' cached header data
    Version As String * 6       ' version Gif89a/Gif87a
    fileName As String          ' currently accessed GIF
End Type
Private Type bkBuffDC
    DC As Long
    oldBmp As Long
    disp3Bmp As Long
    cusBkgBmp As Long
End Type
Public Enum GIFtransStateEnum
    NotTransparent = 0
    AllTransaprent = 1
    PartialTransparent = 2
End Enum
Public Enum ScaleGIFConstants
    NoResizeFrames = 0
    ScaleFrames = 1
    StretchFrames = 2
End Enum
Public Enum StretchModeConstants
    scaleLight = 0
    scaleHeavy = 1
End Enum
Public Enum ucBorderStyleConstants
    bdrNone = 0
    bdrFlat = 1
    bdrRaised = 2
    bdrSunken = 3
End Enum

' DCs only used if GIF includes disposal code of 3 or if user opts for
' a custom background, scaling images or forcing transparencies
Private bkBuff As bkBuffDC
Private bkCustBuff As bkBuffDC

Private gifProps As GIFcoreProperties
Private uMinDelay As Integer     ' any delay less than this value will use this value
Private uScaleFactor As Integer  ' frame scale options
Private uScaleMode As Integer    ' scale mode (either VB's Render or StretchBlt)
Private uAutoSize As Boolean     ' window size lock
Private uBackColor As Long       ' solid backcolor
Private uBorder As Integer       ' border style
Private uAniFrame As Long        ' during animation: the current frame
Private uAniLoops As Integer     ' during animation: the current loop
Private colFrames As Collection  ' collection of frame classes
Private curFileNum As Integer    ' file number for open GIF file

' ***********************************************************************************
' Following are Public Functions & Read/Write Properties. Read-Only Properties at end
' ***********************************************************************************

Private Sub UserControl_Initialize()
    ' global properties at start up
    uMinDelay = 50
    uAutoSize = True
    uBackColor = vbButtonFace
End Sub

Private Sub UserControl_Terminate()
    ' close any open handles, delete DCs, basic clean up stuff
    UnloadGIF
End Sub

Private Sub UserControl_Resize()
    If colFrames Is Nothing Then Exit Sub
    
    Dim bdrSize As Long
    bdrSize = UserControl.Width \ Screen.TwipsPerPixelX - UserControl.ScaleWidth
    
    If uAutoSize Then
    
        Dim bResize As Boolean
        If UserControl.Width <> (gifProps.vWindow.Width + bdrSize) * Screen.TwipsPerPixelX Then
            bResize = True
        End If
        If UserControl.Height <> (gifProps.vWindow.Height + bdrSize) * Screen.TwipsPerPixelY Then
            bResize = True
        End If
        If bResize Then
            UserControl.Width = (gifProps.vWindow.Width + bdrSize) * Screen.TwipsPerPixelX
            UserControl.Height = (gifProps.vWindow.Height + bdrSize) * Screen.TwipsPerPixelY
            Exit Sub
        End If
        
    End If
    
    CalculateScaleFactor
    SetupDC    ' ensure any bitmaps are sized accordingly
    ShowFrame 1 ' show first frame

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "nrLoops", gifProps.Loops, 0
    .WriteProperty "autoSize", uAutoSize, True
    .WriteProperty "bkColor", uBackColor, vbButtonFace
    .WriteProperty "minDelay", uMinDelay, 50
    .WriteProperty "scaleFactor", uScaleFactor, 0
    .WriteProperty "scaleMode", uScaleMode, 0
    
    ' well, can't save a animated GIF like a normal stdPic. If you try,
    ' you'll only get the 1st frame. We save it as a byte array
    Dim fileBytes() As Byte
    If gifProps.FileNum Then
        On Error Resume Next
        ReDim fileBytes(0 To LOF(gifProps.FileNum) - 1)
        .WriteProperty "fileLen", UBound(fileBytes)
        Get #gifProps.FileNum, 1, fileBytes
        .WriteProperty "fileData", fileBytes
        Erase fileBytes
        If Err Then .WriteProperty "fileLen", 0
    Else
        .WriteProperty "fileLen", 0
    End If
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        ' well, can't save a animated GIF like a normal stdPic. If you try,
        ' you'll only get the 1st frame. We read it as a byte array
        Dim nrBytes As Long, tmpFile As String
        Dim fileBytes() As Byte, tByte(0) As Byte
        
        nrBytes = .ReadProperty("fileLen", 0)
        If nrBytes Then
            On Error Resume Next
            ReDim fileBytes(0 To nrBytes)
            fileBytes = .ReadProperty("fileData", tByte)
            gifProps.FileNum = FreeFile()
            tmpFile = GetTempFileName()
            Open tmpFile For Binary As #gifProps.FileNum
            Put #gifProps.FileNum, , fileBytes
            Erase fileBytes
            Close #gifProps.FileNum
            gifProps.FileNum = 0
            If Err = 0 Then
               If LoadGIF(tmpFile) Then gifProps.isCached = 1
            End If
        End If
        gifProps.Loops = .ReadProperty("nrLoops", 0)
        uAutoSize = .ReadProperty("autoSize", True)
        uBackColor = .ReadProperty("bkColor", vbButtonFace)
        uMinDelay = .ReadProperty("minDelay", 50)
        uScaleFactor = .ReadProperty("scaleFactor", 0)
        uScaleMode = .ReadProperty("scaleMode", 0)
    End With
End Sub

' ***********************************************************************************
' Following are Public Functions & Read/Write Properties. Read-Only Properties at end
' ***********************************************************************************

'/==================================================================================
'   FUNCTION RETURNS BEST SCALED SIZE FOR A WINDOW FOR THE LOADED GIF
'/==================================================================================

Public Sub ScaleLogWinToSize(newWidth As Long, newHeight As Long, _
    Optional bCanShrink As Boolean, Optional BorderWidth As Long)

' Parameters
' newWidth :: desired width of usercontrol
' newHeight :: desired height of usercontrol
' bCanShrink :: optional & used for a return value
'   If set to True on exit, the full-size GIF can fit within the passed size
'   If set to False then, the full-size GIF must be resized to fit
' BorderWidth: the total pixels for left/right border edges
'   Helpful if you want to ensure you size the control so entire GIF can
'   be displayed along with a border edge other than none.

    If gifProps.vWindow.Width = 0 Or gifProps.vWindow.Height = 0 Then Exit Sub

    Dim lRatio1 As Single, lRatio2 As Single
    Dim OversizedCx As Boolean, OversizedCy As Boolean
    
    OversizedCx = (newHeight >= gifProps.vWindow.Height)
    OversizedCy = (newWidth >= gifProps.vWindow.Width)
    
    bCanShrink = (OversizedCx = True And OversizedCy = True)
    
    lRatio1 = (newWidth / gifProps.vWindow.Width)
    lRatio2 = (newHeight / gifProps.vWindow.Height)
    If lRatio2 < lRatio1 Then lRatio1 = lRatio2
    newWidth = lRatio1 * gifProps.vWindow.Width
    newHeight = lRatio1 * gifProps.vWindow.Height

    BorderWidth = UserControl.Width \ Screen.TwipsPerPixelX - UserControl.ScaleWidth

End Sub
'/==================================================================================
' 3 FUNCTIONS USED TO RETRIEVE, UPDATE, & CLEAR CONTROL'S DC FOR CUSTOM BACKGROUNDS
'/==================================================================================
    Public Function GetCustomBkgDC() As Long
    ' returns the custom DC so user can draw on to it
    ' Do not cache this value. It can change
    
    ' IMPORTANT: Pass it back before drawing, otherwise leaks may occur
    
    SetupDC True, False     ' set up needed DCs & bitmaps
    SetupDC False, True
    
    SelectObject bkCustBuff.DC, bkBuff.cusBkgBmp    ' select bitmap & pass DC
    GetCustomBkgDC = bkCustBuff.DC
    End Function
    
    Public Sub SetCustomBkgDC()
    ' user finished drawing on DC - if you don't pass DC back, memory leaks
    If bkCustBuff.DC Then SelectObject bkCustBuff.DC, bkCustBuff.cusBkgBmp
    End Sub
    
    Public Sub RemoveCustomBkgDC()
    ' remove the custom DC & bitmaps
    If bkCustBuff.DC Then
        DeleteObject bkBuff.cusBkgBmp
        bkBuff.cusBkgBmp = 0
        With bkCustBuff
            DeleteObject SelectObject(.DC, .oldBmp)
            .cusBkgBmp = 0
            .oldBmp = 0
            DeleteDC .DC
            .DC = 0
        End With
        UserControl.Cls
    End If
    
    End Sub
'/==================================================================================
'  PROPERTY USED TO LOAD A GIF DURING DESIGN TIME.
'   Note that this does not return failure.
'   In code you should use the LoadGIF & UnloadGIF functions
'/==================================================================================
Public Property Get gifFileName() As String
Attribute gifFileName.VB_ProcData.VB_Invoke_Property = "ppGIFpreview"
    If gifProps.fileName = "" Then
        gifFileName = "{none}"
    Else
        gifFileName = gifProps.fileName
    End If
End Property
Public Property Let gifFileName(gifFileName As String)
    If gifFileName = "" Then
        UnloadGIF
    Else
        LoadGIF gifFileName
    End If
End Property
'/==================================================================================
'   PROPERTY TO ENABLE/DISABLE ANIMATION
'/==================================================================================
Public Property Get Animate() As Boolean
    Animate = Len(tmrGIF.Tag) > 0
End Property
Public Property Let Animate(vAnimate As Boolean)
    DoAnimation Not vAnimate
End Property
'/==================================================================================
'   ROUTINE TO PAUSE AND RESUME ANIMATION
'   Note: When using a custom bkg, always pause before resizing usercontrol
'/==================================================================================
Public Sub PauseAnimation(bResume As Boolean)
    If Animate Then
        tmrGIF.Enabled = bResume
        If Not bResume Then tmrGIF.Interval = tmrGIF.Interval / 2 + 1
    End If
End Sub
'/==================================================================================
'   PROPERTY TO LOCK WINDOW SIZE SO ENTIRE FULL-SIZE GIF IS DISPLAYED COMPLETELY
'   If stretching GIF is desired, this property must be set to False
'/==================================================================================
Public Property Get AutoSize() As Boolean
    AutoSize = uAutoSize
End Property
Public Property Let AutoSize(vValue As Boolean)
    If uAutoSize <> vValue Then
        uAutoSize = vValue
        uScaleFactor = NoResizeFrames
        If Not colFrames Is Nothing Then Call UserControl_Resize
    End If
End Property
'/==================================================================================
'   PROPERTY TO SET THE BACKCOLOR OF THE USERCONTROL.
' This generally has no effect on non-transparent GIFs
'/==================================================================================
Public Property Let BackColor(vColor As OLE_COLOR)
    uBackColor = vColor
    UserControl.BackColor = vColor
    ShowFrame 1
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = uBackColor
End Property
'/==================================================================================
'   PROPERTY GETS/SETS THE NR OF LOOPS A GIF WILL COMPLETE BEFORE STOPPING
'   0 :: indefinite
'  -1 :: indicates single frame GIF
'   # :: indicates number of loops
'/==================================================================================
Public Property Get gifMaxLoops() As Integer
    gifMaxLoops = gifProps.Loops
End Property
Public Property Let gifMaxLoops(newMax As Integer)

    If newMax < -1 Then newMax = 1
    gifProps.Loops = newMax
    uAniLoops = newMax
    gifProps.Loops = newMax
   
End Property
'/==================================================================================
'   PROPERTY GETS/SETS MINIMUM FRAME DELAY
'   Any frame delay that is < this value will be set to this value when animating
'/==================================================================================
Public Property Get gifMinFrameDelay() As Integer
    gifMinFrameDelay = uMinDelay
End Property
Public Property Let gifMinFrameDelay(newDelay As Integer)
    If newDelay < 5 Then
        newDelay = 5
    Else
        If newDelay > 60000 Then newDelay = 60000
    End If
    uMinDelay = newDelay
End Property
'/==================================================================================
'   PROPERTY GETS/SETS CURRENT RESIZE METHOD
'   0 :: Frames are not resized - a 1:1 ratio
'   1 :: Frames are scaled proportionally
'   2 :: Frames are scaled unproportionally
'/==================================================================================
Public Property Get Stretch() As ScaleGIFConstants
    Stretch = uScaleFactor
End Property
Public Property Let Stretch(vValue As ScaleGIFConstants)
    If vValue > -1 And vValue < 3 Then
        If vValue <> uScaleFactor Then
            uScaleFactor = vValue
            Call UserControl_Resize
        End If
    End If
End Property
'/==================================================================================
'   PROPERTY GETS/SETS TYPE OF RESIZING ROUTINE TO USE
'   0 :: Uses VBs StdPicture.Render function
'   1 :: Converts each frame to bitmap & uses StretchBlt
'   Quality dependent upon source image and whether enlarging or shrinking the image
'/==================================================================================
Public Property Get StretchMethod() As StretchModeConstants
    StretchMethod = uScaleMode
End Property
Public Property Let StretchMethod(vMethod As StretchModeConstants)
    If vMethod = scaleLight Or vMethod = scaleHeavy Then
        uScaleMode = vMethod
        Call UserControl_Resize
    End If
End Property
'/==================================================================================
'   PROPERTY GETS/SETS TYPE OF BORDER STYLE FOR THE CONTROL
'/==================================================================================
Public Property Get BorderStyle() As ucBorderStyleConstants
    BorderStyle = uBorder
End Property
Public Property Let BorderStyle(vStyle As ucBorderStyleConstants)
    
    If vStyle < bdrNone Or vStyle > bdrSunken Then Exit Property
    If vStyle <> uBorder Then
        uBorder = vStyle
        
        Dim lStyle As Long, lStyleEx As Long
        lStyleEx = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
        lStyle = GetWindowLong(UserControl.hwnd, GWL_STYLE)
        
        Select Case uBorder
        Case bdrNone
            lStyleEx = lStyleEx And Not WS_EX_CLIENTEDGE
            lStyle = lStyle And Not WS_BORDER
            lStyle = lStyle And Not WS_DLGFRAME
        Case bdrFlat
            lStyleEx = lStyleEx And Not WS_EX_CLIENTEDGE
            lStyle = lStyle Or WS_BORDER
            lStyle = lStyle And Not WS_DLGFRAME
        Case bdrRaised
            lStyleEx = lStyleEx And Not WS_EX_CLIENTEDGE
            lStyle = lStyle And Not WS_BORDER
            lStyle = lStyle Or WS_DLGFRAME
        Case bdrSunken
            lStyleEx = lStyleEx Or WS_EX_CLIENTEDGE
            lStyle = lStyle And Not WS_BORDER
            lStyle = lStyle And Not WS_DLGFRAME
        End Select
        
        SetWindowLong UserControl.hwnd, GWL_EXSTYLE, lStyleEx
        SetWindowLong UserControl.hwnd, GWL_STYLE, lStyle
        SetWindowPos UserControl.hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
        Call UserControl_Resize
    
    End If
End Property
'/==================================================================================
'   FUNCTION CALLED FROM YOUR CODE
'   It will load the passed GIF file & return True/False for success/failure
'/==================================================================================
Public Function LoadGIF(gifFileName As String) As Boolean

    On Error Resume Next
    If Len(Dir$(gifFileName)) = 0 Then Exit Function
    If FileLen(gifFileName) < 20 Then Exit Function
    If Err Then
        Err.Clear
        Exit Function
    End If
    
    ' Flags used to help identify BIT occurences
    Const flagDisposal As Byte = &H1C
    Const flagClrTblSize As Byte = &H7
    Const flagClrTblUsed As Byte = &H80
    Const gTableLoc As Long = &HE
    Dim lTableLoc As Long
    
    Dim gTableCount As Byte         ' nr when applied ^2 indicates nr colors in gif table
    Dim lTableCount As Byte         ' same as above for local tables vs global tables
    Dim gByte As Byte               ' generic use byte
    Dim gifString As String         ' generic use string
    Dim fColor As Long              ' frame transparent color
    Dim gTableUsed As Boolean       ' global table in effect
    Dim maxLength As Long           ' max size of largest image frame in bytes
    Dim frameLength As Long         ' individual frame size
    Dim frameCount As Integer       ' number of frames in the gif
    Dim I As Integer                ' generic use (counter)
    Dim cFrame As clsGIFframe
    
    On Error GoTo ExitReadRoutine
    ' work in a backup in case user passes corrupted file
        Dim tmpProps As GIFcoreProperties
        Dim tmpFrames As New Collection
        
    
    With tmpProps
    
        ' load as binary
        curFileNum = FreeFile()
        Open gifFileName For Binary Access Read As #curFileNum
    
        ' reset to defaults
        .Loops = -1
        .UsesBkBuffer = False
        fColor = -1
        
        ' read gif key data
        .Version = ReadGifFile(6, True)
        If LCase(.Version) <> "gif89a" And LCase(.Version) <> "gif87a" Then GoTo ExitReadRoutine
        
        .vWindow.Width = ReadGifFile(2)
        .vWindow.Height = ReadGifFile(2)
    
        ' see if a global color table is in use
        gByte = ReadGifFile()
        gTableUsed = DoBitwise(gByte + 0, flagClrTblUsed)
    
        ' get number of colors in the table. Used to calculate bytes used for table
        gTableCount = DoBitwise(gByte + 0, flagClrTblSize)
    
        ' let's get the bkground color of the gif logical window
        If gTableUsed Then
            ' determine the background index
            .BkgColor = ReadGifFile()
            ColorFromTable &HE, .BkgColor, gTableCount
            If .BkgColor < 0 Then .BkgColor = 0
            ' move the file pointer to the end of the color table
            Seek #curFileNum, gTableLoc + 3 * 2 ^ (gTableCount + 1)
        Else
            ' no global color table; probably uses local color tables
            Seek #curFileNum, 14
            .BkgColor = 0
            gTableCount = 0
        End If
        
    End With
    
    
    ' loop thru the entire file to find all the images & other key data
    With tmpProps
    
        Do
            Select Case ReadGifFile()    ' read a single byte
            Case 0  ' block terminators
            
            Case 33 'Extension Introducer
                ' different extensions handled differently
                
                ' read the extension type
                Select Case ReadGifFile()
                
                Case 255    ' application extension
                    If tmpFrames.Count = 0 Then
                        ' the Netscape extension should always be before any images
                        ' Get the length of extension: will always be 11
                        gByte = ReadGifFile()
                        ' read first 8 bytes & see if it is a netscape extension
                        gifString = UCase(ReadGifFile(8, True))
                        If gifString = "NETSCAPE" Then
                            ' within the data, we can extract the number of loops
                            ' an animated gif is suppose to run. Move ahead 3 bytes
                            Seek curFileNum, Seek(curFileNum) + 3
                            ' now get the data block byte count
                            gByte = ReadGifFile()
                            If gByte = 3 Then   ' valid netscape extension
                                ' move ahead one byte & the next two is the loop count
                                Seek curFileNum, Seek(curFileNum) + 1
                                .Loops = ReadGifFile(2)
                                If .Loops < 1 Then .Loops = 0
                            Else
                                ' not valid netscape extension. Move back a byte
                                ' and we will simply skip the rest of this block
                                Seek curFileNum, Seek(curFileNum) - 1
                            End If
                        Else
                            ' not netscape? Simply finish the block & skip the rest
                            Seek curFileNum, Seek(curFileNum) + 3
                        End If
                    End If
                    SkipGifBlock
                
                Case 254    ' Comment Block
                    ' this contains creator comments that can be read by opening
                    ' the gif in notepad. Mostly at bottom of gif, but can be anywhere
                    SkipGifBlock
                    
                Case 249    ' Graphic Control Label
                    ' this is the individual frame data
                    If .HdrOffset = 0 Then
                        ' first frame; cache where header ends
                        .HdrOffset = Seek(curFileNum) - 2
                    End If
                    ' begin a new frame structure
                    Set cFrame = New clsGIFframe
                    With cFrame
                        ' cahce offset where frame starts & move ahead 2 control bytes
                        .byteOffset(True) = Seek(curFileNum) - 1
                        Seek curFileNum, .byteOffset(True) + 2
                        ' get next byte which contains several properties
                        gByte = ReadGifFile()
                        
                        ' calculate how frame is cleared after it is shown
                        .FrameDisposal = DoBitwise(gByte + 0, flagDisposal)
                        If .FrameDisposal = 3 Then tmpProps.UsesBkBuffer = True
                        
                        ' determine if frame uses transparency or not & which color
                        .isTransparent = DoBitwise(gByte + 0, &H1)
                        
                        ' how long does frame stay before being disposed
                        .frameDelay = ReadGifFile(2) * 10
                        
                        ' if transparent then retrieve index into color table
                        If .isTransparent Then fColor = ReadGifFile()
                            
                        ' move ahead the last byte to finish the block
                        Seek curFileNum, Seek(curFileNum) + 2 + .isTransparent
                        
                    End With
                    
                Case 1  ' Plain Text Extension
                    ' read up on these; just haven't seen any to play with
                    SkipGifBlock
                    
                Case Else   ' Unknown extension
                    SkipGifBlock
                End Select
                
            Case 44 ' Image Descriptor
                ' location of image within logical window
                If .HdrOffset = 0 Then
                    ' Gif87a won't have a block249 so the framecount won't be initialized
                    ' Noticed this with some Gif89a files too
                    Set cFrame = New clsGIFframe
                    cFrame.byteOffset(True) = Seek(curFileNum) - 1
                    .HdrOffset = cFrame.byteOffset(True) - 1
                    ' however above type GIFs are never animated
                End If
                    
                With cFrame
                    .FrameLeft = ReadGifFile(2)
                    .FrameTop = ReadGifFile(2)
                    .FrameWidth = ReadGifFile(2)
                    .FrameHeight = ReadGifFile(2)
                    ' next byte indicates if local color table used
                    ' Need to account for the space in file to move through blocks w/o errors
                    gByte = ReadGifFile()
                    If DoBitwise(gByte + 0, flagClrTblUsed) Then
                        ' local color table used.
                        lTableCount = DoBitwise(gByte, flagClrTblSize)
                        lTableLoc = Seek(curFileNum)
                        Seek curFileNum, lTableLoc + 3 * 2 ^ (lTableCount + 1) + 1
                    Else
                        ' no local table.
                        Seek curFileNum, Seek(curFileNum) + 1
                    End If
                    ' extract transparency color from color table
                    If .isTransparent Then
                        If gTableCount Then
                            ColorFromTable &HE, fColor, gTableCount
                        ElseIf lTableCount Then
                            ColorFromTable lTableLoc, fColor, lTableCount
                        End If
                        .FrameTransparentColor = fColor
                    Else
                        .FrameTransparentColor = -1
                    End If
                    ' done with block, skip the rest & increment our frame count
                    SkipGifBlock
                    .byteOffset(False) = Seek(curFileNum)
                End With
                tmpFrames.Add cFrame
                Set cFrame = Nothing
                frameCount = frameCount + 1
                fColor = -1
                
            Case 59 ' Trailer (end of images)
                ' Although more images may exist, this flag tells us not to use any others
                
                ' finalize the last frame's end offset
                Exit Do
            Case Else
                ' shouldn't happen
                SkipGifBlock
            End Select
        Loop
          
    End With
          
    ' got this far? good to go
    On Error GoTo 0
    ' ensure no animation
    DoAnimation True
    ' close previous file handle if applicable
    If gifProps.FileNum Then Close #gifProps.FileNum
          
    ' clear & transfer byte offsets
    ' clear & transfer frame classes
    Set colFrames = tmpFrames
    Set tmpFrames = Nothing
    ' clear and transfer key core GIF properties
    gifProps = tmpProps
    gifProps.FileNum = curFileNum
          
    With gifProps
        ' cache the header bytes for faster animated frame creation
        ReDim .gifHdr(0 To .HdrOffset - 1)
        Get #.FileNum, 1, .gifHdr()
        
        ' do some final checks here
        
        For I = 1 To frameCount
            Set cFrame = colFrames(I)
            ' calculate max bytes needed for any frame
            frameLength = cFrame.byteOffset(False) - cFrame.byteOffset(True)
            If frameLength > maxLength Then maxLength = frameLength
            ' the logical window size provided in gif should be large enough
            ' to show all frames; however, other documentation on the web
            ' suggests not to rely on it. So we will double check....
            If cFrame.FrameWidth > .vWindow.Width Then .vWindow.Width = cFrame.FrameWidth
            If cFrame.FrameHeight > .vWindow.Height Then .vWindow.Height = cFrame.FrameHeight
        Next
        
        If colFrames.Count > 1 Then
            If .Loops < 0 Then .Loops = 0
        End If
            
        ' adjust our byte array to account for the max frame byte length
        ReDim Preserve .gifHdr(0 To .HdrOffset + maxLength + 1)
    
        .fileName = gifFileName  ' not saved in usercontrol
        
    End With
    
    Call UserControl_Resize
    
    LoadGIF = True
    PropertyChanged "gifFileName"
    Exit Function
    
    
ExitReadRoutine:
    ' Error occurred reading the file, clean up
    If tmpProps.FileNum Then Close #tmpProps.FileNum
    Set tmpFrames = Nothing
    Err.Clear
End Function
'/==================================================================================
'   FUNCTION UNLOADS A GIF & ALL ASSOCIATED SETTINGS/MEMORY OBJECTS
'/==================================================================================
Public Sub UnloadGIF()
    On Error Resume Next
    
    DoAnimation True    ' Stop animation & close file
    If gifProps.FileNum Then
        Close #gifProps.FileNum
        If gifProps.isCached Then Kill gifProps.fileName
        gifProps.isCached = 0
    End If
    ' clean up
    Set colFrames = Nothing
    ReDim gifProps.gifHdr(0)
    gifProps.fileName = ""
    gifProps.FileNum = 0
    If bkBuff.DC Then
        If bkBuff.disp3Bmp Then DeleteObject SelectObject(bkBuff.DC, bkBuff.oldBmp)
        DeleteDC bkBuff.DC
    End If
    If bkCustBuff.DC Then
        DeleteObject SelectObject(bkCustBuff.DC, bkCustBuff.oldBmp)
        If bkBuff.cusBkgBmp Then DeleteObject bkBuff.cusBkgBmp
        DeleteDC bkCustBuff.DC
    End If
    UserControl.Cls
End Sub



' **********************************************************************************
' PUBLIC PROPERTIES -- READ ONLY
' **********************************************************************************

'/==================================================================================
'   PROPERTY RETURNS THE FRAME CLASS FOR THE PASSED INDEX
'/==================================================================================
Public Property Get Frame(Index As Long) As clsGIFframe
    If Not colFrames Is Nothing Then
        If Index > 0 And Index < colFrames.Count + 1 Then Set Frame = colFrames(Index)
    End If
End Property

'/==================================================================================
'   PROPERTY RETURNS A STDPICTURE FOR THE PASSED FRAME INDEX
'/==================================================================================
Public Property Get FrameImage(Index As Long) As StdPicture
' Return the GIF frame as a stdPicture

' Note: Recently researched & became impressed with the full functionality
' of the stdPic.Render function. Once you are comfortable with it, it can
' handle most graphic formats. It does everything BitBlt, StretchBlt,
' DrawIcon & DrawIconEx can do. For our purposes, the Render function does
' all the transparencies & can even shrink/stretch the individual frames.
' Therefore there is no need for transparent bitmap routines to handle
' transparent GIFs, unless user opts for some custom settings.

    If colFrames Is Nothing Then Exit Property
    If Index < 1 Or Index > colFrames.Count Then Exit Property
        
    ' create the single frame GIF image
    Dim gifData() As Byte, cFrame As clsGIFframe
    Set cFrame = colFrames(Index)
    With cFrame
        ' read the gif image block
        ReDim gifData(0 To .byteOffset(False) - .byteOffset(True))
        Get #gifProps.FileNum, .byteOffset(True), gifData()
        ' now append that to the gifHder
        CopyMemory gifProps.gifHdr(gifProps.HdrOffset), gifData(0), UBound(gifData) + 1
        ' here we will add a Terminate flag
        gifProps.gifHdr(UBound(gifData) + gifProps.HdrOffset + 1) = 59
        ' now send byte array to routine to create stdPic on the fly
    End With
    Set FrameImage = PictureFromByteStream(gifProps.gifHdr, UBound(gifData) + gifProps.HdrOffset + 2)
    
End Property

'/==================================================================================
'   PROPERTY RETURNS THE TOTAL TIME AN ANIMATED GIF NEEDS TO COMPLETE A SINGLE LOOP
'/==================================================================================
Public Property Get gifAnimationTime() As Long ' milliseconds
    If colFrames Is Nothing Then Exit Property
    Dim frameNr As Long, ttlTime As Long
    For frameNr = 1 To colFrames.Count
        If colFrames(frameNr).frameDelay < uMinDelay Then
            ttlTime = ttlTime + uMinDelay
        Else
            ttlTime = ttlTime + colFrames(frameNr).frameDelay
        End If
    Next
    gifAnimationTime = ttlTime
End Property
Public Property Let gifAnimationTime(vNull As Long)
' dummy Let so property shows on property page during design time
End Property
'/==================================================================================
'   PROPERTY RETURNS THE NUMBER OF FRAMES WITHIN A GIF
'/==================================================================================
Public Property Get gifFrameCount() As Long
    If colFrames Is Nothing Then Exit Property
    gifFrameCount = colFrames.Count
End Property
Public Property Let gifFrameCount(vNull As Long)
' dummy Let so property shows on property page during design time
End Property
'/==================================================================================
'   PROPERTY RETURNS THE LOGICAL WINDOW WIDTH NEEDED TO DISPLAY GIF A FULL SIZE
'/==================================================================================
Public Property Get gifLogicalWindowCx() As Integer
    ' the width of the logical window -- does not include any usercontrol borders
    gifLogicalWindowCx = gifProps.vWindow.Width
End Property
'/==================================================================================
'   PROPERTY RETURNS THE LOGICAL WINDOW HEIGHT NEEDED TO DISPLAY GIF A FULL SIZE
'/==================================================================================
Public Property Get gifLogicalWindowCy() As Integer
    ' the height of the logical window -- does not include any usercontrol borders
    gifLogicalWindowCy = gifProps.vWindow.Height
End Property
'/==================================================================================
'   PROPERTY RETURNS THE LOGICAL WINDOW BKG COLOR AS DEFINED BY THE GIF
'/==================================================================================
Public Property Get gifWindowBkgColor() As Long
    ' Note: generally this is not important for animated GIFs; as they will
    ' use disposal rules that will prevent any bkgnd color from showing thru
    gifWindowBkgColor = gifProps.BkgColor
End Property
'/==================================================================================
'   PROPERTY RETURNS THE STATE OF GIF TRANSPARENCY
'   0 :: 100% not transparent
'       (most can be converted to 100% transparency on the fly)
'   1 :: 100% transparent
'       (some gifs use transparency but still have solid backcolors)
'   2 :: mix of transparent & non-transparent frames
'       (converting to 100% transparent not likely to be successful)
'/==================================================================================
Public Property Get gifTransparencyState() As GIFtransStateEnum
    If colFrames Is Nothing Then Exit Property
    Dim frameNr As Long, tState As Integer
    For frameNr = 1 To colFrames.Count
        If colFrames(frameNr).isTransparent Then
            tState = tState Or 1
        Else
            tState = tState Or 2
        End If
    Next
    Select Case tState
    Case 0, 2
        gifTransparencyState = NotTransparent
    Case 1
        gifTransparencyState = AllTransaprent
    Case Else
        gifTransparencyState = PartialTransparent
    End Select
End Property
Public Property Let gifTransparencyState(vNull As GIFtransStateEnum)
' dummy Let so property shows on property page during design time
End Property
'/==================================================================================
'   PROPERTY RETURNS VERSION OF THE GIF
'/==================================================================================
Public Property Get gifVersion() As String
    ' The version of the GIF
    gifVersion = gifProps.Version
End Property
Public Property Let gifVersion(vNull As String)
' dummy Let so property shows on property page during design time
End Property
'/==================================================================================
'   PROPERTY RETURNS THE hWnd FOR THE USERCONTROL.
'   Primarily used for the custom property page
'/==================================================================================
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
'/==================================================================================
'   FUNCTION ALLOWS USERS TO REMOVE FRAMES FROM AN ANIMATED GIF FILE
'   Return value indicates success/failure
'/==================================================================================
Public Function RemoveFrame(Index As Long, Optional bLogWindowResized As Boolean) As Boolean
    
    ' Optional parameter indicates if the logical GIF window dimensions changed
    ' as a result of removing the GIF frame
    
    If colFrames Is Nothing Then Exit Function
    If colFrames.Count < 2 Or Index > colFrames.Count Then Exit Function

    colFrames.Remove Index
    RecalcLogWindow bLogWindowResized
    RemoveFrame = True

End Function
'/==================================================================================
'   FUNCTION ALLOWS USERS TO REPOSITION FRAMES FROM AN ANIMATED GIF FILE
'   Return value indicates success/failure
'/==================================================================================
Public Function MoveFrame(Index As Long, X As Long, Y As Long, Optional bLogWindowResized As Boolean) As Boolean
    
    ' Optional parameter indicates if the logical GIF window dimensions changed
    ' as a result of removing the GIF frame
    
    If colFrames Is Nothing Then Exit Function
    If colFrames.Count < 1 Or Index > colFrames.Count Then Exit Function

    colFrames(Index).FrameTop = Y
    colFrames(Index).FrameLeft = X
    RecalcLogWindow bLogWindowResized
    MoveFrame = True

End Function



' **********************************************************************************
'   INTERNAL HELP FUNCTIONS, PARSING & DRAWING ROUTINES
' **********************************************************************************

'/==================================================================================
'   Calculates logical window size as result of frame removal or placement changes
'/==================================================================================
Private Sub RecalcLogWindow(bChanged As Boolean)
    Dim fNr As Long, lWindow As Size
    For fNr = 1 To colFrames.Count
        If colFrames(fNr).FrameWidth + colFrames(fNr).FrameLeft > lWindow.Width Then
            lWindow.Width = colFrames(fNr).FrameWidth + colFrames(fNr).FrameLeft
        End If
        If colFrames(fNr).FrameHeight + colFrames(fNr).FrameTop > lWindow.Height Then
            lWindow.Height = colFrames(fNr).FrameHeight + colFrames(fNr).FrameTop
        End If
    Next
    bChanged = Not (lWindow.Height = gifProps.vWindow.Height And lWindow.Width = gifProps.vWindow.Width)
    gifProps.vWindow = lWindow
    If bChanged = True Then Call UserControl_Resize

End Sub
'/==================================================================================
'   Calculate relative value of byte depending on passed Mask
'/==================================================================================
Private Function DoBitwise(ByVal lFlags As Long, ByVal lMask As Long) As Long
    If lMask > 0 Then
        DoBitwise = (lFlags And lMask)
        Do While (lMask And 1) = 0
            lMask = lMask \ 2
            DoBitwise = DoBitwise \ 2
        Loop
    End If
End Function
'/==================================================================================
'   Read thru bytes until a zero-byte Block Terminator is found
'/==================================================================================
Private Sub SkipGifBlock()
    Dim curByte As Byte, curLoc As Long
    curByte = ReadGifFile()         ' current byte value
    
    Do While curByte > 0
        Seek curFileNum, Seek(curFileNum) + curByte
        curByte = ReadGifFile()
    Loop
End Sub
'/==================================================================================
'   Read one or more bytes from the open gif file
'/==================================================================================
Private Function ReadGifFile(Optional nrBytes As Long = 1, Optional isASCII As Boolean) As Variant

    If LOF(curFileNum) < Seek(curFileNum) + nrBytes - 1 Then
        Err.Raise 53, "ReadGifFile", "End of File"
        Exit Function
    End If
    
    If isASCII Then
        Dim sRtn As String
        sRtn = String$(nrBytes, Chr$(0))
        Get #curFileNum, , sRtn
        ReadGifFile = sRtn
    Else
        Dim dBytes() As Byte, rtnValL As Long
        ReDim dBytes(0 To nrBytes - 1)
        
        Get #curFileNum, , dBytes
        CopyMemory rtnValL, dBytes(0), nrBytes
        ReadGifFile = rtnValL
    End If
End Function
'/==================================================================================
' Calculate frame transparency color
'/==================================================================================
Public Sub ColorFromTable(seekIdx As Long, tIndex As Long, nrColors As Byte)

    ' this appears to be similar algorithm MS Paint uses:
    ' If the index is invalid, use black as the transparency color
    If tIndex < 0 Then Exit Sub
        ' check next to see if index falls in color table range
    If tIndex * 3 > 2 ^ (nrColors * 3 + 1) - 2 Then
        tIndex = 0
        Exit Sub
    End If
    
    Dim curLoc As Long
    ' get current file position so we can reset pointer later
    curLoc = Seek(curFileNum)
    
    ' now get the 3 color bytes & convert to a long value
    Seek curFileNum, seekIdx + tIndex * 3 + 0
    tIndex = ReadGifFile(3)
    
    ' reset pointer
    Seek curFileNum, curLoc
End Sub
'/==================================================================================
' Convert a byte array to a stdPicture
'/==================================================================================
Private Function PictureFromByteStream(bytContent() As Byte, nrBytes As Long) As IPicture
    On Error GoTo HandleError
    ' in original form from author.
    ' I modified above parameters to include the nr of bytes in the array
    ' Therefore commented out lines that don't apply to this class
            
    ' Caution: From experience. If you use a similar routine in any of your projects,
    ' passing a byte array that cannot be converted to a stdPic will lock up the routine.

        Dim o_lngLowerBound As Long
        Dim o_lngByteCount  As Long
        Dim o_hMem  As Long
        Dim o_lpMem  As Long
        Dim IID_IPicture(15)
        Dim istm As stdole.IUnknown
        
        If UBound(bytContent) > 0 Then
            'o_lngLowerBound = LBound(bytContent)
            'o_lngByteCount = (UBound(bytContent) - o_lngLowerBound) + 1
            o_lngByteCount = nrBytes
            o_hMem = GlobalAlloc(&H2, o_lngByteCount)
            If o_hMem <> 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    CopyMemory ByVal o_lpMem, bytContent(o_lngLowerBound), o_lngByteCount
                    Call GlobalUnlock(o_hMem)
                    If CreateStreamOnHGlobal(o_hMem, 1, istm) = 0 Then
                        If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                            ' if byte array contains invalid picture bytes, following locks up
                          Call OleLoadPicture(ByVal ObjPtr(istm), o_lngByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                        End If
                    End If
                End If
            End If
        End If
    
    Exit Function
    
HandleError:
'    If Err.Number = 9 Then
'        'Uninitialized array
'        MsgBox "You must pass a non-empty byte array to this function!"
'    Else
'        MsgBox Err.Number & " - " & Err.Description
'    End If
End Function
'/==================================================================================
' Off-screen DCs & bitmaps used if GIF has disposal code of 3 or
'   user opts for custom background vs solid color background
'/==================================================================================
Private Sub SetupDC(Optional ForceSetup As Boolean, Optional forCustBkg As Boolean)
    ' setup an offscreen DC & create bitmap only if needed
    If forCustBkg Then
        If bkCustBuff.DC = 0 Then
            bkCustBuff.DC = CreateCompatibleDC(UserControl.hdc)
        Else
            Exit Sub
        End If
    Else
        If bkBuff.DC = 0 Then bkBuff.DC = CreateCompatibleDC(UserControl.hdc)
        If ForceSetup = True Then Exit Sub
    End If
    
    With bkBuff
        If gifProps.UsesBkBuffer = True And forCustBkg = False Then
                
            If .disp3Bmp Then DeleteObject SelectObject(.DC, .oldBmp)
            .disp3Bmp = CreateCompatibleBitmap(hdc, gifProps.vWindow.Width * gifProps.ScaleCx, gifProps.vWindow.Height * gifProps.ScaleCy)
            .oldBmp = SelectObject(.DC, .disp3Bmp)
            
        Else    ' no disposal code of 3. Clean up if DC already created
            
            If .DC <> 0 And forCustBkg = False Then
                If .disp3Bmp Then DeleteObject SelectObject(.DC, .oldBmp)
                .oldBmp = 0
                .disp3Bmp = 0
                If bkCustBuff.cusBkgBmp = 0 Then
                    DeleteDC .DC
                    .DC = 0
                End If
            End If
        End If
    End With
    
    If bkCustBuff.DC Then
        Dim hPrevBmp As Long
        With bkCustBuff
            If .cusBkgBmp Then DeleteObject SelectObject(.DC, .oldBmp)
            .cusBkgBmp = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
            .oldBmp = SelectObject(.DC, .cusBkgBmp)
        End With
        With bkBuff
            If .cusBkgBmp Then
                hPrevBmp = SelectObject(.DC, .cusBkgBmp)
                BitBlt bkCustBuff.DC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, .DC, 0, 0, vbSrcCopy
                SelectObject .DC, hPrevBmp
                DeleteObject .cusBkgBmp
            End If
            .cusBkgBmp = CreateCompatibleBitmap(UserControl.hdc, UserControl.ScaleWidth, UserControl.ScaleHeight)
            If hPrevBmp Then
                hPrevBmp = SelectObject(.DC, .cusBkgBmp)
                BitBlt .DC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, bkCustBuff.DC, 0, 0, vbSrcCopy
                SelectObject .DC, hPrevBmp
            End If
        End With
    End If
End Sub
'/==================================================================================
' Shows the 1st Frame of a GIF
'/==================================================================================
Public Sub ShowFrame(Index As Long)
    If colFrames Is Nothing Then Exit Sub
    If Animate Then Exit Sub
    
    With colFrames(Index)
        If .isTransparent Then
            UserControl.BackColor = uBackColor
        Else
            UserControl.BackColor = gifProps.BkgColor
        End If
    End With
    DoUpdateRect 0, False, True
    RenderFrame Index
    DoUpdateRect 0, True, False
    

End Sub
'/==================================================================================
'   FUNCTION DRAWS THE GIF IMAGE DEPENDENT UPON GIF DISPOSAL CODES AND USER-OPTIONS
'/==================================================================================
Private Sub RenderFrame(frameNr As Long)

    Const DSna = &H220326 '0x00220326
    ' A pretty good transparent bitmap maker I use in several projects
    ' Modified here to specifically work with the clsGIFframe collection
    
    ' DCs used
    Dim lHDCMem As Long, lHDCscreen As Long
    Dim lHDCsrc As Long, lHDCMask As Long, lHDCcolor As Long
    ' bmps for above DCs
    Dim lBmMask As Long, lBmAndMem As Long, lBmColor As Long
    ' prev bmp references for above DCs
    Dim lBmObjectOld As Long, lBmMemOld As Long, lBmColorOld As Long, lBmSrcOld As Long
    ' measurements
    Dim dCX As Long, dCY As Long, imgMaxW As Long, imgMaxH As Long
    Dim wRect As RECT, hBrush As Long, lMaskColor As Long
    
    ' get core DC, current frame and destination DC
    Dim lBMPsource As StdPicture, lHDCdest As Long
    
    Set lBMPsource = FrameImage(frameNr)
    lHDCscreen = UserControl.hdc
    ' determine which DC to draw to. If using a user-defined background, then
    ' that will always override the actual usercontrol DC
    If bkCustBuff.DC Then lHDCdest = bkCustBuff.DC Else lHDCdest = lHDCscreen
    
    With colFrames(frameNr)
    
        If .ForceTransparency = False Then
            ' No forced transparency, maybe no need to individually transBlt frames
            If uAutoSize = True Or uScaleFactor = NoResizeFrames Then
                ' a 1:1 ratio. Use Render
                lBMPsource.Render lHDCdest + 0, .FrameLeft + 0, .FrameTop + 0, _
                    .FrameWidth + 0, .FrameHeight + 0, 0, lBMPsource.Height, lBMPsource.Width, -lBMPsource.Height, ByVal 0&
                Set lBMPsource = Nothing
                Exit Sub
                    
            ElseIf uScaleMode = scaleLight Then
                ' user opted Render over TransBlting
                lBMPsource.Render lHDCdest + 0, .FrameLeft * gifProps.ScaleCx, .FrameTop * gifProps.ScaleCy, _
                    .FrameWidth * gifProps.ScaleCx, .FrameHeight * gifProps.ScaleCy, _
                    0, lBMPsource.Height, lBMPsource.Width, -lBMPsource.Height, ByVal 0&
                Set lBMPsource = Nothing
                Exit Sub
            End If
            
        End If
        ' got here that means we need to individuall process the GIF frames
        ' and possibly convert/process them as bitmaps

        ' calculate the adjusted width/height of the image depending on scale ratios
        dCX = .FrameWidth * gifProps.ScaleCx
        dCY = .FrameHeight * gifProps.ScaleCy
        
        ' setup DC to hold passed bitmap -- ony gets created once
        If bkBuff.DC = 0 Then SetupDC True, False
        lHDCsrc = bkBuff.DC
        ' select bitmap into DC
        lBmSrcOld = SelectObject(lHDCsrc, lBMPsource.Handle)
    
        If .isTransparent = False And .ForceTransparency = False Then
            ' no transparency needed & it isn't a transparent frame,
            ' therefore, we will simply Blt vs processing completely - quicker
            StretchBlt lHDCdest, .FrameLeft * gifProps.ScaleCx, .FrameTop * gifProps.ScaleCy, _
                dCX, dCY, lHDCsrc, 0&, 0&, .FrameWidth, .FrameHeight, vbSrcCopy
            DeleteObject SelectObject(lHDCsrc, lBmSrcOld)
            Exit Sub
        End If
        
        ' we are going to convert the image & create a transparent bitmap out of it
        
        ' get the mask color & ensure it is valid
        lMaskColor = .FrameTransparentColor
        If lMaskColor < 0 Then
            If lMaskColor = -1 Then ' use top left corner
                lMaskColor = GetPixel(bkBuff.DC, 0, 0)
            Else
                lMaskColor = GetSysColor(.FrameTransparentColor And &HFF&)
            End If
        End If
        
        ' always use the largest measurement when creating the bitmaps
        If dCX > .FrameWidth Then imgMaxW = dCX Else imgMaxW = .FrameWidth
        If dCY > .FrameHeight Then imgMaxH = dCY Else imgMaxH = .FrameHeight
    
        'Create some DCs & bitmaps & select bitmaps into DCs
        lHDCMask = CreateCompatibleDC(lHDCscreen)
        lHDCMem = CreateCompatibleDC(lHDCscreen)
        lHDCcolor = CreateCompatibleDC(lHDCscreen)
        
        lBmColor = CreateCompatibleBitmap(lHDCscreen, imgMaxW, imgMaxH)
        lBmAndMem = CreateCompatibleBitmap(lHDCscreen, imgMaxW, imgMaxH)
        lBmMask = CreateBitmap(imgMaxW, imgMaxH, 1&, 1&, ByVal 0&)
        
        lBmColorOld = SelectObject(lHDCcolor, lBmColor)
        lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
        lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    
    ' ====================== Start working here ======================
    
        BitBlt lHDCMem, 0&, 0&, dCX, dCY, lHDCdest, _
            .FrameLeft * gifProps.ScaleCx, .FrameTop * gifProps.ScaleCy, vbSrcCopy
    
        SetBkColor lHDCcolor, GetBkColor(lHDCdest)
        SetTextColor lHDCcolor, GetTextColor(lHDCdest)
    
        BitBlt lHDCcolor, 0&, 0&, .FrameWidth, .FrameHeight, lHDCsrc, 0, 0, vbSrcCopy
    
        If lMaskColor = 0 And .ForceTransparency = False Then
            ' black transparency, process differently
            hBrush = CreateSolidBrush(vbWhite)
            SetRect wRect, 0, 0, .FrameWidth, .FrameHeight
            FillRect lHDCMask, wRect, hBrush
            DeleteObject hBrush
            lBMPsource.Render lHDCMask + 0, 0&, 0&, .FrameWidth + 0, .FrameHeight + 0, _
                0, lBMPsource.Height, lBMPsource.Width, -lBMPsource.Height, ByVal 0&
        Else
            SetBkColor lHDCcolor, lMaskColor
            SetTextColor lHDCcolor, vbWhite
            BitBlt lHDCMask, 0&, 0&, imgMaxW, imgMaxH, lHDCcolor, 0&, 0&, vbSrcCopy
        End If
    
        SetTextColor lHDCcolor, vbBlack
        SetBkColor lHDCcolor, vbWhite
        BitBlt lHDCcolor, 0, 0, imgMaxW, imgMaxH, lHDCMask, 0, 0, DSna

        StretchBlt lHDCMem, 0, 0, dCX, dCY, lHDCMask, _
            0&, 0&, .FrameWidth, .FrameHeight, vbSrcAnd
    
        StretchBlt lHDCMem, 0&, 0&, dCX, dCY, lHDCcolor, _
            0, 0, .FrameWidth, .FrameHeight, vbSrcPaint
   
        ' transfer result
        BitBlt lHDCdest, .FrameLeft * gifProps.ScaleCx, .FrameTop * gifProps.ScaleCy, _
            dCX, dCY, lHDCMem, 0&, 0&, vbSrcCopy
    
    End With
    'Delete memory bitmaps & DCs
    DeleteObject SelectObject(lHDCsrc, lBmSrcOld)
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    SelectObject lHDCsrc, lBmSrcOld
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor

End Sub
'/==================================================================================
' Internal function to calculate proportion ratios for GIF frames
'/==================================================================================
Private Sub CalculateScaleFactor()

' this function called whenever the usercontrol is resized
    With gifProps
        If uAutoSize Then   ' window cannot be resized beyond GIF logical window size
            .ScaleCx = 1
            .ScaleCy = 1
        Else
            If uScaleFactor = NoResizeFrames Then
                .ScaleCx = 1
                .ScaleCy = 1
            Else
                .ScaleCx = (UserControl.ScaleWidth / .vWindow.Width)
                .ScaleCy = (UserControl.ScaleHeight / .vWindow.Height)
                If uScaleFactor = ScaleFrames Then ' else stretch
                    ' scale requires same ratio for both cx & cy
                    If .ScaleCy < .ScaleCx Then .ScaleCx = .ScaleCy Else .ScaleCy = .ScaleCx
                End If
            End If
        End If
    End With
End Sub
'/==================================================================================
' Used to facilitate flicker-free drawing without using offscreen DCs & bitmaps
'/==================================================================================
Private Sub DoUpdateRect(frameNr As Long, bInvalidate As Boolean, bErase As Boolean)

    ' Typically offscreen dcs & bitmaps used to create flicker free drawing
    ' However, using InvalidateRect can do the job without the added overhead
    '   when used properly. The simple trick is to invalidate every drawing
    '   action by NOT pass True as the last parameter to the API. Once
    '   all drawing is complete, simply pass True & window refreshes.
    '   Side Note: Use .AutoRedraw on your controls.
    
    Dim eRect As RECT
    
    If frameNr = 0 Then
        SetRect eRect, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Else
        With colFrames(frameNr)
            If uAutoSize = True Or uScaleFactor = NoResizeFrames Then
                SetRect eRect, .FrameLeft, .FrameTop, _
                    .FrameLeft + .FrameWidth, .FrameTop + .FrameHeight
            Else
                SetRect eRect, 0, 0, .FrameWidth * gifProps.ScaleCx, gifProps.ScaleCy + .FrameHeight * gifProps.ScaleCy
                OffsetRect eRect, .FrameLeft * gifProps.ScaleCx, .FrameTop * gifProps.ScaleCy
            End If
        End With
    End If
    
    If bErase Then
        Dim meDC As Long, hBrush As Long, brColor As Long, hPrevBmp As Long
        If bkCustBuff.DC Then
            hPrevBmp = SelectObject(bkBuff.DC, bkBuff.cusBkgBmp)
            BitBlt bkCustBuff.DC, eRect.Left, eRect.Top, eRect.Right - eRect.Left, eRect.Bottom - eRect.Top, _
                bkBuff.DC, eRect.Left, eRect.Top, vbSrcCopy
            SelectObject bkBuff.DC, hPrevBmp
        Else
            meDC = UserControl.hdc
            brColor = uBackColor
            If brColor < 0 Then brColor = GetSysColor(brColor And &HFF&)
            hBrush = CreateSolidBrush(brColor)
            FillRect meDC, eRect, hBrush
            DeleteObject hBrush
        End If
    End If
    
    InvalidateRect UserControl.hwnd, eRect, Abs(bInvalidate)
End Sub
'/==================================================================================
'  FUNCTION CREATES A TEMPORARY FILE NAME NEEDED TO WRITE A STORED GIF PROPERTY
'/==================================================================================
Private Function GetTempFileName() As String

Dim sFile As String, sPath As String, tmpIncr As Integer

' patch for VB bug if app is on root drive
If Right$(App.Path, 1) = "\" Then
    sPath = App.Path
Else
    sPath = App.Path & "\"
End If
sFile = sPath & "~tLV(0).gif"
Do While Len(Dir$(sFile)) > 0
    tmpIncr = tmpIncr + 1
    sFile = sPath & "~tLV(" & tmpIncr & ").gif"
Loop
GetTempFileName = sFile
End Function

'/==================================================================================
' TIMER - Only purpose is to erase & draw GIF frames
'/==================================================================================
Private Sub tmrGIF_Timer()

    tmrGIF.Enabled = False  ' disable timer for now
    
    Dim nextFrame As Long
    Dim meDC As Long
   
    ' determine which DC to draw to. If using a user-defined background, then
    ' that will always override the actual usercontrol DC
    If bkCustBuff.DC Then meDC = bkCustBuff.DC Else meDC = UserControl.hdc
    
    ' increment to next frame
    nextFrame = uAniFrame + 1
    
    If uAniFrame < 1 Then
    
        ' first time thru, only ensure we CLS
        DoUpdateRect 0, False, True
    Else
        
        ' do disposal of current frame
        
        ' see if at end of the loop
        If nextFrame > colFrames.Count Then
            
            nextFrame = 1   ' at end, reset counter
            
            If uAniLoops > 0 Then
                ' loops are being adhered to, see if last loop occurred
                If uAniLoops = 1 Then
                    ' yep, reset flags & raise custom event
                    DoAnimation True
                    RaiseEvent AnimationLoopExpired
                    Exit Sub
                End If
                ' loops continuing, subtract counter & raise custom event
                uAniLoops = uAniLoops - 1
                RaiseEvent AnimationLoopComplete(uAniLoops)
            End If
            DoUpdateRect 0, False, True
        
        Else    ' not the first time thru & not the last frame
            
            Select Case colFrames(uAniFrame).FrameDisposal
            Case 2: ' erase area occupied by this frame, using DC back color
                DoUpdateRect uAniFrame, False, True
            
            Case 3 ' replace with what was on the DC before we drew this frame.
                ' This disposal method requires a secondary buffer.
            
                If bkBuff.DC Then   ' this should have always been created automatically
                    With colFrames(uAniFrame)
                        ' transfer the offscreen DC contents to our DC
                        BitBlt meDC, .FrameLeft * gifProps.ScaleCx, .FrameTop * gifProps.ScaleCy, _
                            .FrameWidth * gifProps.ScaleCx, .FrameHeight * gifProps.ScaleCy, bkBuff.DC, 0, 0, vbSrcCopy
                        ' inform windows DC needs updating
                        DoUpdateRect uAniFrame, False, False
                    End With
                End If
            
            Case Else: ' no action required.
                ' However, if forcing transparency on non-transparent GIF,
                ' then we need to erase the previous frame
                If colFrames(uAniFrame).ForceTransparency Then
                    DoUpdateRect uAniFrame, False, True
                End If
            End Select
            
        End If
    
    End If
    
    ' get the next frame's image
    uAniFrame = nextFrame
    If gifProps.UsesBkBuffer Then
        ' result of frames having disposal code of 3
        ' basically we capture our DC before every update
        With colFrames(uAniFrame)
            BitBlt bkBuff.DC, 0, 0, .FrameWidth * gifProps.ScaleCx, .FrameHeight * gifProps.ScaleCy, _
                meDC, .FrameLeft * gifProps.ScaleCx, .FrameTop * gifProps.ScaleCy, vbSrcCopy
        End With
    End If
    
    ' draw the next frame
    RenderFrame uAniFrame
    
    If bkCustBuff.DC Then
        ' when using custom background, blt from there as all drawing was done there
        ' Otherwise all drawing was done directly on the usercontrol DC
        meDC = UserControl.hdc
        BitBlt meDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, bkCustBuff.DC, 0, 0, vbSrcCopy
    End If
    ' send message to refresh the DC
    DoUpdateRect uAniFrame, True, False
    
    ' clean up & set timer for current frame's dealy timeout
    tmrGIF.Interval = colFrames(uAniFrame).frameDelay
    If tmrGIF.Interval < uMinDelay Then tmrGIF.Interval = uMinDelay
    ' raise custom event
    RaiseEvent AnimationProgress(uAniFrame)
    tmrGIF.Enabled = True
End Sub
'/==================================================================================
' ROUTINE SETS UP OR CANCELS ANIMATION
'/==================================================================================
Private Sub DoAnimation(bStop As Boolean)

If bStop Then
    tmrGIF.Enabled = False
    uAniLoops = 0
    uAniFrame = 0
    On Error Resume Next
    If Len(tmrGIF.Tag) Then
        tmrGIF.Tag = ""
        If Not Ambient.UserMode Then
            PropertyChanged "Animate"
            ShowFrame 1
        End If
    End If
Else
    If Not colFrames Is Nothing Then
        If colFrames.Count > 1 Then
            If Ambient.UserMode Then
                If gifProps.Loops > 0 Then uAniLoops = gifProps.Loops Else uAniLoops = 0
            Else
                uAniLoops = 10
            End If
            uAniFrame = 0
            tmrGIF.Interval = 1
            tmrGIF.Tag = "Animating"
            tmrGIF.Enabled = True
        End If
    End If
End If
End Sub
