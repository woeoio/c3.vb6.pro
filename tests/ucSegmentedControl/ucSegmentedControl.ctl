VERSION 5.00
Begin VB.UserControl ucSegmentedControl 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipBehavior    =   0  'None
   DefaultCancel   =   -1  'True
   PropertyPages   =   "ucSegmentedControl.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   ToolboxBitmap   =   "ucSegmentedControl.ctx":0010
   Windowless      =   -1  'True
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   360
   End
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   1320
      Top             =   1800
   End
End
Attribute VB_Name = "ucSegmentedControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------
'Module Name: ucSegmentedControl
'Autor:  Leandro Ascierto
'Web: www.leandroascierto.com
'Published: 13/04/2024
'LastUpdate: 25/04/2024
'Version: 1.0.1
'-------------------------------------------
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function IsBadCodePtr Lib "kernel32" (ByVal lpFn As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetRect Lib "user32" (lpRect As Any, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "Oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCloneImage Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mCloneImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal mGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal Flags As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal hGraphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mInterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal mGraphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal argb As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByRef mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal stride As Long, ByVal Format As Long, ByRef Scan0 As Any, ByRef BITMAP As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSaveImageToStream Lib "gdiplus" (ByVal Image As Long, ByVal Stream As IUnknown, clsidEncoder As Any, encoderParams As Any) As Long
Private Declare Sub CreateStreamOnHGlobal2 Lib "Ole32.dll" Alias "CreateStreamOnHGlobal" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Any, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
Private Declare Function GetHGlobalFromStream Lib "ole32" (ByVal ppstm As Long, hGlobal As Long) As Long
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByVal pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal HIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function CopyImage Lib "user32.dll" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_CALCRECT As Long = &H400
Private Const DT_VCENTER As Long = &H4
Private Const DT_CENTER As Long = &H1
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_WORDBREAK As Long = &H10
Private Const DT_LEFT As Long = &H0

Private Type ICONINFO
    fIcon           As Long
    xHotspot        As Long
    yHotspot        As Long
    hbmMask         As Long
    hbmColor        As Long
End Type

Private Type BITMAP
  bmType                    As Long
  bmWidth                   As Long
  bmHeight                  As Long
  bmWidthBytes              As Long
  bmPlanes                  As Integer
  bmBitsPixel               As Integer
  bmBits                    As Long
End Type

Private Type BlurParams
    Radius As Single
    ExpandEdge As Long
End Type

Private Type RectF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Private Type COLORMATRIX
    m(0 To 4, 0 To 4)           As Single
End Type

Private Type PicBmp
    Size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion              As Long
    DebugEventCallback          As Long
    SuppressBackgroundThread    As Long
    SuppressExternalCodecs      As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type


Private Const WrapModeTileFlipXY        As Long = &H3
Private Const IDC_HAND                  As Long = 32649
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const SmoothingModeAntiAlias    As Long = 4
Private Const OBJ_BITMAP                As Long = 7
Private Const IMAGE_BITMAP              As Long = 0
Private Const LR_CREATEDIBSECTION       As Long = &H2000
Private Const RotateNoneFlipY           As Long = &H6
Private Const CLR_NONE                  As Long = &HFFFFFFFF

Public Event ItemClick(Index As Long)
Public Event ItemMouseEnter(Index As Long)
Public Event ItemMouseLeave(Index As Long)

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

Private Type tItem
   Caption As String
   TextWidth As Long
   TextHeight As Long
   Image As Long
   ImgArr() As Byte
   Value As Boolean
   Disabled As Boolean
   ToolTipText As String
   RECT As RECT
   DropDown As Boolean
   Tag As String
   Visible As Boolean
End Type

Public Enum eCtrlType
   OptionBar
   CheckBar
   ButtonsBar
End Enum

Public Enum eOrientation
    Horizontal
    Vertical
End Enum

Public Enum eDesignStyle
    PlaneStyle
    ButtonStyle
    TabStyle
End Enum

Dim m_BackColor As OLE_COLOR
Dim m_BackColorOpacity As Integer
Dim m_Border As Boolean
Dim m_BorderColor As OLE_COLOR
Dim m_BorderWidth As Integer
Dim m_MousePointerHands As Boolean
Dim m_Font As StdFont
Dim m_ForeColor As OLE_COLOR
Dim m_Gradient As Boolean
Dim m_BorderRadius As Long
Dim m_Items() As tItem
Dim hCur As Long
Dim c_lhWnd As Long
Dim nScale As Single
Dim bIntercept As Boolean
Dim m_Enter As Boolean
Dim m_PT As POINTAPI
Dim m_Left As Long
Dim m_Top As Long
Dim GdipToken As Long
Dim m_Orientation As eOrientation
Dim m_DesignStyle As eDesignStyle
Dim m_HotItem As Long
Dim m_MinItemWidth As Long
Dim m_IconsAlignTop As Boolean
Dim m_IconsSize As Long
Dim m_DividingLinesVisible As Boolean
Dim m_SelectedBackColor As OLE_COLOR
Dim m_SelectedForeColor As OLE_COLOR
Dim m_TabIndicatorColor As OLE_COLOR
Dim m_TabIndicatorSize As Long
Dim m_Selected As Long
Dim AnimatePos As Long
Dim OldXY As Long
Dim NewXY As Long
Dim InitTime As Double
Dim m_DividingLinesColor As Long
Dim m_Padding As Long
Dim m_CtrlType As eCtrlType
Dim bMouseDown As Boolean
Dim m_UseIconVector As Boolean
Dim Count As Long
Dim m_AutoSize As Boolean
Dim m_FocusedItem  As Long
Dim m_uIPAO As IPAOHookStruct
'*-

Public Sub Draw()
    Dim hPath As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim X As Long, Y As Long
    Dim XX As Long, YY As Long
    Dim lWidth As Long, lHeight As Long
    Dim WW As Long, HH As Long
    Dim BorderWidth As Integer
    Dim i As Long
    Dim TextW As Long, TextH As Long
    Dim R As RECT
    Dim Radius As Single
    Dim BColor As Long
    Dim Padding As Long
    Dim ImgSize As Long
    Dim sh As Long
    Dim MinItemWidth As Long
    Dim hGraphics As Long
    Dim VisibleCount As Long
    
    MinItemWidth = m_MinItemWidth * nScale
    ImgSize = m_IconsSize * nScale
    BorderWidth = m_BorderWidth * nScale
    Radius = m_BorderRadius * nScale
    Padding = m_Padding * nScale
    lWidth = UserControl.ScaleWidth - nScale - BorderWidth
    lHeight = UserControl.ScaleHeight - nScale - BorderWidth
    sh = 2 * nScale + BorderWidth
    
    XX = X
    YY = Y
    WW = lWidth
    HH = lHeight
    
    
    For i = 0 To Count - 1
        If m_Items(i).Visible Then VisibleCount = VisibleCount + 1
    Next
        
    If VisibleCount = 0 Then
        If Not Ambient.UserMode Then
            SetRect R, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
            DrawFocusRect hdc, R
        End If
        Exit Sub
    End If


    For i = 0 To Count - 1
        R.Left = 0
        R.Top = 0
        DrawTextW hdc, StrPtr(m_Items(i).Caption), -1, R, DT_CALCRECT Or DT_LEFT
        TextW = R.Right
        TextH = R.Bottom
        m_Items(i).TextWidth = TextW
        m_Items(i).TextHeight = TextH
    
        
        If m_AutoSize Then
            If m_Orientation = Horizontal Then
               WW = TextW + Padding * 2
               If m_Items(i).DropDown Then WW = WW + m_Items(i).TextHeight / 2.5 + Padding
               If m_Items(i).Image <> 0 Then WW = WW + ImgSize
               
               If WW < MinItemWidth Then WW = MinItemWidth
            Else
               HH = TextH + Padding * 2
            End If
        Else
            If Count > 0 Then
                If m_Orientation = Horizontal Then
                   WW = lWidth / VisibleCount
                   If WW < MinItemWidth Then WW = MinItemWidth
                Else
                   HH = lHeight / VisibleCount
                End If
            End If
        End If
        
        With m_Items(i).RECT
            .Top = YY
            .Left = XX
            .Right = XX + WW
            .Bottom = YY + HH
        End With
    
        If m_Items(i).Visible Then
            If m_Orientation = Horizontal Then
                XX = XX + WW
            Else
                YY = YY + HH
            End If
        End If
    Next
    
    'AUTOSIZE
    If m_AutoSize Then
        If m_Orientation = Horizontal Then
            UserControl.Size (XX + nScale + BorderWidth) * Screen.TwipsPerPixelX, ScaleHeight * Screen.TwipsPerPixelY
        Else
            UserControl.Size ScaleWidth * Screen.TwipsPerPixelX, (YY + nScale + BorderWidth) * Screen.TwipsPerPixelY
        End If
    End If
    
    GdipCreateFromHDC hdc, hGraphics
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    XX = 0
    YY = 0
    WW = UserControl.ScaleWidth ' m_Items(Count - 1).RECT.Right
    HH = UserControl.ScaleHeight ' m_Items(Count - 1).RECT.Bottom
    
    
    
    
'*1
'==================================================================================================
'ESTILO BUTTON
'==================================================================================================
 
If m_DesignStyle = ButtonStyle Then
    
    hPath = CreateRoundPath(XX, YY, WW - nScale, HH - nScale, Radius, Radius, Radius, Radius)
    
    'BRUSH
    hBrush = CreateBrush(m_BackColor, XX, YY, WW, HH)
    'PINTAR FONDO
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
    
    'HOTITEM
    If m_HotItem <> -1 And m_CtrlType <> ButtonsBar Then
        If Not m_Items(m_HotItem).Disabled And m_Items(m_HotItem).Visible Then
            BColor = m_BackColor
            If IsDarkColor(BColor) Then
               BColor = ShiftColor(BColor, vbWhite, IIF(bMouseDown, 100, 190))
            Else
               BColor = ShiftColor(BColor, vbBlack, IIF(bMouseDown, 200, 240))
            End If
        
            With m_Items(m_HotItem).RECT
               
                XX = .Left + sh
                
                WW = .Right - .Left - sh * 2
                If m_Orientation = Horizontal Then
                    HH = .Bottom - .Top - sh * 2
                    YY = .Top + sh
                Else
                    YY = .Top + sh / 2
                    HH = .Bottom - .Top - sh
                End If
            End With
            
            hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
    
            hBrush = CreateBrush(BColor, XX, YY, WW, HH)
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
        End If
    End If
    
    'LINEAS
    If m_DividingLinesVisible Then
        For i = 0 To VisibleCount - 1 - 1
            If m_Items(i + 1).Value = False And m_Items(i).Value = False And m_Items(i).Visible Then
                If m_Orientation = Horizontal Then
                    XX = m_Items(i).RECT.Right
                    YY = m_Items(i).RECT.Bottom * 25 / 100
                    HH = YY + m_Items(i).RECT.Bottom / 2
                    WW = XX
                Else
                    XX = Padding
                    YY = m_Items(i).RECT.Bottom
                    WW = UserControl.ScaleWidth - Padding
                    HH = YY
                End If
                GdipCreatePen1 ConvertColor(m_DividingLinesColor, 100), nScale, UnitPixel, hPen
                GdipDrawLineI hGraphics, hPen, XX, YY, WW, HH
                GdipDeletePen hPen
            End If
        Next
    End If
    
    'SELECION BOTON
    For i = 0 To Count - 1
        Dim b As Boolean
        If m_CtrlType = ButtonsBar Then
           b = m_HotItem = i
           If m_Items(i).Value Then b = True
        Else
           b = m_Items(i).Value
        End If
            
        If b And m_Items(i).Visible = True Then
            With m_Items(i).RECT
                 If m_Orientation = Horizontal Then
                    XX = .Left + AnimatePos + sh
                    YY = .Top + sh
                    HH = .Bottom - .Top - sh * 1.5
                 Else
                    XX = .Left + sh
                    YY = .Top + AnimatePos + sh / 2
                    HH = .Bottom - .Top - sh
                 End If
                 WW = .Right - .Left - sh * 1.5
            End With

            hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
            
            BColor = m_BackColor
            If IsDarkColor(BColor) Then
               BColor = ShiftColor(BColor, vbWhite, 40)
            Else
               BColor = ShiftColor(BColor, vbBlack, 40)
            End If
            
            GdipCreateSolidFill ConvertColor(BColor, 5), hBrush
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
            
            With m_Items(i).RECT
                 If m_Orientation = Horizontal Then
                    XX = .Left + sh + AnimatePos
                    YY = .Top + sh
                    HH = .Bottom - .Top - sh * 1.5 - nScale '* 2
                 Else
                    XX = .Left + sh
                    YY = .Top + AnimatePos + sh / 2
                    HH = .Bottom - .Top - nScale * 2 - sh
                 End If
                 WW = .Right - .Left - sh * 1.5 - nScale '* 2
            End With
            
            hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
            If m_CtrlType = ButtonsBar Then
                If bMouseDown And m_HotItem = i Then
                    BColor = ShiftColor(m_SelectedBackColor, m_BackColor, 127)
                Else
                    BColor = m_SelectedBackColor
                End If
            Else
                BColor = m_SelectedBackColor
            End If
            hBrush = CreateBrush(BColor, XX, YY, WW, HH)
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
                        
            GdipCreatePen1 ConvertColor(m_BorderColor, 50), nScale, UnitPixel, hPen
            GdipDrawPath hGraphics, hPen, hPath
            GdipDeletePen hPen
                        
            GdipDeletePath hPath
        End If
    Next
    
    'IMAGENES
    For i = 0 To Count - 1
        If m_Items(i).Visible = True Then
            If m_Items(i).Image <> 0 Then
                DrawImage hGraphics, i, ImgSize
            End If
            
            If m_Items(i).DropDown Then
                DrawDropDown hGraphics, i
            End If
        End If
    Next
    
    GdipDeleteGraphics hGraphics
    
    'CAPTION
    For i = 0 To Count - 1
        If m_Items(i).Visible = True Then
            If Len(m_Items(i).Caption) Then
                R = m_Items(i).RECT
                R.Top = R.Top - sh / 4
                DrawCaption hdc, i, R, ImgSize
            End If
            
            If i = m_FocusedItem Then
                R = m_Items(i).RECT
                InflateRect R, -4, -4
                DrawFocusRect hdc, R
            End If
        End If
    Next
    
    If m_Border = True Then
        XX = BorderWidth / 2
        YY = BorderWidth / 2
        'WW = m_Items(Count - 1).RECT.Right - BorderWidth '* 2 - nScale / 2
        'HH = m_Items(Count - 1).RECT.Bottom - BorderWidth ' * 2 - nScale / 2
        
        WW = UserControl.ScaleWidth - BorderWidth * 2 '- nScale / 2
        HH = UserControl.ScaleHeight - BorderWidth * 2 ' - nScale / 2
        
        GdipCreateFromHDC hdc, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
        
        GdipCreatePen1 ConvertColor(m_BorderColor, 100), BorderWidth, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
        GdipDeletePath hPath
        GdipDeleteGraphics hGraphics
    End If
     
    Exit Sub
End If

'*2
'==================================================================================================
'ESTILO TAB
'==================================================================================================

If m_DesignStyle = TabStyle Then
    
    hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
    
    'BRUSH
    hBrush = CreateBrush(m_BackColor, XX, YY, WW, HH)
    
    'PINTAR FONDO
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
        
    'HOTITEM
    If m_HotItem <> -1 And m_CtrlType <> ButtonsBar Then
        If Not m_Items(m_HotItem).Disabled And m_Items(m_HotItem).Visible = True Then

            BColor = m_BackColor
            If IsDarkColor(BColor) Then
               BColor = ShiftColor(BColor, vbWhite, IIF(bMouseDown, 100, 190))
            Else
               BColor = ShiftColor(BColor, vbBlack, IIF(bMouseDown, 200, 240))
            End If

            With m_Items(m_HotItem).RECT
                If m_Orientation = Horizontal Then
                    XX = .Left
                    YY = .Top + BorderWidth
                    WW = .Right - .Left
                    HH = .Bottom - .Top - BorderWidth ' * 2
                Else
                    XX = .Left + BorderWidth
                    YY = .Top + BorderWidth
                    WW = .Right - .Left - BorderWidth * 2
                    HH = .Bottom - .Top - BorderWidth * 2
                End If
            End With
            
            hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
    
            'GdipCreateSolidFill ConvertColor(BColor, 100), hBrush
            hBrush = CreateBrush(BColor, XX, YY, WW, HH)
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
        End If
    End If
        
    'LINEAS
    If m_DividingLinesVisible Then
        For i = 0 To VisibleCount - 1 - 1
            If m_Items(i).Visible Then
                If m_Orientation = Horizontal Then
                    XX = m_Items(i).RECT.Right
                    YY = m_Items(i).RECT.Bottom * 25 / 100
                    HH = YY + m_Items(i).RECT.Bottom / 2
                    WW = XX
                Else
                    XX = Padding
                    YY = m_Items(i).RECT.Bottom
                    WW = UserControl.ScaleWidth - Padding
                    HH = YY
                End If
                GdipCreatePen1 ConvertColor(m_DividingLinesColor, 100), nScale, UnitPixel, hPen
                GdipDrawLineI hGraphics, hPen, XX, YY, WW, HH
                GdipDeletePen hPen
            End If
        Next
    End If
    
    'SELECION TAB
    For i = 0 To Count - 1
 
        If m_CtrlType = ButtonsBar Then
           b = m_HotItem = i
           If m_Items(i).Value Then b = True
        Else
           b = m_Items(i).Value
        End If
    
        If b And m_Items(i).Visible Then
            With m_Items(i).RECT
                 If m_Orientation = Horizontal Then
                    XX = .Left + AnimatePos
                    YY = .Top + BorderWidth
                    WW = .Right - .Left
                    HH = .Bottom - .Top - BorderWidth * 2
                 Else
                    XX = .Left + BorderWidth
                    YY = .Top + BorderWidth + AnimatePos
                    WW = .Right - .Left - BorderWidth * 2 + nScale
                    HH = .Bottom - .Top - BorderWidth * 2
                 End If
            End With
            
            hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)

            If m_CtrlType = ButtonsBar Then
                If bMouseDown Then
                    BColor = ShiftColor(m_SelectedBackColor, m_BackColor, 127)
                Else
                    BColor = m_SelectedBackColor
                End If
            Else
                BColor = m_SelectedBackColor
            End If
            
            hBrush = CreateBrush(BColor, XX, YY, WW, HH)

            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath

            With m_Items(i).RECT
                Dim TIS As Long
                TIS = m_TabIndicatorSize * nScale
                 If m_Orientation = Horizontal Then
                    XX = .Left + AnimatePos + Radius / 2
                    YY = .Bottom - TIS '- BorderWidth '* 2
                    WW = .Right - .Left - Radius
                    HH = TIS
                 Else
                    XX = .Left + BorderWidth
                    YY = .Top + AnimatePos + Radius / 2 + nScale
                    WW = TIS
                    HH = .Bottom - .Top - Radius - nScale * 2
                 End If
            End With
            
            hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)

            GdipCreateSolidFill ConvertColor(m_TabIndicatorColor, 100), hBrush
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
            
        End If
    Next
    
    'IMAGENES
    For i = 0 To Count - 1
        If m_Items(i).Visible Then
            If m_Items(i).Image <> 0 Then
                DrawImage hGraphics, i, ImgSize
            End If
            If m_Items(i).DropDown Then
            
                DrawDropDown hGraphics, i
            End If
        End If
    Next
    
    GdipDeleteGraphics hGraphics
    
    'CAPTION
    For i = 0 To Count - 1
        If m_Items(i).Visible Then
            If Len(m_Items(i).Caption) Then
                UserControl.ForeColor = vbBlack
                R = m_Items(i).RECT
                DrawCaption hdc, i, R, ImgSize
            End If
            
            If i = m_FocusedItem Then
                R = m_Items(i).RECT
                InflateRect R, -4, -4
                DrawFocusRect hdc, R
            End If
        End If
    Next
    
    If m_Border = True Then
        XX = BorderWidth / 2
        YY = BorderWidth / 2
        'WW = m_Items(Count - 1).RECT.Right - BorderWidth '* 2
        'HH = m_Items(Count - 1).RECT.Bottom - BorderWidth '* 2
        WW = UserControl.ScaleWidth - BorderWidth * 1.5
        HH = UserControl.ScaleHeight - BorderWidth * 1.5
        GdipCreateFromHDC hdc, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
        
        GdipCreatePen1 ConvertColor(m_BorderColor, 100), BorderWidth, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
        GdipDeletePath hPath
        GdipDeleteGraphics hGraphics
    End If
    
     
     Exit Sub
End If

'*3

'==================================================================================================
'ESTILO PLANO
'==================================================================================================
If m_DesignStyle = PlaneStyle Then
    Dim j As Long
    
    For i = 0 To Count - 1
        If m_Items(i).Visible Then
            If m_Items(i).Value Then
                BColor = ShiftColor(m_BackColor, m_SelectedBackColor, IIF(bMouseDown And m_HotItem = i, 100, AnimatePos))
            Else
                BColor = m_BackColor
              
                If m_HotItem = i And Not m_Items(i).Disabled Then
                    If m_CtrlType = ButtonsBar Then
                        If bMouseDown Then
                            BColor = ShiftColor(m_SelectedBackColor, m_BackColor, 127)
                        Else
                            BColor = m_SelectedBackColor
                        End If
                    Else
                        If IsDarkColor(BColor) Then
                           BColor = ShiftColor(BColor, vbWhite, IIF(bMouseDown, 220, 240))
                        Else
                           BColor = ShiftColor(BColor, vbBlack, IIF(bMouseDown, 220, 240))
                        End If
                    End If
                End If
            End If
            
            
            R = m_Items(i).RECT
            XX = R.Left
            YY = R.Top
            WW = R.Right - R.Left
            HH = R.Bottom - R.Top
            
            
            GdipCreateFromHDC hdc, hGraphics
            GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
            
            'PATH
            If m_Orientation = Horizontal Then
                If j = 0 And j = VisibleCount - 1 Then
                 hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
                ElseIf j = 0 Then
                 hPath = CreateRoundPath(XX, YY, WW, HH, Radius, 0, 0, Radius)
                ElseIf j = VisibleCount - 1 Then
                 hPath = CreateRoundPath(XX, YY, WW, HH, 0, Radius, Radius, 0)
                Else
                 hPath = CreateRoundPath(XX, YY, WW, HH, 0, 0, 0, 0)
                End If
            Else
                If j = 0 Then
                 hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, 0, 0)
                ElseIf j = VisibleCount - 1 Then
                 hPath = CreateRoundPath(XX, YY, WW, HH, 0, 0, Radius, Radius)
                Else
                 hPath = CreateRoundPath(XX, YY, WW, HH, 0, 0, 0, 0)
                End If
            End If
            
            'BRUSH
            hBrush = CreateBrush(BColor, XX, YY, WW, HH)
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
            WW = m_Items(i).RECT.Right - m_Items(i).RECT.Left
            HH = m_Items(i).RECT.Bottom - m_Items(i).RECT.Top

    
            'IMAGEN
            If m_Items(i).Image <> 0 Then
               DrawImage hGraphics, i, ImgSize
            End If
            
            If j > 0 And m_DividingLinesVisible Then
                GdipCreatePen1 ConvertColor(m_DividingLinesColor, 100), nScale, UnitPixel, hPen
                If m_Orientation = Horizontal Then
                    GdipDrawLineI hGraphics, hPen, XX, YY, XX, HH
                Else
                    GdipDrawLineI hGraphics, hPen, XX, YY, XX + WW, YY
                End If
                GdipDeletePen hPen
            End If
            
            If m_Items(i).DropDown Then
                DrawDropDown hGraphics, i
            End If
            GdipDeleteGraphics hGraphics
    
            'TEXTO
            If Len(m_Items(i).Caption) Then
                DrawCaption hdc, i, R, ImgSize
            End If
            
            'FOCO
            If i = m_FocusedItem Then
                R = m_Items(i).RECT
                InflateRect R, -4, -4
                DrawFocusRect hdc, R
            End If
            
            j = j + 1
        End If
        
    Next
    
    If m_Border = True Then
        
        XX = BorderWidth / 2
        YY = BorderWidth / 2
        'WW = m_Items(Count - 1).RECT.Right - BorderWidth + nScale
        WW = UserControl.ScaleWidth - BorderWidth * 2 '+ nScale
        HH = UserControl.ScaleHeight - BorderWidth * 2 '- nScale
        GdipCreateFromHDC hdc, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        hPath = CreateRoundPath(XX, YY, WW, HH, Radius, Radius, Radius, Radius)
        
        GdipCreatePen1 ConvertColor(m_BorderColor, 100), BorderWidth, UnitPixel, hPen
        GdipDrawPath hGraphics, hPen, hPath
        GdipDeletePen hPen
        GdipDeletePath hPath
        GdipDeleteGraphics hGraphics
    End If
    
    Exit Sub
End If
   
End Sub

Private Sub DrawCaption(hdc As Long, i As Long, R As RECT, ImgSize As Long)
    Dim XX As Long, YY As Long, WW As Long, HH As Long
    Dim Padding As Long
    
    Padding = m_Padding * nScale
    WW = m_Items(i).RECT.Right - m_Items(i).RECT.Left
    HH = m_Items(i).RECT.Bottom - m_Items(i).RECT.Top
  
    If m_IconsAlignTop Then
        If m_Items(i).Image <> 0 Then
            YY = R.Top + HH / 2 - (m_Items(i).TextHeight + ImgSize + 4 * nScale) / 2 + ImgSize + 4 * nScale
        Else
            YY = R.Top + HH / 2 - m_Items(i).TextHeight / 2
        End If
        XX = m_Items(i).RECT.Left + WW / 2 - m_Items(i).TextWidth / 2
    Else
        If m_Orientation = Vertical Then
            If m_Items(i).Image <> 0 Then
                XX = m_Items(i).RECT.Left + m_Padding * nScale + ImgSize + m_Padding * nScale / 2
            Else
                XX = m_Items(i).RECT.Left + m_Padding * nScale
            End If
        Else
            If m_Items(i).Image <> 0 Then
                XX = m_Items(i).RECT.Left + WW / 2 - (ImgSize + Padding + m_Items(i).TextWidth) / 2 + ImgSize + Padding
            Else
                XX = m_Items(i).RECT.Left + WW / 2 - m_Items(i).TextWidth / 2
            End If
        End If
        YY = R.Top + HH / 2 - m_Items(i).TextHeight / 2
    End If
    
    If m_Items(i).DropDown And m_Orientation = Horizontal Then
       XX = XX - m_Items(i).TextHeight / 2.5 - m_Padding * nScale
       R.Right = R.Right - m_Items(i).TextHeight / 2.5 * 2 - m_Padding
    End If
    
    If XX > R.Left Then
       R.Left = XX
    End If
    
    
    R.Top = YY
    
    If m_Items(i).Disabled Then
        UserControl.ForeColor = &HCCCCCC
    Else
        If m_CtrlType = ButtonsBar Then
            If m_HotItem = i Or m_Items(i).Value Then
              UserControl.ForeColor = m_SelectedForeColor
            Else
              UserControl.ForeColor = m_ForeColor
            End If
        Else
            If m_Items(i).Value Then
              UserControl.ForeColor = m_SelectedForeColor
            Else
              UserControl.ForeColor = m_ForeColor
            End If
        End If
    End If
    DrawTextW hdc, StrPtr(m_Items(i).Caption), -1, R, DT_LEFT Or DT_WORD_ELLIPSIS
End Sub

Private Sub DrawDropDown(hGraphics As Long, i As Long)
    Dim XX As Single, YY As Single, WW As Single, HH As Single
    Dim FColor As Long
    Dim hPen As Long
    
    If m_Items(i).Disabled Then
        FColor = &HCCCCCC
    Else
        If m_CtrlType = ButtonsBar Then
            If m_HotItem = i Then
              FColor = m_SelectedForeColor
            Else
              FColor = m_ForeColor
            End If
        Else
            If m_Items(i).Value Then
              FColor = m_SelectedForeColor
            Else
              FColor = m_ForeColor
            End If
        End If
    End If

    HH = m_Items(i).TextHeight / 2.5 '* nScale
    WW = HH
    If m_Orientation = Horizontal Then
        XX = m_Items(i).RECT.Right - m_Padding * nScale - WW * 2
    Else
        XX = m_Items(i).RECT.Right - m_Padding * nScale - WW
    End If

    YY = m_Items(i).RECT.Top + (m_Items(i).RECT.Bottom - m_Items(i).RECT.Top) / 2 - HH / 2

    GdipCreatePen1 ConvertColor(FColor, 100), m_Items(i).TextHeight / UserControl.Font.Size, UnitPixel, hPen
    If m_Orientation = Horizontal Then
        GdipDrawLine hGraphics, hPen, XX, YY, XX + WW, YY + HH
        GdipDrawLine hGraphics, hPen, XX + WW, YY + HH, XX + WW * 2, YY
    Else
        GdipDrawLine hGraphics, hPen, XX, YY - HH / 2, XX + WW, YY + HH / 2
        GdipDrawLine hGraphics, hPen, XX + WW, YY + HH / 2, XX, YY + HH + HH / 2
    End If
    
    GdipDeletePen hPen
End Sub

Private Sub DrawImage(hGraphics As Long, i As Long, ImgSize As Long)
    Dim XX As Long, YY As Long, WW As Long, HH As Long
    Dim FColor As Long
    WW = m_Items(i).RECT.Right - m_Items(i).RECT.Left
    HH = m_Items(i).RECT.Bottom - m_Items(i).RECT.Top
    If m_IconsAlignTop Then
        If Len(m_Items(i).Caption) Then
            YY = m_Items(i).RECT.Top + HH / 2 - (m_Items(i).TextHeight + ImgSize + m_Padding * nScale) / 2
        Else
            YY = m_Items(i).RECT.Top + HH / 2 - ImgSize / 2
        End If
        XX = m_Items(i).RECT.Left + WW / 2 - ImgSize / 2
    Else

        If Len(m_Items(i).Caption) Then
            If m_Orientation = Vertical Then
                XX = m_Items(i).RECT.Left + m_Padding * nScale
            Else
                XX = m_Items(i).RECT.Left + WW / 2 - (m_Items(i).TextWidth + ImgSize + m_Padding * nScale) / 2
            End If
        Else
            XX = m_Items(i).RECT.Left + WW / 2 - ImgSize / 2
        End If
        
        YY = m_Items(i).RECT.Top + HH / 2 - ImgSize / 2
        
    End If
    'XX = m_Items(i).Rect.Left + Padding

    If m_Items(i).Disabled Then
        FColor = &HCCCCCC
    Else
        If m_CtrlType = ButtonsBar Then
            If m_HotItem = i Then
              FColor = m_SelectedForeColor
            Else
              FColor = m_ForeColor
            End If
        Else
            If m_Items(i).Value Then
              FColor = m_SelectedForeColor
            Else
              FColor = m_ForeColor
            End If
        End If
    End If

    
    If m_Items(i).DropDown And m_Orientation = Horizontal Then
        XX = XX - m_Padding * nScale - m_Items(i).TextHeight / 2.5 '*
    End If

    If Not m_UseIconVector Then FColor = -1
    
    ImageListDraw m_Items(i).Image, hGraphics, XX, YY, ImgSize, ImgSize, FColor, m_Items(i).Disabled
    
End Sub


Private Function ImageListDraw(ByVal hImage As Long, ByVal hGraphics As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional PictureColor As Long = CLR_NONE, Optional Disabled As Boolean) As Boolean
    Dim tMatrixColor As COLORMATRIX, tMatrixGray As COLORMATRIX
    Dim hAttributes As Long
    Dim ImgW As Long, ImgH As Long
    
    Call GdipCreateImageAttributes(hAttributes)
    GdipGetImageWidth hImage, ImgW
    GdipGetImageHeight hImage, ImgH
    
    With tMatrixColor
        If PictureColor <> CLR_NONE Then
            If (PictureColor And &H80000000) Then PictureColor = GetSysColor(PictureColor And &HFF&)
            Dim R As Byte, G As Byte, b As Byte

            b = ((PictureColor \ &H10000) And &HFF)
            G = ((PictureColor \ &H100) And &HFF)
            R = (PictureColor And &HFF)

            .m(0, 0) = R / 255
            .m(1, 0) = G / 255
            .m(2, 0) = b / 255
            .m(0, 4) = R / 255
            .m(1, 4) = G / 255
            .m(2, 4) = b / 255
        Else
            .m(0, 0) = 1
            .m(1, 1) = 1
            .m(2, 2) = 1
        End If
        .m(3, 3) = 1
        .m(4, 4) = 1

        If Disabled Then
            .m(0, 0) = 0.299
            .m(1, 0) = 0.299
            .m(2, 0) = 0.299
            .m(0, 1) = 0.587
            .m(1, 1) = 0.587
            .m(2, 1) = 0.587
            .m(0, 2) = 0.114
            .m(1, 2) = 0.114
            .m(2, 2) = 0.114
        End If
    End With
    
    'GdipCreateFromHDC hdc, hGraphics
    Call GdipSetInterpolationMode(hGraphics, 7&)  '//InterpolationModeHighQualityBicubic
    Call GdipSetPixelOffsetMode(hGraphics, 4&)
    GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0
    ImageListDraw = GdipDrawImageRectRectI(hGraphics, hImage, Left, Top, Width, Height, 0, 0, ImgW, ImgH, UnitPixel, hAttributes) = 0&
    'Call GdipDeleteGraphics(hGraphics)
    Call GdipDisposeImageAttributes(hAttributes)
End Function

Private Function CreateBrush(ByVal Color As Long, Left, Top, Width, Height) As Long
    Dim RECTL As RECTL
    Dim Color1 As Long, Color2 As Long
    If m_Gradient Then
        SetRect RECTL, Left, Top, Width, Height
        If Not IsDarkColor(m_BackColor) Or IsDarkColor(Color) Then
          Color1 = ConvertColor(ShiftColor(Color, vbWhite, 100), m_BackColorOpacity)
          Color2 = ConvertColor(Color, m_BackColorOpacity)
        Else
          Color1 = ConvertColor(Color, m_BackColorOpacity)
          Color2 = ConvertColor(ShiftColor(Color, vbBlack, 230), m_BackColorOpacity)
        End If

        GdipCreateLineBrushFromRectWithAngleI RECTL, Color1, Color2, 90, 0, WrapModeTileFlipXY, CreateBrush
    Else
        GdipCreateSolidFill ConvertColor(Color, m_BackColorOpacity), CreateBrush
    End If
End Function

Private Function CreateRoundPath(ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, ByVal LT As Single, ByVal RT As Single, ByVal RB As Single, ByVal LB As Single) As Long
    Dim hPath As Long
    
    
    If LT > Width Then LT = Width
    If LT > Height Then LT = Height
    If RT > Width Then RT = Width
    If RT > Height Then RT = Height
    If RB > Width Then RB = Width
    If RB > Height Then RB = Height
    If LB > Width Then LB = Width
    If LB > Height Then LB = Height

    If GdipCreatePath(&H0, hPath) = 0& Then
    
            If LT = 0 Then
              GdipAddPathLineI hPath, Left, Top, Left + Width - RT, Top
            Else
              GdipAddPathArcI hPath, Left, Top, LT, LT, 180, 90
            End If
            
            If RT = 0 Then
              GdipAddPathLineI hPath, Left + Width, Top, Left + Width, Top + Height - RB
            Else
              GdipAddPathArcI hPath, Left + Width - RT, Top, RT, RT, 270, 90
            End If
            
            If RB = 0 Then
              GdipAddPathLineI hPath, Left + Width - RB, Top + Height, Left + LB, Top + Height
            Else
              GdipAddPathArcI hPath, Left + Width - RB, Top + Height - RB, RB, RB, 0, 90
            End If
            
            If LB = 0 Then
              GdipAddPathLineI hPath, Left + Width - RB, Top + Height, Left + LB, Top + Height
            Else
              GdipAddPathArcI hPath, Left, Top + Height - LB, LB, LB, 90, 90
            End If

            GdipClosePathFigure hPath
        'End If
        CreateRoundPath = hPath
    End If
    
End Function

Public Function GetWindowsDPI() As Double
    Dim hdc As Long, lPx  As Double
    hdc = GetDC(0)
    lPx = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc

    If (lPx = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = lPx / 96#
    End If
End Function


Public Sub MoveItem(ByVal Index As Long, ByVal NewPos As Long)
    Dim Item As tItem
    Dim i As Long
    
    Item = m_Items(Index)
    
    
    If Index < NewPos Then
        For i = Index To NewPos - 1
            m_Items(i) = m_Items(i + 1)
        Next
        m_Items(NewPos) = Item
    Else
        For i = Index To NewPos + 1 Step -1
            m_Items(i) = m_Items(i - 1)
        Next
        m_Items(NewPos) = Item
    End If
    Refresh
End Sub


Public Sub AddItem(Optional ByVal Caption As String, _
                    Optional Image As Variant, _
                    Optional Value As Boolean, _
                    Optional ToolTipText As String, _
                    Optional Tag As String, _
                    Optional Disabled As Boolean, _
                    Optional DropDown As Boolean)
                    
    ReDim Preserve m_Items(Count)
    With m_Items(Count)
        .Caption = Caption
        If Not IsEmpty(Image) Then

            ItemLoadImage Count, Image
        End If
        .Value = Value
        .ToolTipText = ToolTipText
        .Tag = Tag
        .Disabled = Disabled
        .DropDown = DropDown
        .Visible = True
    End With
    PropertyChanged "Caption_" & Count
    Count = Count + 1
    
    Refresh
End Sub

Public Function ItemsCount() As Long
    ItemsCount = Count
End Function

Public Function Clear()
    Erase m_Items
    Count = 0
    Refresh
End Function

Public Function RemoveItem(Index As Long)
    Dim i As Long
    
    If m_Items(Index).Image <> 0 Then
        GdipDisposeImage m_Items(Index).Image
    End If
    
    For i = Index To Count - 2
        m_Items(i) = m_Items(i + 1)
    Next
    Count = Count - 1
    If Count > 0 Then
        ReDim Preserve m_Items(Count - 1)
    Else
        ReDim m_Items(0)
    End If
    Refresh
End Function

Public Property Let ItemCaption(ByVal Index As Long, ByVal Caption As String)
    m_Items(Index).Caption = Caption
    Refresh
End Property

Public Property Get ItemCaption(ByVal Index As Long) As String
    ItemCaption = m_Items(Index).Caption
End Property

Public Property Get ItemDropDown(ByVal Index As Long) As Boolean
    ItemDropDown = m_Items(Index).DropDown
End Property

Public Property Let ItemDropDown(ByVal Index As Long, ByVal Value As Boolean)
    m_Items(Index).DropDown = Value
    Refresh
End Property

Public Property Let ItemTag(ByVal Index As Long, ByVal Tag As String)
    m_Items(Index).Tag = Tag
    Refresh
End Property

Public Property Get ItemTag(ByVal Index As Long) As String
    ItemTag = m_Items(Index).Tag
End Property

Public Property Let ItemToolTipText(ByVal Index As Long, ByVal ToolTipText As String)
    m_Items(Index).ToolTipText = ToolTipText
    Refresh
End Property

Public Property Get ItemToolTipText(ByVal Index As Long) As String
    ItemToolTipText = m_Items(Index).ToolTipText
End Property

Public Property Let ItemValue(ByVal Index As Long, ByVal Value As Boolean)
    If m_CtrlType = OptionBar Then
       Dim i As Long
       For i = 0 To Count - 1
          m_Items(i).Value = False
       Next
    End If
    m_Items(Index).Value = Value
    If Value Then m_Selected = Index
    Refresh
End Property

Public Property Get ItemValue(ByVal Index As Long) As Boolean
    ItemValue = m_Items(Index).Value
End Property

Public Property Get SelItem() As Long
   SelItem = m_Selected
End Property

Public Property Let SelItem(ByVal Index As Long)
    'If m_CtrlType = OptionBar Then
       Dim i As Long
       For i = 0 To Count - 1
          m_Items(i).Value = False
       Next
    'End If
    If Index > -1 Then
        m_Items(Index).Value = True
    End If
   m_Selected = Index
   PropertyChanged "Selected"
   Refresh
End Property

Public Property Let ItemDisabled(ByVal Index As Long, ByVal Disabled As Boolean)
    m_Items(Index).Disabled = Disabled
    Refresh
End Property

Public Property Get ItemDisabled(ByVal Index As Long) As Boolean
    ItemDisabled = m_Items(Index).Disabled
End Property

Public Property Let ItemVisible(ByVal Index As Long, ByVal Visible As Boolean)
    m_Items(Index).Visible = Visible
    Refresh
End Property

Public Property Get ItemVisible(ByVal Index As Long) As Boolean
    ItemVisible = m_Items(Index).Visible
End Property

Public Property Let FocusedItem(ByVal Value As Long)
    m_FocusedItem = Value
    Refresh
End Property

Public Property Get FocusedItem() As Long
    FocusedItem = m_FocusedItem
End Property



Public Property Let ItemImage(ByVal Index As Long, ByVal Image As Variant)
    If m_Items(Index).Image <> 0 Then
        GdipDisposeImage m_Items(Index).Image
        m_Items(Index).Image = 0
    End If
  
    'If Image <> Null Then
        ItemLoadImage Index, Image
    'End If
    Refresh
End Property

Public Property Get ItemImage(ByVal Index As Long) As Variant
    ItemImage = m_Items(Index).Image
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    Refresh
End Property

Public Property Get BackColorOpacity() As Integer
    BackColorOpacity = m_BackColorOpacity
End Property

Public Property Let BackColorOpacity(ByVal New_BackColorOpacity As Integer)
    m_BackColorOpacity = New_BackColorOpacity
    PropertyChanged "BackColorOpacity"
    Refresh
End Property

Public Property Get Border() As Boolean
    Border = m_Border
End Property

Public Property Let Border(ByVal New_Border As Boolean)
    m_Border = New_Border
    PropertyChanged "Border"
    
    Refresh
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    Refresh
End Property

Public Property Get BorderWidth() As Integer
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    m_BorderWidth = New_BorderWidth
    PropertyChanged "BorderWidth"
    Refresh
End Property


Public Property Get BorderRadius() As Integer
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal New_Value As Integer)
    m_BorderRadius = New_Value
    PropertyChanged "BorderRadius"
    Refresh
End Property


Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(New_Font As StdFont)
    With m_Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    Set UserControl.Font = m_Font
    PropertyChanged "Font"

    Refresh
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Refresh
End Property

Public Property Get SelectedBackColor() As OLE_COLOR
    SelectedBackColor = m_SelectedBackColor
End Property

Public Property Let SelectedBackColor(ByVal New_Color As OLE_COLOR)
    m_SelectedBackColor = New_Color
    PropertyChanged "SelectedBackColor"
    Refresh
End Property

Public Property Get SelectedForeColor() As OLE_COLOR
    SelectedForeColor = m_SelectedForeColor
End Property

Public Property Let SelectedForeColor(ByVal New_Color As OLE_COLOR)
    m_SelectedForeColor = New_Color
    PropertyChanged "SelectedForeColor"
    Refresh
End Property

Public Property Get TabIndicatorColor() As OLE_COLOR
    TabIndicatorColor = m_TabIndicatorColor
End Property

Public Property Let TabIndicatorColor(ByVal New_Color As OLE_COLOR)
    m_TabIndicatorColor = New_Color
    PropertyChanged "TabIndicatorColor"
    Refresh
End Property

Public Property Get TabIndicatorSize() As Long
    TabIndicatorSize = m_TabIndicatorSize
End Property

Public Property Let TabIndicatorSize(ByVal New_Value As Long)
    m_TabIndicatorSize = New_Value
    PropertyChanged "TabIndicatorSize"
    Refresh
End Property


Public Property Get DividingLinesColor() As OLE_COLOR
    DividingLinesColor = m_DividingLinesColor
End Property

Public Property Let DividingLinesColor(ByVal New_Color As OLE_COLOR)
    m_DividingLinesColor = New_Color
    PropertyChanged "DividingLinesColor"
    Refresh
End Property


Public Property Get Orientation() As eOrientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal New_Value As eOrientation)
    m_Orientation = New_Value
    PropertyChanged "Orientation"
    Refresh
End Property

Public Property Get DesignStyle() As eDesignStyle
    DesignStyle = m_DesignStyle
End Property

Public Property Let DesignStyle(ByVal New_Value As eDesignStyle)
    m_DesignStyle = New_Value
    PropertyChanged "DesignStyle"
    Refresh
End Property

Public Property Get MinItemWidth() As Long
    MinItemWidth = m_MinItemWidth
End Property

Public Property Let MinItemWidth(ByVal New_Value As Long)
    m_MinItemWidth = New_Value
    PropertyChanged "MinItemWidth"
    Refresh
End Property

Public Property Get IconsAlignTop() As Boolean
    IconsAlignTop = m_IconsAlignTop
End Property

Public Property Let IconsAlignTop(ByVal New_Value As Boolean)
    m_IconsAlignTop = New_Value
    PropertyChanged "IconsAlignTop"
    Refresh
End Property

Public Property Get IconsSize() As Long
    IconsSize = m_IconsSize
End Property

Public Property Let IconsSize(ByVal New_Value As Long)
    m_IconsSize = New_Value
    PropertyChanged "IconsSize"
    Refresh
End Property

Public Property Get Padding() As Long
    Padding = m_Padding
End Property

Public Property Let Padding(ByVal New_Value As Long)
    m_Padding = New_Value
    PropertyChanged "Padding"
    Refresh
End Property


Public Property Get DividingLinesVisible() As Boolean
    DividingLinesVisible = m_DividingLinesVisible
End Property

Public Property Let DividingLinesVisible(ByVal New_Value As Boolean)
    m_DividingLinesVisible = New_Value
    PropertyChanged "DividingLinesVisible"
    Refresh
End Property

Public Property Get Gradient() As Boolean
    Gradient = m_Gradient
End Property

Public Property Let Gradient(ByVal New_Gradient As Boolean)
    m_Gradient = New_Gradient
    PropertyChanged "Gradient"
    Refresh
End Property

Public Property Get ControlType() As eCtrlType
    ControlType = m_CtrlType
End Property

Public Property Let ControlType(ByVal New_Value As eCtrlType)
    m_CtrlType = New_Value
    PropertyChanged "CtrlType"
    Refresh
End Property

Public Property Get UseIconVector() As Boolean
    UseIconVector = m_UseIconVector
End Property

Public Property Let UseIconVector(ByVal New_Value As Boolean)
    m_UseIconVector = New_Value
    PropertyChanged "UseIconVector"
    Refresh
End Property


Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_Value As Boolean)
    m_AutoSize = New_Value
    PropertyChanged "AutoSize"
    Refresh
End Property

Public Property Get Object() As Object
    Set Object = UserControl.Extender
End Property

'Public Property Get WordWrap() As Boolean
'    WordWrap = m_WordWrap
'End Property
'
'Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
'    m_WordWrap = New_WordWrap
'    PropertyChanged "WordWrap"
'    Refresh
'End Property

Public Function Move(Left As Single, Optional Top, Optional Width, Optional Height)
    If Not IsMissing(Height) Then
        Extender.Move Left, Top, Width, Height
    ElseIf Not IsMissing(Width) Then
        Extender.Move Left, Top, Width
    ElseIf Not IsMissing(Top) Then
        Extender.Move Left, Top
    Else
        Extender.Move Left
    End If
End Function

Public Property Get Visible() As Boolean
    Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal New_Value As Boolean)
    Extender.Visible = New_Value
End Property

Public Property Get Tag() As Variant
    Tag = Extender.Tag
End Property

Public Property Let Tag(ByVal New_Value As Variant)
    Extender.Tag = New_Value
End Property

Private Function SaveImage(i) As Byte()
    Dim hBmp As Long, Grph As Long
    Dim ImgW As Long, ImgH As Long

   
    GdipGetImageWidth m_Items(i).Image, ImgW
    GdipGetImageHeight m_Items(i).Image, ImgH
    

    If GdipCreateBitmapFromScan0(m_IconsSize, m_IconsSize, 0&, &HE200B, ByVal 0&, hBmp) = 0 Then
        If GdipGetImageGraphicsContext(hBmp, Grph) = 0 Then
        
            Call GdipSetInterpolationMode(Grph, 7&)  '//InterpolationModeHighQualityBicubic
            Call GdipSetPixelOffsetMode(Grph, 4&)
            
            Call GdipDrawImageRectRectI(Grph, m_Items(i).Image, 0, 0, m_IconsSize, m_IconsSize, 0, 0, ImgW, ImgH, &H2)
            GdipDeleteGraphics Grph
            
            SaveImage = BitmapToArray(hBmp)
        End If
    End If
End Function


Private Sub Timer1_Timer()
  Dim p As Single, tt As Single, d As Long, Duration As Single
  tt = Timer - InitTime
  
  If m_DesignStyle = PlaneStyle Then
    d = 255
    Duration = 0.2
  Else
    d = OldXY - NewXY
    Duration = 0.1
  End If
  
  p = tt * 100 / Duration
  AnimatePos = d - (d * p / 100)
  
  UserControl.Refresh
 
  If InitTime + Duration < Timer Then
    Timer1.Interval = 0
    AnimatePos = 0
  End If
End Sub


Private Sub UserControl_ExitFocus()
    m_FocusedItem = -1
    Refresh
End Sub

Private Sub UserControl_GotFocus()
    Call mIOleInPlaceActivate.SetIPAO(m_uIPAO, Me)
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    On Error Resume Next

    If UserControl.Enabled Then
        HitResult = vbHitResultHit
        Dim pt As POINTAPI
        Dim lHwnd As Long
        GetCursorPos pt
        lHwnd = WindowFromPoint(pt.X, pt.Y)
        
        If m_Enter = False Then
        
            ScreenToClient c_lhWnd, pt
            m_PT.X = pt.X - X
            m_PT.Y = pt.Y - Y
        
            m_Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
            m_Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)
        
            m_Enter = True
            tmrMOUSEOVER.Interval = 1
        
            RaiseEvent MouseEnter
        End If
        
        bIntercept = True
    ElseIf Not Ambient.UserMode Then
        HitResult = vbHitResultHit
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI
    m_Selected = -1
    m_HotItem = -1
    m_FocusedItem = -1
    Count = 0

    Call mIOleInPlaceActivate.InitIPAO(m_uIPAO, Me)
End Sub


Private Sub UserControl_InitProperties()

    m_BackColor = Ambient.BackColor
    m_BackColorOpacity = 100
    m_BorderColor = vbActiveBorder
    m_DividingLinesColor = vbActiveBorder
    Set m_Font = UserControl.Ambient.Font
    m_ForeColor = vbButtonText
    m_SelectedForeColor = vbHighlightText
    m_SelectedBackColor = vbHighlight
    m_BorderWidth = 1
    m_Padding = 10
    m_IconsSize = 16
    m_TabIndicatorSize = 4
    c_lhWnd = UserControl.ContainerHwnd
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, m As Long
    RaiseEvent KeyDown(KeyCode, Shift)
    
    If KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then
'        If m_FocusedItem < Count - 1 Then
'            m_FocusedItem = m_FocusedItem + 1
'            Refresh
'        End If
        m = m_FocusedItem
        Do While m_FocusedItem < Count

          m_FocusedItem = m_FocusedItem + 1
          
           If m_FocusedItem > Count - 1 Then
            m_FocusedItem = m
            Exit Do
          End If
          
          If m_Items(m_FocusedItem).Visible Then
             Exit Do
          End If
        Loop
        Debug.Print m_FocusedItem
        Refresh
    ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then
        m = m_FocusedItem
        Do While m_FocusedItem > -1
            m_FocusedItem = m_FocusedItem - 1
             If m_FocusedItem < 0 Then
              m_FocusedItem = m
              Exit Do
            End If
            
            If m_Items(m_FocusedItem).Visible Then
               Exit Do
            End If
        Loop

        Refresh
    ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then
   
        If m_FocusedItem > -1 Then
        
            If m_CtrlType = ButtonsBar Then
            
            ElseIf m_CtrlType = OptionBar Then
                Call UpdateActivePos
                
                For i = 0 To Count - 1
                    If m_FocusedItem = i Then
                        If Not m_Items(i).Value And Not m_Items(i).Disabled And m_Items(i).Visible Then
                            
                            m_Items(i).Value = True
                            If m_Orientation = Horizontal Then
                               NewXY = m_Items(i).RECT.Left
                            Else
                               NewXY = m_Items(i).RECT.Top
                            End If
                            m_Selected = i
                            InitTime = Timer
                            Timer1.Interval = 1
                      
                        End If
                    Else
                       m_Items(i).Value = False
                    End If
                Next
            ElseIf m_CtrlType = CheckBar Then
                m_Items(m_FocusedItem).Value = Not m_Items(m_FocusedItem).Value
                Refresh
            End If
            RaiseEvent ItemClick(m_FocusedItem)
        End If
    End If
    
   
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub UserControl_Paint()
    Call Draw
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    c_lhWnd = UserControl.ContainerHwnd
    
    With PropBag
        m_BackColor = .ReadProperty("BackColor", Ambient.BackColor)
        m_SelectedBackColor = .ReadProperty("SelectedBackColor", vbHighlight)
        m_BackColorOpacity = .ReadProperty("BackColorOpacity", 100)
        m_Border = .ReadProperty("Border", False)
        m_BorderColor = .ReadProperty("BorderColor", vbActiveBorder)
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_BorderRadius = .ReadProperty("BorderRadius", 10)
        m_DividingLinesVisible = .ReadProperty("DividingLinesVisible", True)
        m_ForeColor = .ReadProperty("ForeColor", vbButtonText)
        m_SelectedForeColor = .ReadProperty("SelectedForeColor", vbHighlightText)
        m_TabIndicatorColor = .ReadProperty("TabIndicatorColor", vbHighlight)
        m_TabIndicatorSize = .ReadProperty("TabIndicatorSize", 4)
        m_DividingLinesColor = .ReadProperty("DividingLinesColor", vbActiveBorder)
        m_IconsSize = .ReadProperty("IconsSize", 16)
        m_Orientation = .ReadProperty("Orientation", Horizontal)
        m_DesignStyle = .ReadProperty("DesignStyle", PlaneStyle)
        m_MinItemWidth = .ReadProperty("MinItemWidth", 32)
        m_Padding = .ReadProperty("Padding", 10)
        m_IconsAlignTop = .ReadProperty("IconsAlignTop", False)
        m_UseIconVector = .ReadProperty("UseIconVector", False)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set m_Font = .ReadProperty("Font", UserControl.Ambient.Font)
        m_Gradient = .ReadProperty("Gradient", False)
        m_CtrlType = .ReadProperty("CtrlType", OptionBar)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbArrow)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_MousePointerHands = .ReadProperty("MousePointerHands", False)
        m_Selected = .ReadProperty("Selected", -1)
        UserControl.OLEDropMode = .ReadProperty("OLEDropMode", 0&)
        m_AutoSize = .ReadProperty("AutoSize", False)
        
        Set UserControl.Font = m_Font
        
        If m_MousePointerHands Then
            If Ambient.UserMode Then
                UserControl.MousePointer = vbCustom
                UserControl.MouseIcon = GetSystemHandCursor
            End If
        End If
        
        Dim i As Long, SaveCount As Long
        
        SaveCount = .ReadProperty("Count", 0)
   
        For i = 0 To SaveCount - 1
            Dim b()  As Byte
            Me.AddItem VarToStr(.ReadProperty("Caption_" & i, vbNullString))
            If .ReadProperty("HasImg_" & i, False) Then
                b = .ReadProperty("image_" & i, 0)
                Me.ItemImage(i) = b
            End If
            
            m_Items(i).Tag = .ReadProperty("Tag_" & i, Empty)
            m_Items(i).ToolTipText = .ReadProperty("ToolTipText_" & i, vbNullString)
            m_Items(i).Disabled = .ReadProperty("Disabled_" & i, False)
            m_Items(i).DropDown = .ReadProperty("DropDown_" & i, False)
            m_Items(i).Value = .ReadProperty("Value_" & i, False)
        Next
    End With
        
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", m_BackColor, Ambient.BackColor
        .WriteProperty "BackColorOpacity", m_BackColorOpacity, 100
        .WriteProperty "Border", m_Border, False
        .WriteProperty "BorderColor", m_BorderColor, vbActiveBorder
        .WriteProperty "BorderWidth", m_BorderWidth, 1
        .WriteProperty "BorderRadius", m_BorderRadius, 10
        .WriteProperty "DividingLinesVisible", m_DividingLinesVisible, True
        .WriteProperty "SelectedBackColor", m_SelectedBackColor, vbHighlight
        .WriteProperty "SelectedForeColor", m_SelectedForeColor, vbHighlightText
        .WriteProperty "TabIndicatorColor", m_TabIndicatorColor, vbHighlight
        .WriteProperty "TabIndicatorSize", m_TabIndicatorSize, 4
        .WriteProperty "DividingLinesColor", m_DividingLinesColor, vbActiveBorder
        .WriteProperty "Orientation", m_Orientation, Horizontal
        .WriteProperty "DesignStyle", m_DesignStyle, PlaneStyle
        .WriteProperty "MinItemWidth", m_MinItemWidth, 32
        .WriteProperty "Padding", m_Padding, 10
        .WriteProperty "IconsAlignTop", m_IconsAlignTop, False
        .WriteProperty "UseIconVector", m_UseIconVector, False
        .WriteProperty "IconsSize", m_IconsSize, 16
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Font", m_Font, UserControl.Ambient.Font
        .WriteProperty "ForeColor", m_ForeColor, vbButtonText
        .WriteProperty "Gradient", m_Gradient, False
        .WriteProperty "CtrlType", m_CtrlType, OptionBar
        .WriteProperty "OLEDropMode", UserControl.OLEDropMode, 0&
        .WriteProperty "Selected", m_Selected, -1
        .WriteProperty "AutoSize", m_AutoSize, False
        .WriteProperty "MousePointer", UserControl.MousePointer, vbArrow
        .WriteProperty "Selected", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointerHands", m_MousePointerHands, False

        '--------------------------
        .WriteProperty "Count", Count, 0
        Dim i As Long
        For i = 0 To Count - 1
            Dim b()  As Byte
            If m_Items(i).Image <> 0 Then
                'B = SaveImage(I)  'comprime el tamaño, pero con el dpi mas alto no favorece la calidad.
                b = BitmapToArray(m_Items(i).Image)
                .WriteProperty "image_" & i, b, 0
            End If
            .WriteProperty "HasImg_" & i, m_Items(i).Image <> 0, False
            .WriteProperty "Caption_" & i, StrToVar(m_Items(i).Caption), vbNullString
            .WriteProperty "Tag_" & i, m_Items(i).Tag, Empty
            .WriteProperty "ToolTipText_" & i, m_Items(i).ToolTipText, vbNullString
            .WriteProperty "Disabled_" & i, m_Items(i).Disabled, False
            .WriteProperty "DropDown_" & i, m_Items(i).DropDown, False
            .WriteProperty "Value_" & i, m_Items(i).Value, False
        Next
        
    End With

End Sub

Private Sub UserControl_Terminate()
    Dim i As Long
    For i = 0 To Count - 1
        If m_Items(i).Image <> 0 Then
            GdipDisposeImage m_Items(i).Image
        End If
    Next
    
    If hCur Then DestroyCursor hCur

    Call GdiplusShutdown(GdipToken)
    Call mIOleInPlaceActivate.TerminateIPAO(m_uIPAO)
End Sub

Private Function ConvertColor(ByVal Color As Long, ByVal Opacity As Long) As Long
    Dim BGRA(0 To 3) As Byte
    OleTranslateColor Color, 0, VarPtr(Color)
  
    BGRA(3) = CByte((Abs(Opacity) / 100) * 255)
    BGRA(0) = ((Color \ &H10000) And &HFF)
    BGRA(1) = ((Color \ &H100) And &HFF)
    BGRA(2) = (Color And &HFF)
    CopyMemory ConvertColor, BGRA(0), 4&
End Function

Private Function StrToVar(ByVal Text As String) As Variant
    If Text = vbNullString Then
        StrToVar = Empty
    Else
        Dim b() As Byte
        b() = Text
        StrToVar = b()
    End If
End Function

Private Function VarToStr(ByVal Bytes As Variant) As String
    If IsEmpty(Bytes) Then
        VarToStr = vbNullString
    Else
        Dim b() As Byte
        b() = Bytes
        VarToStr = b()
    End If
End Function


Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function ItemLoadImage(Index As Long, SrcImg As Variant)
    Dim DataArr() As Byte
    
    If m_Items(Index).Image <> 0 Then GdipDisposeImage m_Items(Index).Image: m_Items(Index).Image = 0

    
    Select Case VarType(SrcImg)
        Case vbString
            If PathIsURL(SrcImg) Then
            
                If Left$(LCase(SrcImg), 5) = "data:" Then
                    Base64Decode Split(SrcImg, ",")(1), DataArr
                    m_Items(Index).Image = LoadImageFromArray(DataArr)
                Else
                    Dim oXMLHTTP As Object
                    Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")

                    oXMLHTTP.Open "GET", SrcImg, True
                    oXMLHTTP.send
                    While oXMLHTTP.readyState <> 4
                        DoEvents
                    Wend
                    If oXMLHTTP.Status = 200 Then
                        DataArr() = oXMLHTTP.responseBody
                        m_Items(Index).ImgArr = DataArr
                        m_Items(Index).Image = LoadImageFromArray(m_Items(Index).ImgArr)
                    End If
                    'LoadPicture = PictureFromURL(SrcImg, UseCache, DrawProgress)
                    'Exit Function
                End If
            Else
                Call GdipLoadImageFromFile(ByVal StrPtr(SrcImg), m_Items(Index).Image)
            End If
            
        Case vbLong
            Dim hBmp As Long
            Dim IIF As ICONINFO
            Dim tBmp As BITMAP
            
            If GetObjectType(SrcImg) = OBJ_BITMAP Then
                If GetObject(SrcImg, Len(tBmp), tBmp) Then
                    
                    If tBmp.bmBitsPixel = 32 And tBmp.bmBits > 0 Then
                        GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, PixelFormat32bppPARGB, tBmp.bmBits, m_Items(Index).Image
                        GdipImageRotateFlip m_Items(Index).Image, RotateNoneFlipY
                    Else
                        Call GdipCreateBitmapFromHBITMAP(SrcImg, 0, m_Items(Index).Image)
                    End If
                    
                End If
            Else
        
                If TypeName(SrcImg) = "Long" Then 'StdPicture.handle
                    If GetIconInfo(SrcImg, IIF) Then
                    
                        hBmp = CopyImage(IIF.hbmColor, IMAGE_BITMAP, 0, 0, LR_CREATEDIBSECTION Or &H2)
                        Call GetObject(hBmp, Len(tBmp), tBmp)
                        GdipCreateBitmapFromScan0 tBmp.bmWidth, tBmp.bmHeight, tBmp.bmWidthBytes, PixelFormat32bppPARGB, tBmp.bmBits, m_Items(Index).Image
                        GdipImageRotateFlip m_Items(Index).Image, RotateNoneFlipY
                        DeleteObject hBmp
                        DeleteObject IIF.hbmColor
                        DeleteObject IIF.hbmMask
                    Else
                        On Error Resume Next
                        If IsBadCodePtr(SrcImg) = 0 Then
                            GdipCloneImage SrcImg, m_Items(Index).Image ' OBJECT GDI PLUS BITMAP
                        End If
                        On Error GoTo 0
                    End If
                Else 'StdPicture
                    GdipCreateBitmapFromHICON SrcImg, m_Items(Index).Image
                End If
            End If
            
        Case vbDataObject
            Call GdipLoadImageFromStream(SrcImg, m_Items(Index).Image)
            
        Case (vbArray Or vbByte)
            m_Items(Index).ImgArr = SrcImg
            m_Items(Index).Image = LoadImageFromArray(m_Items(Index).ImgArr)
            'DataArr() = SrcImg
            'm_Items(Index).Image = LoadImageFromArray(DataArr)
    End Select
    
End Function

Private Function Base64Decode(ByVal sIn As String, ByRef bvOut() As Byte) As Boolean 'By Cocus
    Dim lLenOut As Long
    Const CRYPT_STRING_BASE64 = &H1
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, 0, VarPtr(lLenOut), 0, 0)
    If lLenOut = 0 Then Exit Function
    ReDim bvOut(lLenOut - 1)
    Call CryptStringToBinaryA(sIn, Len(sIn), CRYPT_STRING_BASE64, VarPtr(bvOut(0)), VarPtr(lLenOut), 0, 0)
    Base64Decode = True
End Function

Private Function LoadImageFromArray(ByRef bvData() As Byte) As Long
    On Local Error GoTo LoadImageFromArray_Error
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal2(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        GdipLoadImageFromStream IStream, LoadImageFromArray
    End If
    Set IStream = Nothing

    
LoadImageFromArray_Error:

End Function

Public Sub tmrMOUSEOVER_Timer()
    Dim pt As POINTAPI
    Dim Left As Long, Top As Long
    Dim RECT As RECT
  
    GetCursorPos pt
    ScreenToClient c_lhWnd, pt
    
    Left = ScaleX(Extender.Left, vbContainerSize, UserControl.ScaleMode)
    Top = ScaleY(Extender.Top, vbContainerSize, UserControl.ScaleMode)

    With RECT
        .Left = m_PT.X - (m_Left - Left)
        .Top = m_PT.Y - (m_Top - Top)
        .Right = .Left + UserControl.ScaleWidth
        .Bottom = .Top + UserControl.ScaleHeight
    End With
    
   
    
    If PtInRect(RECT, pt.X, pt.Y) = 0 Then
        'WriteValue &H10, 0
        If m_HotItem <> -1 Then RaiseEvent ItemMouseLeave(m_HotItem)
        m_HotItem = -1
        UserControl.Refresh
        m_Enter = False
        tmrMOUSEOVER.Interval = 0
        RaiseEvent MouseLeave
    End If
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    For i = 0 To Count - 1
        If m_Items(i).Visible And PtInRect(m_Items(i).RECT, X, Y) Then
           If m_CtrlType = ButtonsBar Then
                m_HotItem = i
                bMouseDown = True
                UserControl.Refresh
           End If
           Exit For
        End If
    Next
    If hCur Then SetCursor hCur
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UpdateActivePos()
    Dim i As Long
    OldXY = -1
    For i = 0 To Count - 1
      If m_Items(i).Value = True Then
           If m_Orientation = Horizontal Then
              OldXY = m_Items(i).RECT.Left
           Else
              OldXY = m_Items(i).RECT.Top
           End If
          Exit For
      End If
    Next
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    bMouseDown = False
    Dim Index As Long
    Index = -1
    
    For i = 0 To Count - 1
        If PtInRect(m_Items(i).RECT, X, Y) Then
            If m_Items(i).Disabled Then Exit Sub
        End If
    Next
    
    If m_CtrlType = ButtonsBar Then
        For i = 0 To Count - 1
            If m_Items(i).Visible And PtInRect(m_Items(i).RECT, X, Y) Then
               If Not m_Items(i).Disabled Then
                 Index = i
                 Refresh
                 Exit For
               End If
            End If
        Next
        
    ElseIf m_CtrlType = OptionBar Then
    
        Call UpdateActivePos
        For i = 0 To Count - 1
            If m_Items(i).Visible And PtInRect(m_Items(i).RECT, X, Y) Then
                If Not m_Items(i).Value And Not m_Items(i).Disabled Then
                    
                    m_Items(i).Value = True
                    If m_Orientation = Horizontal Then
                       NewXY = m_Items(i).RECT.Left
                    Else
                       NewXY = m_Items(i).RECT.Top
                    End If
                    m_Selected = i
                    InitTime = Timer
                    Timer1.Interval = 1
                    Index = i
                End If
            Else
               m_Items(i).Value = False
            End If
        Next
        
    ElseIf m_CtrlType = CheckBar Then
        For i = 0 To Count - 1
            If m_Items(i).Visible And PtInRect(m_Items(i).RECT, X, Y) Then
                If Not m_Items(i).Disabled Then
                    m_Items(i).Value = Not m_Items(i).Value
                    UserControl.Refresh
                    Index = i
                    Exit For
                End If
            End If
        Next
    End If
    
    If Index > -1 Then RaiseEvent ItemClick(Index)
    RaiseEvent Click
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    For i = 0 To Count - 1
        If m_Items(i).Visible And PtInRect(m_Items(i).RECT, X, Y) Then
           If i <> m_HotItem Then
            If m_HotItem <> -1 Then RaiseEvent ItemMouseLeave(m_HotItem)
            m_HotItem = i
            
            UserControl.Extender.ToolTipText = m_Items(i).ToolTipText
            If m_Items(i).Disabled Then
                UserControl.MousePointer = vbNoDrop
            Else
                UserControl.MousePointer = vbCustom
            End If
            UserControl.Refresh
            RaiseEvent ItemMouseEnter(m_HotItem)
           End If
           Exit For
        End If
    Next

    If hCur Then SetCursor hCur
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Let MousePointer(ByVal newValue As MousePointerConstants)
    UserControl.MousePointer = newValue
    PropertyChanged "MousePointer"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Set MouseIcon(ByVal newValue As IPictureDisp)
    UserControl.MouseIcon = newValue
    PropertyChanged "MouseIcon"
End Property

Public Property Get MouseIcon() As IPictureDisp
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MousePointerHands(ByVal newValue As Boolean)
    m_MousePointerHands = newValue
    If newValue Then
        If Ambient.UserMode Then
            UserControl.MousePointer = vbCustom
            UserControl.MouseIcon = GetSystemHandCursor
        End If
    Else
        If hCur Then DestroyCursor hCur: hCur = 0
        UserControl.MousePointer = vbDefault
        UserControl.MouseIcon = Nothing
    End If
    PropertyChanged "MousePointerHands"
End Property

Public Property Get MousePointerHands() As Boolean
    MousePointerHands = m_MousePointerHands
End Property

Public Function GetSystemHandCursor() As Picture
    Dim Pic As PicBmp, IPic As IPicture, GUID(0 To 3) As Long
    
    If hCur Then DestroyCursor hCur: hCur = 0
    
    hCur = LoadCursor(ByVal 0&, IDC_HAND)
     
    GUID(0) = &H7BF80980
    GUID(1) = &H101ABF32
    GUID(2) = &HAA00BB8B
    GUID(3) = &HAB0C3000
 
    With Pic
        .Size = Len(Pic)
        .type = vbPicTypeIcon
        .hBmp = hCur
        .hPal = 0
    End With
 
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
 
    Set GetSystemHandCursor = IPic
    
End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    
    If lAlpha < 0 Then lAlpha = 0
    If (clrFirst And &H80000000) Then clrFirst = GetSysColor(clrFirst And &HFF&)
    If (clrSecond And &H80000000) Then clrSecond = GetSysColor(clrSecond And &HFF&)
  
    CopyMemory clrFore(0), clrFirst, 4
    CopyMemory clrBack(0), clrSecond, 4
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

Private Function BitmapToArray(lBitmap As Long) As Byte()
Dim oStream As IUnknown
Dim eGUID   As GUID

    Set oStream = pvStreamFromArray(0&, 0&)
    If Not oStream Is Nothing Then
        CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), eGUID
        If GdipSaveImageToStream(lBitmap, oStream, eGUID, ByVal 0&) = 0& Then
            Call pvStreamToArray(ObjPtr(oStream), BitmapToArray)
        End If
    End If
End Function

Private Function pvStreamFromArray(ArrayPtr As Long, Length As Long) As stdole.IUnknown
On Error GoTo e
Dim o_hMem As Long
Dim o_lpMem  As Long
     
    If ArrayPtr = 0& Then
        CreateStreamOnHGlobal 0&, 1&, pvStreamFromArray
    ElseIf Length <> 0& Then
        o_hMem = GlobalAlloc(&H2&, Length)
        If o_hMem <> 0 Then
            o_lpMem = GlobalLock(o_hMem)
            If o_lpMem <> 0 Then
                CopyMemory ByVal o_lpMem, ByVal ArrayPtr, Length
                Call GlobalUnlock(o_hMem)
                Call CreateStreamOnHGlobal(o_hMem, 1&, pvStreamFromArray)
            End If
        End If
    End If
    
e:
End Function


Private Function pvStreamToArray(hStream As Long, arrayBytes() As Byte) As Boolean
Dim o_hMem        As Long
Dim o_lpMem       As Long
Dim o_lByteCount  As Long
    
    If hStream Then
        If GetHGlobalFromStream(ByVal hStream, o_hMem) = 0 Then
            o_lByteCount = GlobalSize(o_hMem)
            If o_lByteCount > 0 Then
                o_lpMem = GlobalLock(o_hMem)
                If o_lpMem <> 0 Then
                    ReDim arrayBytes(0 To o_lByteCount - 1)
                    
                    CopyMemory arrayBytes(0), ByVal o_lpMem, o_lByteCount
                    GlobalUnlock o_hMem
                    pvStreamToArray = True
                End If
            End If
        End If
        
    End If
End Function
