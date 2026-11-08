VERSION 5.00
Begin VB.UserControl ucSlider 
   BackColor       =   &H00C6C6C6&
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   ClipBehavior    =   0  'None
   PropertyPages   =   "ucSlider.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   318
   Windowless      =   -1  'True
   Begin VB.Timer tmrMOUSEOVER 
      Left            =   1320
      Top             =   1800
   End
End
Attribute VB_Name = "ucSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'----------------------------------
'Module Name: ucSlider
'Autor: Leandro Ascierto
'Web: www.leandroascierto.com
'Date: 24/01/2022
'Version 1.0.0
'----------------------------------
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal HDC As Long, ByRef lpRect As RECT) As Long
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal length As Long)
Private Declare Function MulDiv Lib "Kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function PtInRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "Oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByVal lColorRef As Long) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "Ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
Private Declare Function GdiplusStartup Lib "gdiplus" (token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As Any, ByRef Image As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal HDC As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mImage As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipImageRotateFlip Lib "gdiplus" (ByVal Image As Long, ByVal rfType As RotateFlipType) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipAddPathRectangleI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RectF, ByVal mStringFormat As Long, ByRef mBoundingBox As RectF, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RectF, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
'Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As StringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As StringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As StringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectI Lib "GdiPlus.dll" (ByRef mRect As RectL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mMode As LinearGradientMode, ByVal mWrapMode As WrapMode, ByRef mLineGradient As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Type RectF
    Left As Single
    Top As Single
    Width As Single
    Height As Single
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

Private Type RectL
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

Private Enum CaptionAlignmentH
    cLeft
    ccEnter
    cRight
End Enum

Private Enum CaptionAlignmentV
    cTop
    cMiddle
    cBottom
End Enum

Private Enum RotateFlipType
    RotateNoneFlipNone = &H0
    Rotate90FlipNone = &H1
    Rotate180FlipNone = &H2
    Rotate270FlipNone = &H3
    RotateNoneFlipX = &H4
    Rotate90FlipX = &H5
    Rotate180FlipX = &H6
    Rotate270FlipX = &H7
    RotateNoneFlipY = &H6
    Rotate90FlipY = &H7
    Rotate180FlipY = &H4
    Rotate270FlipY = &H5
    RotateNoneFlipXY = &H2
    Rotate90FlipXY = &H3
    Rotate180FlipXY = &H0
    Rotate270FlipXY = &H1
End Enum

Private Enum LinearGradientMode
    LinearGradientModeHorizontal = &H0
    LinearGradientModeVertical = &H1
    LinearGradientModeForwardDiagonal = &H2
    LinearGradientModeBackwardDiagonal = &H3
End Enum

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum
  
Private Enum StringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum
  
'Public Enum StringTrimming
'    StringTrimmingNone = &H0
'    StringTrimmingCharacter = &H1
'    StringTrimmingWord = &H2
'    StringTrimmingEllipsisCharacter = &H3
'    StringTrimmingEllipsisWord = &H4
'    StringTrimmingEllipsisPath = &H5
'End Enum

'Public Enum StringFormatFlags
'    StringFormatFlagsNone = &H0
'    StringFormatFlagsDirectionRightToLeft = &H1
'    StringFormatFlagsDirectionVertical = &H2
'    StringFormatFlagsNoFitBlackBox = &H4
'    StringFormatFlagsDisplayFormatControl = &H20
'    StringFormatFlagsNoFontFallback = &H400
'    StringFormatFlagsMeasureTrailingSpaces = &H800
'    StringFormatFlagsNoWrap = &H1000
'    StringFormatFlagsLineLimit = &H2000
'    StringFormatFlagsNoClip = &H4000
'End Enum

Private Enum HotkeyPrefix
    HotkeyPrefixNone = &H0
    HotkeyPrefixShow = &H1
    HotkeyPrefixHide = &H2
End Enum

Private Enum WrapMode
    WrapModeTile = &H0
    WrapModeTileFlipX = &H1
    WrapModeTileFlipy = &H2
    WrapModeTileFlipXY = &H3
    WrapModeClamp = &H4
End Enum

Public Enum eTickStyle
    sldLeftTop
    sldBottomRight
    sldBoth
    sldNone
End Enum

Public Enum eKnobStyle
    ksElipse
    ksRoundRect
    ksRectangle
End Enum

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event MouseOver()
Public Event MouseOut()
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event Change()

Private Const IDC_HAND                  As Long = 32649
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const UnitPixel                 As Long = &H2&
Private Const LOGPIXELSX                As Long = 88
Private Const LOGPIXELSY                As Long = 90
Private Const SmoothingModeAntiAlias    As Long = 4

Dim mFocus As Boolean
Dim nScale As Long
Dim m_Enter As Boolean
Dim c_lhWnd As Long
Dim m_PT As POINTAPI
Dim m_Left As Long
Dim m_Top As Long
Dim bIntercept As Boolean
Dim m_Over As Boolean
Dim hCur As Long
Dim GdipToken As Long
Dim ArrImgKnob() As Byte
Dim hImgKnob As Long
Dim PosDif As Long
Dim m_GradientStyle As Boolean
Dim m_TrackSize As Long
Dim m_TrackBackColor As OLE_COLOR
Dim m_TrackForeColorPositive As OLE_COLOR
Dim m_TrackForeColorNegative As OLE_COLOR
Dim m_TickTextColor As OLE_COLOR
Dim m_TickLinesColor As OLE_COLOR
Dim m_TickLines As eTickStyle
Dim m_TickText As eTickStyle
Dim m_BiselColor As OLE_COLOR
Dim m_BiselSize As Long
Dim m_KnobColor As OLE_COLOR
Dim m_TicksLongFrequency As Long
Dim m_TicksSmallHiden  As Boolean
Dim m_TicksSize As Long
Dim m_TickDistance As Long
Dim m_KnobStyle As eKnobStyle
Dim m_KnobWidth As Long
Dim m_KnobHeight As Long
Dim m_Min As Single
Dim m_Max As Single
Dim m_Value As Single
Dim m_MousePointerHands As Long
Dim mPercent As Single
Dim m_OrientationHorizontal As Boolean
Dim MA As Long 'MarginA
Dim MB As Long 'MarginB
Dim RectKnob As RectL
'Private WithEvents Ext As VBControlExtender

Public Property Get BiselSize() As Long
    BiselSize = m_BiselSize
End Property

Public Property Let BiselSize(ByVal New_Value As Long)
    m_BiselSize = New_Value
    PropertyChanged "BiselSize"
    Refresh
End Property


Public Property Let BiselColor(ByVal New_Value As OLE_COLOR)
    m_BiselColor = New_Value
    PropertyChanged "BiselColor"
    Refresh
End Property

Public Property Get BiselColor() As OLE_COLOR
    BiselColor = m_BiselColor
End Property

Public Property Let GradientStyle(ByVal New_Value As Boolean)
    m_GradientStyle = New_Value
    PropertyChanged "GradientStyle"
    Refresh
End Property

Public Property Get GradientStyle() As Boolean
    GradientStyle = m_GradientStyle
End Property

Public Property Get KnobStyle() As eKnobStyle
    KnobStyle = m_KnobStyle
End Property

Public Property Let KnobStyle(ByVal New_Value As eKnobStyle)
    m_KnobStyle = New_Value
    PropertyChanged "KnobStyle"
    Refresh
End Property

Public Property Get KnobTop() As Long
    KnobTop = RectKnob.Top
End Property

Public Property Get KnobLeft() As Long
    KnobLeft = RectKnob.Left
End Property

Public Property Let KnobWidth(ByVal New_Value As Long)
    m_KnobWidth = New_Value
    PropertyChanged "KnobWidth"
    Refresh
End Property

Public Property Get KnobWidth() As Long
    KnobWidth = m_KnobWidth
End Property

Public Property Let KnobHeight(ByVal New_Value As Long)
    m_KnobHeight = New_Value
    PropertyChanged "KnobHeight"
    Refresh
End Property

Public Property Get KnobHeight() As Long
    KnobHeight = m_KnobHeight
End Property

Public Property Let KnobColor(ByVal New_Value As OLE_COLOR)
    m_KnobColor = New_Value
    PropertyChanged "KnobColor"
    Refresh
End Property

Public Property Get KnobColor() As OLE_COLOR
    KnobColor = m_KnobColor
End Property

Public Function LoadImgKnob(ByRef bvStream() As Byte) As Boolean
    If LoadImageFromStream(bvStream, hImgKnob) Then
        ArrImgKnob = bvStream
        If m_OrientationHorizontal = True Then
            GdipImageRotateFlip hImgKnob, Rotate90FlipNone 'Rotate90FlipXY
        End If
    End If
    Call PropertyChanged("ArrImgKnob")
    Refresh
End Function

Public Function ImgKnob() As Byte()
    ImgKnob = ArrImgKnob
End Function

Public Sub DeleteImgKnob()
    If hImgKnob <> 0 Then GdipDisposeImage hImgKnob: hImgKnob = 0
    Erase ArrImgKnob
    Call PropertyChanged("ArrImgKnob")
    Refresh
End Sub

Public Property Let TrackBackColor(ByVal New_Value As OLE_COLOR)
    m_TrackBackColor = New_Value
    PropertyChanged "TrackBackColor"
    Refresh
End Property

Public Property Get TrackBackColor() As OLE_COLOR
    TrackBackColor = m_TrackBackColor
End Property


Public Property Let TrackForeColorPositive(ByVal New_Value As OLE_COLOR)
    m_TrackForeColorPositive = New_Value
    PropertyChanged "TrackForeColorPositive"
    Refresh
End Property

Public Property Get TrackForeColorPositive() As OLE_COLOR
    TrackForeColorPositive = m_TrackForeColorPositive
End Property

Public Property Let TrackForeColorNegative(ByVal New_Value As OLE_COLOR)
    m_TrackForeColorNegative = New_Value
    PropertyChanged "TrackForeColorNegative"
    Refresh
End Property

Public Property Get TrackForeColorNegative() As OLE_COLOR
    TrackForeColorNegative = m_TrackForeColorNegative
End Property

Public Property Let TrackSize(ByVal New_Value As Long)
    m_TrackSize = New_Value
    PropertyChanged "TrackSize"
    Refresh
End Property

Public Property Get TrackSize() As Long
    TrackSize = m_TrackSize
End Property

Public Property Let TickText(ByVal New_Value As eTickStyle)
    m_TickText = New_Value
    PropertyChanged "TickText"
    Refresh
End Property

Public Property Get TickText() As eTickStyle
    TickText = m_TickText
End Property

Public Property Let TickLines(ByVal New_Value As eTickStyle)
    m_TickLines = New_Value
    PropertyChanged "TickLines"
    Refresh
End Property

Public Property Get TickLines() As eTickStyle
    TickLines = m_TickLines
End Property

Public Property Let TickDistance(ByVal New_Value As Long)
    m_TickDistance = New_Value
    PropertyChanged "TickDistance"
    Refresh
End Property

Public Property Get TickDistance() As Long
    TickDistance = m_TickDistance
End Property

Public Property Let TickLinesColor(ByVal New_Value As OLE_COLOR)
    m_TickLinesColor = New_Value
    PropertyChanged "TickLinesColor"
    Refresh
End Property

Public Property Get TickLinesColor() As OLE_COLOR
    TickLinesColor = m_TickLinesColor
End Property

Public Property Let TickTextColor(ByVal New_Value As OLE_COLOR)
    m_TickTextColor = New_Value
    PropertyChanged "TickTextColor"
    Refresh
End Property

Public Property Get TickTextColor() As OLE_COLOR
    TickTextColor = m_TickTextColor
End Property

Public Property Get TicksSize() As Long
    TicksSize = m_TicksSize
End Property

Public Property Let TicksSize(ByVal New_Value As Long)
    m_TicksSize = New_Value
    PropertyChanged "TicksSize"
    Refresh
End Property

Public Property Get OrientationHorizontal() As Boolean
    OrientationHorizontal = m_OrientationHorizontal
End Property

Public Property Let OrientationHorizontal(ByVal New_Value As Boolean)
    m_OrientationHorizontal = New_Value
    PropertyChanged "OrientationHorizontal"
    Refresh
End Property


Public Sub Draw()
    Dim hGraphics As Long
    Dim hPath As Long
    Dim hBrush As Long
    Dim hPen As Long
    Dim W As Long, H As Long, T As Long, L As Long
    Dim Color1 As Long
    Dim Color2 As Long
    Dim BS As Single 'BiselSize
    Dim TS As Single 'TrackWidth
    Dim RL As RectL
    Dim KH As Long 'KnobHeight
    Dim KW As Long 'KnobWidth
    Dim KR As Long 'KnobRadio
    Dim lTicksSize As Long
    Dim mTotalLines As Long
    Dim i As Long
    Dim bDrawTick As Long
    Dim Mid As Long
    Dim TickDist As Long
    Dim p As Single
    Dim LineTop  As Long
    Dim LineLeft  As Long
    Dim TextH As Long, TextW As Long
    Dim R As Long
    
    KW = m_KnobWidth * nScale
    KH = m_KnobHeight * nScale
    
    If m_KnobStyle = ksElipse Then
        KR = KW
    ElseIf m_KnobStyle = ksRoundRect Then
        KR = KW / 4
    Else
        KR = 0
    End If

    TS = m_TrackSize * nScale
    
    mTotalLines = m_Max - m_Min
   
    BS = m_BiselSize * nScale
         
    If m_OrientationHorizontal Then
        MA = IIf(BS > KW / 2, BS, KW / 2)
        MB = IIf(BS > KW / 1.6, BS, KW / 1.6)
        W = UserControl.ScaleWidth - MA - MB - 1
        H = UserControl.ScaleHeight - 1
        L = MA + (W * (CLng((W * mPercent / 100) / (W / mTotalLines)) / mTotalLines))
        RectKnob.Left = L - KW / 2
        RectKnob.Top = H / 2 - KH / 2
    Else
        MA = IIf(BS > KH / 2, BS, KH / 2)
        MB = IIf(BS > KH / 1.6, BS, KH / 1.6)
        W = UserControl.ScaleWidth - 1
        H = UserControl.ScaleHeight - MA - MB - 1
        T = MA + (H * (CLng((H * mPercent / 100) / (H / mTotalLines)) / mTotalLines))
        RectKnob.Top = T - KH / 2
        RectKnob.Left = W / 2 - KW / 2
    End If

    If GdipCreateFromHDC(HDC, hGraphics) <> 0 Then Exit Sub
    GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
    
    'Base Gradient, bisel 1
    If m_GradientStyle Then
        Color1 = ConvertColor(ShiftColor(vbBlack, m_BiselColor, 100), 100)
        Color2 = ConvertColor(ShiftColor(vbWhite, m_BiselColor, 200), 100)
        
        If m_OrientationHorizontal Then
            RL.Left = MA - BS
            RL.Top = H / 2 - TS / 2 - BS
            RL.Height = TS + BS * 2
            RL.Width = W + BS * 2
            GdipCreateLineBrushFromRectI RL, Color1, Color2, LinearGradientModeVertical, WrapModeTileFlipX, hBrush
        Else
            RL.Left = W / 2 - TS / 2 - BS
            RL.Top = MA - BS
            RL.Width = TS + BS * 2
            RL.Height = H + BS * 2
            GdipCreateLineBrushFromRectI RL, Color1, Color2, LinearGradientModeHorizontal, WrapModeTileFlipX, hBrush
        End If
        
    
        hPath = CreateRoundPath(RL.Left, RL.Top, RL.Width, RL.Height, TS * 2)
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
        GdipDeletePath hPath
    End If
    
    'Track
    If m_OrientationHorizontal Then
        RL.Left = MA
        RL.Top = H / 2 - TS / 2
        RL.Width = W
        RL.Height = TS
    Else
        RL.Left = W / 2 - TS / 2
        RL.Top = MA
        RL.Width = TS
        RL.Height = H
    End If

    hPath = CreateRoundPath(RL.Left, RL.Top, RL.Width, RL.Height, TS)
    If m_GradientStyle Then
        Color1 = ConvertColor(ShiftColor(vbBlack, m_TrackBackColor, 10), 100)
        Color2 = ConvertColor(ShiftColor(vbWhite, m_TrackBackColor, 30), 100)
        GdipCreateLineBrushFromRectI RL, Color1, Color2, LinearGradientModeVertical, WrapModeTileFlipX, hBrush
    Else
        GdipCreateSolidFill ConvertColor(m_TrackBackColor, 100), hBrush
    End If

    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
'    If mFocus Then
'        GdipCreatePen1 ConvertColor(vbHighlight, 50), nScale * 1, UnitPixel, hPen
'        GdipDrawPath hGraphics, hPen, hPath
'        GdipDeletePen hPen
'    End If
    GdipDeletePath hPath
    
    
    '-------------

    If m_Max > 0 Then
        If m_OrientationHorizontal Then
            If m_Min >= 0 Then
                RL.Left = MA
                RL.Top = H / 2 - TS / 2
                RL.Width = L - MA
                RL.Height = TS
                R = TS
            Else
                p = Abs(m_Min) / mTotalLines
                RL.Left = MA + (W * p)
                RL.Top = H / 2 - TS / 2
                RL.Width = L - RL.Left
                RL.Height = TS
                R = 0
            End If
        Else
            If m_Min >= 0 Then
                RL.Left = W / 2 - TS / 2
                RL.Top = T
                RL.Width = TS
                RL.Height = MA + H - T
                R = TS
            Else
                p = Abs(m_Min) / mTotalLines
                RL.Left = W / 2 - TS / 2
                RL.Top = T
                RL.Width = TS
                RL.Height = MA + H - T - (H * p)
                R = 0
            End If
        End If
        
        
        If RL.Width > 0 And RL.Height > 0 Then
            hPath = CreateRoundPath(RL.Left, RL.Top, RL.Width, RL.Height, R)
            
            If m_GradientStyle Then
                Color1 = ConvertColor(ShiftColor(vbBlack, m_TrackForeColorPositive, 70), 100)
                Color2 = ConvertColor(ShiftColor(vbWhite, m_TrackForeColorPositive, 70), 100)
                GdipCreateLineBrushFromRectI RL, Color1, Color2, LinearGradientModeVertical, WrapModeTileFlipX, hBrush
            Else
                GdipCreateSolidFill ConvertColor(m_TrackForeColorPositive, 100), hBrush
            End If
            
            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
        End If
    End If
    
    If m_Min < 0 Then
        If m_OrientationHorizontal Then
            If m_Max < 0 Then
                RL.Left = L
                RL.Top = H / 2 - TS / 2
                RL.Width = MA + W - L
                RL.Height = TS
                R = TS
            Else
                p = Abs(m_Min) / mTotalLines
                RL.Left = L
                RL.Top = H / 2 - TS / 2
                RL.Width = MA + (W * p) - L
                RL.Height = TS
                R = 0
            End If
        Else
            If m_Max < 0 Then
                RL.Left = W / 2 - TS / 2
                RL.Top = MA
                RL.Width = TS
                RL.Height = T
                R = TS
            Else
                p = Abs(m_Min) / mTotalLines
                RL.Left = W / 2 - TS / 2
                RL.Top = MA + H - (H * p)
                RL.Width = TS
                RL.Height = T - RL.Top
                R = 0
            End If

        End If
        If RL.Width > 0 And RL.Height > 0 Then
            hPath = CreateRoundPath(RL.Left, RL.Top, RL.Width, RL.Height, R)
            
            If m_GradientStyle Then
                Color1 = ConvertColor(ShiftColor(vbBlack, m_TrackForeColorNegative, 70), 100)
                Color2 = ConvertColor(ShiftColor(vbWhite, m_TrackForeColorPositive, 70), 100)
                GdipCreateLineBrushFromRectI RL, Color1, Color2, m_TrackForeColorNegative, WrapModeTileFlipX, hBrush
            Else
                GdipCreateSolidFill ConvertColor(m_TrackForeColorNegative, 100), hBrush
            End If

            GdipFillPath hGraphics, hBrush, hPath
            GdipDeleteBrush hBrush
            GdipDeletePath hPath
        End If
    End If
    
    '---------------------
    TickDist = m_TickDistance * nScale
    If m_OrientationHorizontal Then
        Mid = H / 2
    Else
        Mid = W / 2
    End If
        
    TextH = UserControl.TextHeight("10") / nScale

    
    GdipCreatePen1 ConvertColor(m_TickLinesColor, 100), nScale, UnitPixel, hPen
    
    For i = 0 To mTotalLines
        p = i * 100 / mTotalLines


        If i Mod m_TicksLongFrequency = 0 Then
            bDrawTick = True
            lTicksSize = m_TicksSize * 2 * nScale
        Else
            If m_TicksSmallHiden Then bDrawTick = False
            lTicksSize = m_TicksSize * nScale
        End If
    
        If bDrawTick Then
            If m_OrientationHorizontal Then
                
                LineLeft = MA + (W * p / 100)
               
                If m_TickLines = sldLeftTop Or m_TickLines = sldBoth Then
                    GdipDrawLine hGraphics, hPen, LineLeft, Mid - TickDist, LineLeft, Mid - TickDist - lTicksSize
                End If
                
                If m_TickText = sldLeftTop Or m_TickText = sldBoth Then
                    If i Mod m_TicksLongFrequency = 0 Then
                        TextW = UserControl.TextWidth(i + m_Min) + 20
                        DrawText hGraphics, i + m_Min, LineLeft / nScale - TextW / 2, (Mid - TickDist - lTicksSize * 2) / nScale - TextH, TextW, TextH, UserControl.Font, m_TickTextColor, 100, ccEnter, cMiddle, False
                    End If
                End If
                
                If m_TickLines = sldBottomRight Or m_TickLines = sldBoth Then
                    GdipDrawLine hGraphics, hPen, LineLeft, Mid + TickDist, LineLeft, Mid + TickDist + lTicksSize
                End If
                
                If m_TickText = sldBottomRight Or m_TickText = sldBoth Then
                    If i Mod m_TicksLongFrequency = 0 Then
                        TextW = UserControl.TextWidth(i + m_Min) + 20
                        DrawText hGraphics, i + m_Min, LineLeft / nScale - TextW / 2, (Mid + TickDist + lTicksSize * 2) / nScale, TextW, TextH, UserControl.Font, m_TickTextColor, 100, ccEnter, cMiddle, False
                    End If
                End If
            Else
                LineTop = MA + H - (H * p / 100)
               
                If m_TickLines = sldLeftTop Or m_TickLines = sldBoth Then
                    GdipDrawLine hGraphics, hPen, Mid - TickDist, LineTop, Mid - TickDist - lTicksSize, LineTop
                End If
                
                If m_TickText = sldLeftTop Or m_TickText = sldBoth Then
                    If i Mod m_TicksLongFrequency = 0 Then
                        DrawText hGraphics, i + m_Min, (Mid - TickDist - lTicksSize * 2) / nScale - 100, LineTop / nScale - TextH / 2, 100, TextH, UserControl.Font, m_TickTextColor, 100, cRight, cMiddle, False
                    End If
                End If
                
                If m_TickLines = sldBottomRight Or m_TickLines = sldBoth Then
                    GdipDrawLine hGraphics, hPen, Mid + TickDist, LineTop, Mid + TickDist + lTicksSize, LineTop
                End If
                
                If m_TickText = sldBottomRight Or m_TickText = sldBoth Then
                    If i Mod m_TicksLongFrequency = 0 Then
                        DrawText hGraphics, i + m_Min, (Mid + TickDist + lTicksSize * 2) / nScale, LineTop / nScale - TextH / 2, 100, TextH, UserControl.Font, m_TickTextColor, 100, cLeft, cMiddle, False
                    End If
                End If
            End If
        End If
    Next

    GdipDeletePen hPen
    
    
    If hImgKnob <> 0 Then
        If m_OrientationHorizontal Then
            RL.Left = L - KW / 2
            RL.Top = H / 2 - KH / 2
            RL.Width = KW
            RL.Height = KH
        Else
            RL.Left = W / 2 - KW / 2
            RL.Top = T - KH / 2
            RL.Width = KW
            RL.Height = KH
        End If
        GdipDrawImageRectI hGraphics, hImgKnob, RL.Left, RL.Top, RL.Width, RL.Height
        GdipDeleteGraphics hGraphics
        If mFocus Then
            Dim RECT As RECT
            RECT.Right = UserControl.ScaleWidth
            RECT.Bottom = UserControl.ScaleHeight
            UserControl.DrawWidth = nScale
            DrawFocusRect HDC, RECT
        End If
        Exit Sub
    End If
    
    '-----------------------------
    'Knob Shadow
    If m_OrientationHorizontal Then
        RL.Left = L - KW / 2
        RL.Top = H / 2 - KH / 2
        RL.Width = KW * 1.1
        RL.Height = KH * 1.1
    Else
        RL.Left = W / 2 - KW / 2
        RL.Top = T - KH / 2
        RL.Width = KW * 1.1
        RL.Height = KH * 1.1
    End If
    Color1 = ConvertColor(vbBlack, 50)
    Color2 = 0

    hPath = CreateRoundPath(RL.Left, RL.Top, RL.Width, RL.Height, KR)
    
    GdipCreatePathGradientFromPath hPath, hBrush
    
    GdipSetPathGradientCenterColor hBrush, Color1              'Asigna un color central a la brocha
    GdipSetPathGradientSurroundColorsWithCount hBrush, Color2, 1
    
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeletePath hPath
    
    'Knob
    RL.Width = KW
    RL.Height = KH
    If m_GradientStyle Then
        Color1 = ConvertColor(ShiftColor(vbWhite, m_KnobColor, 200), 100)
        Color2 = ConvertColor(ShiftColor(vbBlack, m_KnobColor, 50), 100)
        GdipCreateLineBrushFromRectI RL, Color1, Color2, LinearGradientModeVertical, WrapModeTileFlipX, hBrush
    Else
        GdipCreateSolidFill ConvertColor(m_KnobColor, 100), hBrush
    End If

    If m_GradientStyle Then
        hPath = CreateRoundPath(RL.Left, RL.Top, RL.Width, RL.Height, KR)
        GdipCreatePen1 ConvertColor(vbBlack, 10), 1 * nScale, UnitPixel, hPen
    Else
        hPath = CreateRoundPath(RL.Left + nScale, RL.Top + nScale, RL.Width - nScale * 2, RL.Height - nScale * 2, KR)
        GdipCreatePen1 ConvertColor(ShiftColor(vbBlack, m_KnobColor, 50), 100), 1 * nScale * 2, UnitPixel, hPen
    End If
    
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    
    GdipDrawPath hGraphics, hPen, hPath
    GdipDeletePen hPen
    
    GdipDeletePath hPath
    
    
    If m_GradientStyle Then
        'Flip Knob
        RL.Left = RL.Left + KW / 6
        RL.Top = RL.Top + KH / 6
        RL.Width = KW / 1.5
        RL.Height = KH / 1.5
        Color1 = ConvertColor(ShiftColor(vbBlack, m_KnobColor, 50), 100)
        Color2 = ConvertColor(ShiftColor(vbWhite, m_KnobColor, 200), 100)
        GdipCreateLineBrushFromRectI RL, Color1, Color2, LinearGradientModeVertical, WrapModeTileFlipX, hBrush
        hPath = CreateRoundPath(RL.Left, RL.Top, RL.Width, RL.Height, KR / 1.5)
        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush
        GdipDeletePath hPath
    End If

    GdipDeleteGraphics hGraphics
    
    If mFocus Then
        RECT.Right = UserControl.ScaleWidth
        RECT.Bottom = UserControl.ScaleHeight
        UserControl.DrawWidth = nScale
        DrawFocusRect HDC, RECT
    End If
End Sub


Private Function DrawText(ByVal hGraphics As Long, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal oFont As StdFont, ByVal lForeColor As OLE_COLOR, Optional ByVal ColorOpacity As Integer = 100, Optional HAlign As CaptionAlignmentH, Optional VAlign As CaptionAlignmentV, Optional bWordWrap As Boolean) As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim layoutRect As RectF
    Dim lFontSize As Long
    Dim lFontStyle As GDIPLUS_FONTSTYLE
    Dim hFont As Long

    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
        If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
    End If
    
    If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
        'If Not bWordWrap Then GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
        GdipSetStringFormatAlign hFormat, HAlign
        GdipSetStringFormatLineAlign hFormat, VAlign
    End If
        
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
        

    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(HDC, LOGPIXELSY), 72)

    layoutRect.Left = X * nScale: layoutRect.Top = Y * nScale
    layoutRect.Width = Width * nScale: layoutRect.Height = Height * nScale

    GdipCreateSolidFill ConvertColor(lForeColor, ColorOpacity), hBrush

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipDrawString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, hBrush
    
    Dim BB As RectF, CF As Long, LF As Long

    Call GdipCreateFont(hFontFamily, lFontSize, lFontStyle, UnitPixel, hFont)
    GdipMeasureString hGraphics, StrPtr(text), -1, hFont, layoutRect, hFormat, BB, CF, LF

    
    If bWordWrap Then
        DrawText = BB.Height / nScale
    Else
        DrawText = BB.Width / nScale
    End If
    
    GdipDeleteFont hFont
    GdipDeleteBrush hBrush
    GdipDeleteStringFormat hFormat
    GdipDeleteFontFamily hFontFamily
End Function

Private Sub UserControl_EnterFocus()
    If UserControl.Enabled Then
        mFocus = True
        Refresh
        'Set Ext = UserControl.Extender
        modIOleInPlaceActiveObject.SetIPAO Me
    End If
End Sub

Private Sub UserControl_ExitFocus()
    mFocus = False
    Refresh
    'Set Ext = Nothing
End Sub

'NOT WORK WITH ARRAY CONTROLS OR UNIQUE CONTROL
''prevents ArrowKey-based Focus-changes in windowless Controls (O.Schmidt 2010)
'Private Sub Ext_Validate(Cancel As Boolean)  'simply set Cancel=True, to prevent Focus-loss in case of ArrowKeys
'  Dim Shift As Integer
'  If GetAsyncKeyState(vbKeyShift) And &H8000 Then Shift = Shift Or vbShiftMask
'  If GetAsyncKeyState(vbKeyControl) And &H8000 Then Shift = Shift Or vbCtrlMask
'  If GetAsyncKeyState(vbKeyMenu) And &H8000 Then Shift = Shift Or vbAltMask
'
'  If GetAsyncKeyState(vbKeyLeft) And &H8000 Then Cancel = True: UserControl_KeyDown vbKeyLeft, Shift
'  If GetAsyncKeyState(vbKeyRight) And &H8000 Then Cancel = True: UserControl_KeyDown vbKeyRight, Shift
'  If GetAsyncKeyState(vbKeyUp) And &H8000 Then Cancel = True: UserControl_KeyDown vbKeyUp, Shift
'  If GetAsyncKeyState(vbKeyDown) And &H8000 Then Cancel = True: UserControl_KeyDown vbKeyDown, Shift
'End Sub

Private Sub UserControl_Initialize()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
    nScale = GetWindowsDPI
    modIOleInPlaceActiveObject.InitIPAO
End Sub

Private Function CreateRoundPath(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal Radius As Single) As Long
    Dim hPath As Long
    If GdipCreatePath(&H0, hPath) = 0& Then
    
        If Radius > Width / 2 Then Radius = Width / 2
        If Radius > Height / 2 Then Radius = Height / 2
    
        If Radius = 0 Then
            GdipAddPathRectangleI hPath, Left, Top, Width, Height
        Else
            Radius = Radius * 2
            GdipAddPathArcI hPath, Left, Top, Radius, Radius, 180, 90
            GdipAddPathArcI hPath, Left + Width - Radius, Top, Radius, Radius, 270, 90
            GdipAddPathArcI hPath, Left + Width - Radius, Top + Height - Radius, Radius, Radius, 0, 90
            GdipAddPathArcI hPath, Left, Top + Height - Radius, Radius, Radius, 90, 90
            GdipClosePathFigure hPath
        End If
        CreateRoundPath = hPath
    End If
    
End Function

Public Property Get Min() As Single
    Min = m_Min
End Property

Public Property Let Min(ByVal New_Value As Single)
    m_Min = New_Value
    mPercent = 100 - (m_Value - m_Min) * 100 / (m_Max - m_Min)
    PropertyChanged "Min"
    Refresh
End Property

Public Property Get Max() As Single
    Max = m_Max
End Property

Public Property Let Max(ByVal New_Value As Single)
    m_Max = New_Value
    mPercent = 100 - (m_Value - m_Min) * 100 / (m_Max - m_Min)
    PropertyChanged "Max"
    Refresh
End Property

Public Property Get Value() As Single
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Single)
    If New_Value < m_Min Then New_Value = m_Min
    If New_Value > m_Max Then New_Value = m_Max
    m_Value = New_Value
    
    If m_OrientationHorizontal Then
        mPercent = (m_Value - m_Min) * 100 / (m_Max - m_Min)
    Else
        mPercent = 100 - (m_Value - m_Min) * 100 / (m_Max - m_Min)
    End If
    
    RectKnob.Width = m_KnobWidth
    RectKnob.Height = m_KnobHeight
    
    If mPercent < 0 Then mPercent = 0
    If mPercent > 100 Then mPercent = 100

    PropertyChanged "Value"
    Refresh
    
    RaiseEvent Change
End Property

Public Property Get TicksLongFrequency() As Long
    TicksLongFrequency = m_TicksLongFrequency
End Property

Public Property Let TicksLongFrequency(ByVal New_Value As Long)
    m_TicksLongFrequency = New_Value
    PropertyChanged "TicksLongFrequency"
    Refresh
End Property

Public Property Get TicksSmallHiden() As Boolean
    TicksSmallHiden = m_TicksSmallHiden
End Property

Public Property Let TicksSmallHiden(ByVal New_Value As Boolean)
    m_TicksSmallHiden = New_Value
    PropertyChanged "TicksSmallHiden"
    Refresh
End Property

Public Function GetWindowsDPI() As Double
    Dim HDC As Long, LPX  As Double
    HDC = GetDC(0)
    LPX = CDbl(GetDeviceCaps(HDC, LOGPIXELSX))
    ReleaseDC 0, HDC

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = UserControl.Font
End Property

Public Property Set Font(New_Font As StdFont)
    With UserControl.Font
        .Name = New_Font.Name
        .Size = New_Font.Size
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
        .Strikethrough = New_Font.Strikethrough
        .Weight = New_Font.Weight
        .Charset = New_Font.Charset
    End With
    PropertyChanged "Font"
    Refresh
End Property

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    On Error Resume Next
    
    HitResult = vbHitResultHit
    
    If UserControl.Enabled Then
        If Ambient.UserMode Then
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
            
            If lHwnd = c_lhWnd Then
                If m_Over = False Then
                    m_Over = True
                    RaiseEvent MouseOver
                End If
            Else
                If m_Over = True Then
                    m_Over = False
                    RaiseEvent MouseOut
                End If
            End If
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    c_lhWnd = UserControl.ContainerHwnd
    UserControl.Font = Ambient.Font
    m_TickLinesColor = vbButtonShadow
    m_TickTextColor = Ambient.ForeColor
    m_TickText = sldBoth
    m_TickLines = sldBoth
    m_TrackBackColor = Ambient.BackColor
    m_TrackForeColorPositive = vbHighlight
    m_TrackForeColorNegative = vbRed
    m_Max = 100
    m_TicksLongFrequency = 10
    m_BiselColor = Ambient.BackColor
    m_KnobColor = Ambient.BackColor
    m_KnobStyle = ksElipse
    m_TrackSize = 10
    m_TickDistance = 12
    m_TicksSize = 5
    m_GradientStyle = True
    m_Value = 0
    mPercent = 100
    m_BiselSize = 5
    m_KnobWidth = 30
    m_KnobHeight = 30
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If m_OrientationHorizontal = True Then
        If KeyCode = vbKeyLeft Then
            Me.Value = Me.Value - 1
        ElseIf KeyCode = vbKeyRight Then
            Me.Value = Me.Value + 1
        End If
    Else
        If KeyCode = vbKeyUp Then
            Me.Value = Me.Value + 1
        ElseIf KeyCode = vbKeyDown Then
            Me.Value = Me.Value - 1
        End If
    End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
     RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_Paint()
    Draw
End Sub

Public Sub Refresh()
    UserControl.Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        UserControl.Enabled = .ReadProperty("Enabled", True)
        UserControl.Font = .ReadProperty("Font", Ambient.Font)
        m_TicksLongFrequency = .ReadProperty("TicksLongFrequency", 10)
        m_TicksSmallHiden = .ReadProperty("TicksSmallHiden", False)
        m_TicksSize = .ReadProperty("TicksSize", 2)
        m_Min = .ReadProperty("Min", 0)
        m_Max = .ReadProperty("Max", 100)
        m_OrientationHorizontal = .ReadProperty("OrientationHorizontal", False)
        m_TickTextColor = .ReadProperty("TickTextColor", Ambient.ForeColor)
        m_TickLinesColor = .ReadProperty("TickLinesColor", vbButtonShadow)
        m_TickDistance = .ReadProperty("TickDistance", 12)
        m_TrackSize = .ReadProperty("TrackSize", 10)
        m_TrackBackColor = .ReadProperty("TrackBackColor", Ambient.BackColor)
        m_TrackForeColorPositive = .ReadProperty("TrackForeColorPositive", vbHighlight)
        m_TrackForeColorNegative = .ReadProperty("TrackForeColorNegative", vbRed)
        m_TickText = .ReadProperty("TickText", sldBoth)
        m_TickLines = .ReadProperty("TickLines", sldBoth)
        m_KnobStyle = .ReadProperty("KnobStyle", ksElipse)
        m_KnobColor = .ReadProperty("KnobColor", Ambient.BackColor)
        m_BiselColor = .ReadProperty("BiselColor", Ambient.BackColor)
        m_BiselSize = .ReadProperty("BiselSize", 5)
        m_GradientStyle = .ReadProperty("GradientStyle", True)
        m_KnobWidth = .ReadProperty("KnobWidth", 30)
        m_KnobHeight = .ReadProperty("KnobHeight", 30)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbArrow)
        UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        m_MousePointerHands = .ReadProperty("MousePointerHands", False)
        On Error Resume Next
        ArrImgKnob = .ReadProperty("ArrImgKnob")
        
        If m_MousePointerHands Then
            If Ambient.UserMode Then
                UserControl.MousePointer = vbCustom
                UserControl.MouseIcon = GetSystemHandCursor
            End If
        End If

        LoadImageFromStream ArrImgKnob, hImgKnob
        
        If hImgKnob And m_OrientationHorizontal = True Then
            GdipImageRotateFlip hImgKnob, Rotate90FlipNone 'Rotate90FlipXY
        End If

        Me.Value = .ReadProperty("Value", 0)
    End With
    
    c_lhWnd = UserControl.ContainerHwnd
End Sub

Private Sub UserControl_Show()
    Me.Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "Font", UserControl.Font, Ambient.Font
        .WriteProperty "TicksLongFrequency", m_TicksLongFrequency, 10
        .WriteProperty "TicksSmallHiden", m_TicksSmallHiden, False
        .WriteProperty "TicksSize", m_TicksSize, 2
        .WriteProperty "Min", m_Min, 0
        .WriteProperty "Max", m_Max, 100
        .WriteProperty "Value", m_Value, 0
        .WriteProperty "OrientationHorizontal", m_OrientationHorizontal, False
        .WriteProperty "TickTextColor", m_TickTextColor, Ambient.ForeColor
        .WriteProperty "TickLinesColor", m_TickLinesColor, vbButtonShadow
        .WriteProperty "TickDistance", m_TickDistance, 12
        .WriteProperty "TickText", m_TickText, sldBoth
        .WriteProperty "TickLines", m_TickLines, sldBoth
        .WriteProperty "TrackSize", m_TrackSize, 10
        .WriteProperty "TrackBackColor", m_TrackBackColor, Ambient.BackColor
        .WriteProperty "TrackForeColorPositive", m_TrackForeColorPositive, vbHighlight
        .WriteProperty "TrackForeColorNegative", m_TrackForeColorNegative, vbRed
        .WriteProperty "KnobColor", m_KnobColor, Ambient.BackColor
        .WriteProperty "BiselColor", m_BiselColor, Ambient.BackColor
        .WriteProperty "BiselSize", m_BiselSize, 5
        .WriteProperty "GradientStyle", m_GradientStyle, True
        .WriteProperty "MousePointer", UserControl.MousePointer, vbArrow
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointerHands", m_MousePointerHands, False
        .WriteProperty "KnobWidth", m_KnobWidth, False
        .WriteProperty "KnobHeight", m_KnobHeight, False
        .WriteProperty "KnobStyle", m_KnobStyle, ksElipse
        .WriteProperty "ArrImgKnob", ArrImgKnob, 0
    End With
End Sub

Private Sub UserControl_Terminate()
    If hCur Then DestroyCursor hCur
    If hImgKnob Then GdipDisposeImage hImgKnob
    Call GdiplusShutdown(GdipToken)
    modIOleInPlaceActiveObject.TerminateIPAO
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


Private Function LoadImageFromStream(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromStream_Error
    
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromStream = True
        End If
    End If

    Set IStream = Nothing
    
LoadImageFromStream_Error:
End Function

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Sub tmrMOUSEOVER_Timer()
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
    
    bIntercept = False
    SendMessage c_lhWnd, WM_MOUSEMOVE, 0&, ByVal pt.X Or pt.Y * &H10000
    
    If bIntercept = False Then
        If m_Over = True Then
            m_Over = False
            RaiseEvent MouseOut
        End If
    End If
    
    
    If PtInRect(RECT, pt.X, pt.Y) = 0 Then
        m_Enter = False
        tmrMOUSEOVER.Interval = 0
        RaiseEvent MouseLeave
    End If
    
End Sub

Public Property Get IsMouseOver() As Boolean
    IsMouseOver = m_Over
End Property

Public Property Get IsMouseEnter() As Boolean
    IsMouseEnter = m_Enter
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If hCur Then SetCursor hCur
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If m_OrientationHorizontal Then
        If X >= RectKnob.Left And X <= RectKnob.Left + RectKnob.Width Then
            PosDif = X - RectKnob.Left - m_KnobWidth * nScale / 2
        Else
            PosDif = 0
            UserControl_MouseMove Button, Shift, X, Y
        End If
    Else
        If Y >= RectKnob.Top And Y <= RectKnob.Top + RectKnob.Height Then
            PosDif = Y - RectKnob.Top - m_KnobHeight * nScale / 2
        Else
            PosDif = 0
            UserControl_MouseMove Button, Shift, X, Y
        End If
    End If
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Debug.Print X, Y
    
    If hCur Then SetCursor hCur
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
    If Button = vbLeftButton Then
        
        Dim H As Long, W As Long

        If m_OrientationHorizontal Then
            W = UserControl.ScaleWidth - MA - MB - 1
            mPercent = (X - PosDif - MA) * 100 / W
        Else
            H = UserControl.ScaleHeight - MA - MB - 1
            mPercent = (Y - PosDif - MA) * 100 / H
        End If

        If mPercent < 0 Then mPercent = 0
        If mPercent > 100 Then mPercent = 100

        If m_OrientationHorizontal Then
            m_Value = CLng((m_Max - m_Min) * mPercent / 100) + m_Min
        Else
            m_Value = CLng((m_Max - m_Min) * (100 - mPercent) / 100) + m_Min
        End If
        
        Refresh
        If PosDif = 0 Then DoEvents ' primero que dibuje para obtener el RectKnob.left y RectKnob.top
        RaiseEvent Change
    End If

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
    Dim Pic As PicBmp, ipic As IPicture, GUID(0 To 3) As Long
    
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
 
    Call OleCreatePictureIndirect(Pic, GUID(0), 1, ipic)
 
    Set GetSystemHandCursor = ipic
    
End Function

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub


'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
  
    OleTranslateColor clrFirst, 0, VarPtr(clrFore(0))
    OleTranslateColor clrSecond, 0, VarPtr(clrBack(0))
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function
