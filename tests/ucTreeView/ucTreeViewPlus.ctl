VERSION 5.00
Begin VB.UserControl ucTreeViewPlus 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Proyecto1.ucScrollbar ucScrollbar2 
      Height          =   135
      Left            =   240
      Top             =   3360
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   238
      Orientation     =   1
      SmoothScrollFactor=   0,15
      BeginProperty ThumbTooltipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.ucScrollbar ucScrollbar1 
      Height          =   3015
      Left            =   4560
      Top             =   120
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   5318
      SmoothScrollFactor=   0,15
      BeginProperty ThumbTooltipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ucTreeViewPlus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "Kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub CreateStreamOnHGlobal Lib "Ole32.dll" (ByRef hGlobal As Any, ByVal fDeleteOnRelease As Long, ByRef ppstm As Any)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ptr() As Any) As Long
Private Declare Function PathIsURL Lib "shlwapi.dll" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function CryptStringToBinaryA Lib "crypt32.dll" (ByVal pszString As String, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, ByVal pcbBinary As Long, ByVal pdwSkip As Long, ByVal pdwFlags As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hdc As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectType Lib "gdi32.dll" (ByVal hgdiobj As Long) As Long
Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef Brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal Brush As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, hGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal hGraphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "gdiplus" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mpath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mpath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mpath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipClosePathFigure Lib "GdiPlus.dll" (ByVal mpath As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByRef mFilename As Long, ByRef mImage As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imageattr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imageattr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imageattr As Long, ByVal ColorAdjust As Long, ByVal EnableFlag As Boolean, ByRef MatrixColor As COLORMATRIX, ByRef MatrixGray As COLORMATRIX, ByVal flags As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal hGraphics As Long, ByVal hImage As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstWidth As Long, ByVal DstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mHeight As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal mImage As Long, ByRef mWidth As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal mHicon As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal mHbm As Long, ByVal mhPal As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" (ByVal Stream As IUnknown, ByRef Image As Long) As Long
Private Declare Function GdipCreateBitmapFromScan0 Lib "GdiPlus.dll" (ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStride As Long, ByVal mPixelFormat As Long, ByVal mScan0 As Long, ByRef mBitmap As Long) As Long
Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal Image As Long, hGraphics As Long) As Long
Private Declare Function GdipSetClipPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mpath As Long, ByVal mCombineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal InterpolationMode As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "gdiplus" (ByVal graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipAddPathRectangleI Lib "GdiPlus.dll" (ByVal mpath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipAddPathPolygon Lib "GdiPlus.dll" (ByVal mpath As Long, ByRef mPoints As POINTF, ByVal mCount As Long) As Long
Private Declare Function GdipGetPathWorldBoundsI Lib "GdiPlus.dll" (ByVal Path As Long, ByRef bounds As RECTL, ByVal matrix As Long, ByVal pen As Long) As Long
Private Declare Function GdipCreateMatrix Lib "gdiplus" (matrix As Long) As Long
Private Declare Function GdipTranslateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal offsetX As Single, ByVal offsetY As Single, ByVal order As MatrixOrder) As Long
Private Declare Function GdipRotateMatrix Lib "gdiplus" (ByVal matrix As Long, ByVal Angle As Single, ByVal order As MatrixOrder) As Long
Private Declare Function GdipTransformPath Lib "gdiplus" (ByVal Path As Long, ByVal matrix As Long) As Long
Private Declare Function GdipDrawLine Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Single, ByVal mY1 As Single, ByVal mX2 As Single, ByVal mY2 As Single) As Long
Private Declare Function GdipSetPenDashStyle Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashStyle As DashStyle) As Long
Private Declare Function GdipDrawEllipse Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipFillEllipse Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Single, ByVal mY As Single, ByVal mWidth As Single, ByVal mHeight As Single) As Long
Private Declare Function GdipGetImagePixelFormat Lib "gdiplus" (ByVal Image As Long, ByRef PixelFormat As Long) As Long
Private Declare Function GdipGetImageDimension Lib "gdiplus" (ByVal Image As Long, ByRef Width As Single, ByRef Height As Single) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" (ByVal BITMAP As Long, ByRef Rect As Rect, ByVal flags As ImageLockMode, ByVal PixelFormat As Long, ByRef LockedBitmapData As BitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" (ByVal BITMAP As Long, ByRef LockedBitmapData As BitmapData) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)

Private Type PicBmp
    Size As Long
    type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Private Enum DashStyle
    DashStyleSolid = &H0
    DashStyleDash = &H1
    DashStyleDot = &H2
    DashStyleDashDot = &H3
    DashStyleDashDotDot = &H4
    DashStyleCustom = &H5
End Enum

Private Type GDIPlusStartupInput
    GdiPlusVersion                      As Long
    DebugEventCallback                  As Long
    SuppressBackgroundThread            As Long
    SuppressExternalCodecs              As Long
End Type

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type POINTF
    x As Single
    Y As Single
End Type

Private Type COLORMATRIX
    M(0 To 4, 0 To 4)           As Single
End Type

Private Enum MatrixOrder
    MatrixOrderPrepend = 0
    MatrixOrderAppend = 1
End Enum

Private Type ARGB
    Blue            As Byte
    Green           As Byte
    Red             As Byte
    Alpha           As Byte
End Type

Private Type ICONINFO
    fIcon           As Long
    xHotspot        As Long
    yHotspot        As Long
    hbmMask         As Long
    hbmColor        As Long
End Type

Private Type BitmapData
    Width           As Long
    Height          As Long
    Stride          As Long
    PixelFormat     As Long
    Scan0           As Long
    Reserved        As Long
End Type

Private Enum ImageLockMode
    ImageLockModeRead = &H1
    ImageLockModeWrite = &H2
    ImageLockModeUserInputBuf = &H4
End Enum

Private Const SmoothingModeAntiAlias    As Long = 4
Private Const UnitPixel                 As Long = 2&
Private Const PixelFormat32bppPARGB     As Long = &HE200B
Private Const CombineModeIntersect      As Long = &H1
Private Const PixelFormat32bppARGB      As Long = &H26200A
Private Const PixelFormat32bppRGB       As Long = &H22009

Private Const IDC_HAND                  As Long = 32649

Private Const DT_SINGLELINE     As Long = &H20
Private Const DT_LEFT           As Long = &H0
Private Const DT_VCENTER        As Long = &H4
Private Const DT_BOTTOM As Long = &H8
Private Const DT_TOP As Long = &H0
Private Const DT_WORD_ELLIPSIS  As Long = &H40000
Private Const DT_CENTER         As Long = &H1
Private Const CLR_NONE          As Long = &HFFFFFFFF
Private Const DT_RIGHT          As Long = &H2
Private Const DT_CALCRECT       As Long = &H400

Private Const WM_GETTEXTLENGTH As Long = &HE
Private Const WM_GETTEXT As Long = &HD
Private Const WM_SETTEXT As Long = &HC
Private Const EM_SETSEL As Long = &HB1

Public Event OLECompleteDrag(Effect As Long)
Public Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
Public Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Public Event OLESetData(Data As DataObject, DataFormat As Integer)
Public Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Public Event NodeMouseEnter(Node As Node)
Public Event NodeMouseLeave(Node As Node)
Public Event NodeClick(Node As Node)
Public Event NodeDblClick(Node As Node)
Public Event NodeCheck(Node As Node)
Public Event NodeExpand(Node As Node)
Public Event NodeCollapse(Node As Node)
Public Event SelNodeChange(Node As Node)
Public Event NodePostPaint(Node As Node, ByVal hdc As Long, ByVal hGraphics As Long)
Public Event AfterLabelEdit(Node As Node, Cancel As Boolean, NewString As String)
Public Event BeforeLabelEdit(Node As Node, ByRef bCancel As Boolean)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseEnter()
Public Event MouseLeave()
Public Event ScrollChange(ByVal Value As Long, ByVal Max As Long)

Public Enum tvPlusOLEDropModeConstants
    OLEDropModeNone = vbOLEDropNone
    OLEDropModeManual = vbOLEDropManual
End Enum

Public Enum ePlusMinusStyle
    TVP_Arrow = 0
    TVP_PlusMinus = 1
End Enum

Private m_MousePointerHands As Boolean
Private hCur As Long
Private cNodes As New Node
Private NSpacing As Long
Private lMargin As Long
Private MNHeight As Long
Private m_CheckBoxes As Boolean
Private TreeWidth As Long
Private TreeHeight As Long
Private Rect As Rect
Private lToggleIconSize As Long
Private cImageList As Collection
Private ImgLstWidth As Long
Private ImgLstHeight As Long
Private m_ImagesRadius As Long
Private DpiFactor As Double
Private GdipToken As Long
Private CheckSize As Long
Private isSegoeFontInstaled As Boolean
Private m_SegoeFont As StdFont
Private m_Wingdings2 As StdFont
Private m_Wingdings3 As StdFont
Private mhGraphics As Long
Private m_HotItem As Node
Private m_SelItem As Node
Private bExit As Boolean
Private m_Redraw As Boolean
Private m_ContainerBkColor As Long
Private m_FullRowSelect As Boolean
Private m_BorderRadius As Long
Private m_BackColor As Long
Private m_BorderWidth As Long
Private m_SelectionColor As Long
Private m_Font As StdFont
Private m_SubTextFont As StdFont
Private m_SubTextColor As Long
Private m_BorderColor As Long
Private m_TextColor As Long
Private m_ShowTreeLines As Boolean
Private m_PlusMinusVisible As Boolean
Private m_PlusMinusColor As OLE_COLOR
Private m_PlusMinusHotColor As OLE_COLOR
Private bAlternate As Boolean
Private m_PlusMinusStyle As ePlusMinusStyle
Private m_TreeLinesColor As OLE_COLOR
Private m_AlternateNodeColors As Boolean
Private m_CheckBoxColor As OLE_COLOR
Private m_NodeSpacing As Long
Private m_MinNodeHeight As Long
Private m_Margin As Long
Private m_ToggleIconSize As Long
Private m_ClientWidth As Long
Private m_ClientHeight As Long
Private BW As Long
Private m_HotToggle As Node
Private m_AllowEdit As Boolean
Private m_PlusMinusAlignRight As Boolean
Private m_MarginBottom As Long
Private lMBottom As Long
Private searchText As String
Private lastKeyPressTime As Single
Private searchResetTime As Single


Public Property Get ClientWidth() As Long
    ClientWidth = m_ClientWidth
End Property

Public Property Get ClientHeight() As Long
    ClientHeight = m_ClientHeight
End Property

Public Property Get HitTest(ByVal x As Single, ByVal Y As Single) As Node
    Dim oHI As Node
    Set oHI = m_HotItem
    Set m_HotItem = Nothing

    Rect.Top = lMargin - ucScrollbar1.Value
    RecurHotItem x, Y, cNodes, 0
    Set HitTest = m_HotItem
End Property

Public Property Get GetHotNode() As Node
    Set GetHotNode = m_HotItem
End Property

Public Property Let AllowEdit(ByVal New_Value As Boolean)
    m_AllowEdit = New_Value
    PropertyChanged "AllowEdit"
End Property

Public Property Get AllowEdit() As Boolean
    AllowEdit = m_AllowEdit
End Property

Public Property Get ToggleIconSize() As Long
     ToggleIconSize = m_ToggleIconSize
End Property

Public Property Let ToggleIconSize(ByVal New_Value As Long)
    m_ToggleIconSize = New_Value
    Call PropertyChanged("ToggleIconSize")
    Draw
End Property

Public Property Get PlusMinusAlignRight() As Long
     PlusMinusAlignRight = m_PlusMinusAlignRight
End Property

Public Property Let PlusMinusAlignRight(ByVal New_Value As Long)
    m_PlusMinusAlignRight = New_Value
    Call PropertyChanged("PlusMinusAlignRight")
    Draw
End Property

Public Property Get MarginBottom() As Long
     MarginBottom = m_MarginBottom
End Property

Public Property Let MarginBottom(ByVal New_Value As Long)
    m_MarginBottom = New_Value
    Call PropertyChanged("MarginBottom")
    Draw
End Property

Public Property Get Margin() As Long
     Margin = m_Margin
End Property

Public Property Let Margin(ByVal New_Value As Long)
    m_Margin = New_Value
    Call PropertyChanged("Margin")
    Draw
End Property

Public Property Get MinNodeHeight() As Long
     MinNodeHeight = m_MinNodeHeight
End Property

Public Property Let MinNodeHeight(ByVal New_Value As Long)
    m_MinNodeHeight = New_Value
    Call PropertyChanged("MinNodeHeight")
    Draw
End Property

Public Property Get NodeSpacing() As Long
     NodeSpacing = m_NodeSpacing
End Property

Public Property Let NodeSpacing(ByVal New_Value As Long)
    m_NodeSpacing = New_Value
    Call PropertyChanged("NodeSpacing")
    Draw
End Property

Public Property Get CheckBoxColor() As OLE_COLOR
    CheckBoxColor = m_CheckBoxColor
End Property

Public Property Let CheckBoxColor(ByVal New_Value As OLE_COLOR)
    m_CheckBoxColor = New_Value
    Call PropertyChanged("CheckBoxColor")
    Draw
End Property

Public Property Get PlusMinusColor() As OLE_COLOR
    PlusMinusColor = m_PlusMinusColor
End Property

Public Property Let PlusMinusColor(ByVal New_Value As OLE_COLOR)
    m_PlusMinusColor = New_Value
    Call PropertyChanged("PlusMinusColor")
    Draw
End Property

Public Property Get PlusMinusHotColor() As OLE_COLOR
    PlusMinusHotColor = m_PlusMinusHotColor
End Property

Public Property Let PlusMinusHotColor(ByVal New_Value As OLE_COLOR)
    m_PlusMinusHotColor = New_Value
    Call PropertyChanged("PlusMinusHotColor")
    Draw
End Property

Public Property Get AlternateNodeColors() As Boolean
    AlternateNodeColors = m_AlternateNodeColors
End Property

Public Property Let AlternateNodeColors(ByVal New_Value As Boolean)
    m_AlternateNodeColors = New_Value
    Call PropertyChanged("AlternateNodeColors")
    Draw
End Property

Public Property Get TreeLinesColor() As OLE_COLOR
    TreeLinesColor = m_TreeLinesColor
End Property

Public Property Let TreeLinesColor(ByVal New_Value As OLE_COLOR)
    m_TreeLinesColor = New_Value
    Call PropertyChanged("TreeLinesColor")
    Draw
End Property

Public Property Let PlusMinusStyle(ByVal New_Value As ePlusMinusStyle)
    m_PlusMinusStyle = New_Value
    Call PropertyChanged("PlusMinusStyle")
    Draw
End Property

Public Property Get PlusMinusStyle() As ePlusMinusStyle
    PlusMinusStyle = m_PlusMinusStyle
End Property

Public Property Let TextColor(ByVal New_Value As OLE_COLOR)
    m_TextColor = New_Value
    Call PropertyChanged("TextColor")
    Draw
End Property

Public Property Get TextColor() As OLE_COLOR
    TextColor = m_TextColor
End Property

Public Property Let SubTextColor(ByVal New_Value As OLE_COLOR)
    m_SubTextColor = New_Value
    Call PropertyChanged("SubTextColor")
    Draw
End Property

Public Property Get SubTextColor() As OLE_COLOR
    SubTextColor = m_SubTextColor
End Property


Public Property Get SubTextFont() As StdFont
    Set SubTextFont = m_SubTextFont
End Property

Public Property Set SubTextFont(ByVal New_Value As StdFont)
    Set m_SubTextFont = New_Value
    Draw
End Property

Public Property Let SubTextFont(ByVal New_Value As StdFont)
    Set m_SubTextFont = New_Value
    PropertyChanged "SubTextFont"
    Draw
End Property

Public Property Let ContainerBkColor(ByVal New_Value As OLE_COLOR)
    m_ContainerBkColor = New_Value
    UserControl.BackColor = New_Value
    Call PropertyChanged("ContainerBkColor")
    Draw
End Property

Public Property Get ContainerBkColor() As OLE_COLOR
    ContainerBkColor = m_ContainerBkColor
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_Value As OLE_COLOR)
    m_BackColor = New_Value
    ucScrollbar1.StyleBackColor = New_Value
    ucScrollbar2.StyleBackColor = New_Value
    PropertyChanged "BackColor"
    Draw
End Property

Public Property Let CheckBoxes(ByVal New_Value As Boolean)
    m_CheckBoxes = New_Value
    PropertyChanged "CheckBoxes"
    Draw
End Property

Public Property Get CheckBoxes() As Boolean
    CheckBoxes = m_CheckBoxes
End Property

Public Property Let ShowTreeLines(ByVal New_Value As Boolean)
    m_ShowTreeLines = New_Value
    PropertyChanged "ShowTreeLines"
    Draw
End Property

Public Property Get ShowTreeLines() As Boolean
    ShowTreeLines = m_ShowTreeLines
End Property

Public Property Let PlusMinusVisible(ByVal New_Value As Boolean)
    m_PlusMinusVisible = New_Value
    PropertyChanged "PlusMinusVisible"
    Draw
End Property

Public Property Get PlusMinusVisible() As Boolean
    PlusMinusVisible = m_PlusMinusVisible
End Property

Public Property Get Redraw() As Boolean
    Redraw = m_Redraw
End Property

Public Property Let Redraw(ByVal New_Value As Boolean)
    If New_Value = True And m_Redraw <> New_Value Then
        m_Redraw = New_Value
        Draw
    Else
        m_Redraw = New_Value
    End If
End Property

Public Property Get FullRowSelect() As Boolean
    FullRowSelect = m_FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_Value As Boolean)
    m_FullRowSelect = New_Value
    PropertyChanged "FullRowSelect"
    Draw
End Property

Public Property Get SelectionColor() As OLE_COLOR
    SelectionColor = m_SelectionColor
End Property

Public Property Let SelectionColor(ByVal New_Value As OLE_COLOR)
    m_SelectionColor = New_Value
    PropertyChanged "SelectionColor"
    Draw
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Value As StdFont)
    Set m_Font = New_Value
    Draw
End Property

Public Property Let Font(ByVal New_Value As StdFont)
    Set m_Font = New_Value
    PropertyChanged "Font"
    Draw
End Property

Public Property Get BorderRadius() As Long
    BorderRadius = m_BorderRadius
End Property

Public Property Let BorderRadius(ByVal New_Value As Long)
    m_BorderRadius = New_Value
    PropertyChanged "BorderRadius"
    Update
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal New_Value As OLE_COLOR)
    m_BorderColor = New_Value
    PropertyChanged "BorderColor"
    Draw
End Property

Public Property Get BorderWidth() As Long
    BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_Value As Long)
    m_BorderWidth = New_Value
    PropertyChanged "BorderWidth"
    Update
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Value As Boolean)
    UserControl.Enabled = New_Value
    PropertyChanged "Enabled"
End Property

Public Property Get DragIcon() As IPictureDisp
Set DragIcon = Extender.DragIcon
End Property

Public Property Let DragIcon(ByVal Value As IPictureDisp)
    Extender.DragIcon = Value
End Property

Public Property Set DragIcon(ByVal Value As IPictureDisp)
    Set Extender.DragIcon = Value
End Property

Public Property Get DragMode() As Integer
    DragMode = Extender.DragMode
End Property

Public Property Let DragMode(ByVal Value As Integer)
    Extender.DragMode = Value
End Property

Public Property Get OLEDropMode() As tvPlusOLEDropModeConstants
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal Value As tvPlusOLEDropModeConstants)
    Select Case Value
        Case OLEDropModeNone, OLEDropModeManual
            UserControl.OLEDropMode = Value
        Case Else
            Err.Raise 380
    End Select
    UserControl.PropertyChanged "OLEDropMode"
End Property

Public Sub Drag(Optional ByRef Action As Variant)
    If IsMissing(Action) Then Extender.Drag Else Extender.Drag Action
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
            'UserControl.MousePointer = vbCustom
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

Public Property Get ScrollBarStyle() As sbStyleCts
    ScrollBarStyle = ucScrollbar1.Style
End Property

Public Property Let ScrollBarStyle(ByVal New_Value As sbStyleCts)
    ucScrollbar1.Style = New_Value
    ucScrollbar2.Style = New_Value
    If New_Value = sGoogle Then
        ucScrollbar1.Width = 8 * DpiFactor
        ucScrollbar2.Height = 8 * DpiFactor
        ucScrollbar1.ShowButtons = False
        ucScrollbar2.ShowButtons = False
    Else
        ucScrollbar1.Width = 14 * DpiFactor
        ucScrollbar2.Height = 14 * DpiFactor
        ucScrollbar1.ShowButtons = True
        ucScrollbar2.ShowButtons = True
    End If
    
    PropertyChanged "ScrollBarStyle"
    
    If m_Redraw Then Draw
End Property

Public Property Get EditHwnd() As Long
    EditHwnd = Text1.hwnd
End Property

Public Sub CancelEdit()
    Text1.Text = vbNullString
    Text1.Visible = False
End Sub

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Public Property Get Nodes() As Node
    Set Nodes = cNodes
End Property

Public Sub StartEditNode(N As Node)
    Text1.Move N.Left, N.Top, N.Right - N.Left, N.Bottom - N.Top
    Text1.Text = N.Text
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
    Text1.Visible = True
    Text1.SetFocus
End Sub


Public Sub OLEDrag()
    UserControl.OLEDrag
End Sub

Public Sub Refresh()
    Draw
End Sub

'*3
Public Sub Update()
    TreeHeight = lMargin * 2
    TreeWidth = 0
    
    If m_Font Is Nothing Then Exit Sub
    RecurCalcTreeWidth cNodes, 0, True

    If Ambient.UserMode Then
        If TreeHeight > UserControl.ScaleHeight Then
            With ucScrollbar1
                .Max = TreeHeight - UserControl.ScaleHeight + BW * 2
                .TrackMouseWheelOnHwnd UserControl.hwnd
                .SmallChange = 1
                .WheelChange = MNHeight
                .LargeChange = UserControl.ScaleHeight / 4
                .Visible = True
            End With
        Else
            With ucScrollbar1
                .TrackMouseWheelOnHwndStop
                .Visible = False
            End With
        End If
    End If
    
    m_ClientWidth = UserControl.ScaleWidth - IIf(ucScrollbar1.Visible, ucScrollbar1.Width, 0) '- BW * 2
    
    
    If TreeWidth > m_ClientWidth Then
        With ucScrollbar2
            .Visible = True
            .Max = TreeWidth - m_ClientWidth + lMargin + BW * 2
            .LargeChange = UserControl.ScaleWidth
        End With
    Else
        With ucScrollbar2
            .Visible = False
            .Value = 0
        End With
    End If
    
    m_ClientWidth = UserControl.ScaleWidth - IIf(ucScrollbar1.Visible, ucScrollbar1.Width, 0) '- BW * 2
    
    If ucScrollbar2.Visible Then
        ucScrollbar1.Max = ucScrollbar1.Max + ucScrollbar2.Height
        m_ClientHeight = UserControl.ScaleHeight - ucScrollbar2.Height '- BW * 2
    Else
        m_ClientHeight = UserControl.ScaleHeight '- BW * 2
    End If
    
    If ucScrollbar1.Visible = False Then ucScrollbar1.Value = 0
    If ucScrollbar2.Visible = False Then ucScrollbar2.Value = 0
    
    AjustScrolls
    
    Draw
End Sub

Public Property Let SelNode(TheNode As Node)
    Set m_SelItem = TheNode
    Rect.Top = lMargin
    EnsureVisible cNodes
    RaiseEvent SelNodeChange(m_SelItem)
End Property

Public Property Get SelNode() As Node
    Set SelNode = m_SelItem
End Property

Public Sub EnsureVisible(TheNode As Node)
  Dim N As Node
    Dim TIS As Long
    Dim TvHeight As Long
    If TheNode Is Nothing Then Exit Sub
    TIS = lToggleIconSize '* DpiFactor
    TvHeight = UserControl.ScaleHeight - IIf(ucScrollbar2.Visible, ucScrollbar2.Height, 0)
    For Each N In TheNode.Childs

        If N Is m_SelItem Then
            If Rect.Top < ucScrollbar1.Value Then
                ucScrollbar1.Value = Rect.Top - lMargin
                Exit Sub
            ElseIf Rect.Top + N.Height > ucScrollbar1.Value + TvHeight Then
                ucScrollbar1.Value = Rect.Top - TvHeight + N.Height + lMargin + m_MarginBottom * DpiFactor
                Exit Sub
            Else
                Draw
                Exit Sub
            End If
        End If

        Rect.Top = Rect.Top + N.Height + m_MarginBottom * DpiFactor
        
        If N.Expanded Then
            EnsureVisible N
        End If

    Next
End Sub

Public Property Get ImagesRadius() As Long
    ImagesRadius = m_ImagesRadius
End Property

Public Property Let ImagesRadius(ByVal New_Value As Long)
    m_ImagesRadius = New_Value
    PropertyChanged "ImagesRadius"
    Draw
End Property

Public Function ImageListInitialize(Optional ByVal ImageWidth As Integer = 16, Optional ByVal ImageHeight As Integer = 16) As Boolean
    ImgLstWidth = ImageWidth * DpiFactor
    ImgLstHeight = ImageHeight * DpiFactor
    ImageListClear
End Function

Public Property Get ImgListWidth() As Long
   ImgListWidth = ImgLstWidth
End Property

Public Property Get ImgListHeight() As Long
   ImgListHeight = ImgLstHeight
End Property

Public Function ImageListAddImage(SrcImg As Variant, Optional key As Variant) As Long
    Dim hImage As Long
    hImage = LoadPictureEx(SrcImg, ImgLstWidth, ImgLstHeight)
    If hImage <> 0& Then
        cImageList.Add hImage, key
        ImageListAddImage = cImageList.Count
    End If
End Function

Public Function ImageListRemoveImage(index As Variant) As Long
    GdipDisposeImage cImageList(index)
    cImageList.Remove index
    ImageListRemoveImage = cImageList.Count
End Function

Public Function ImageListInsertImage(SrcImg As Variant, Optional key As Variant, Optional Before, Optional After) As Long
    Dim hImage As Long
    hImage = LoadPictureEx(SrcImg, ImgLstWidth, ImgLstHeight)
    If hImage <> 0& Then
        cImageList.Add hImage, key, Before, After
        ImageListInsertImage = cImageList.Count
    End If
End Function

Public Function ImageListGetImage(index As Variant) As Long
    ImageListGetImage = cImageList(index)
End Function

Public Sub ImageListClear()
    Dim i As Long
    For i = 1 To cImageList.Count
        GdipDisposeImage cImageList(i)
    Next
    Set cImageList = New Collection
End Sub

Public Property Get ImageListCount() As Long
    ImageListCount = cImageList.Count
End Property

Public Function ImageListDraw(ByVal hdc As Long, ByVal index As Variant, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, Optional PictureColor As Long = CLR_NONE, Optional Disabled As Boolean) As Boolean
    Dim tMatrixColor As COLORMATRIX, tMatrixGray As COLORMATRIX
    Dim hAttributes As Long, hGraphics As Long
    Dim hImage As Long
    
    On Error Resume Next
    hImage = cImageList(index)
    If hImage = 0 Then
        Debug.Print "THE IMAGE: " & index & " DON'T EXIST"
        Exit Function
    End If
    
    Call GdipCreateImageAttributes(hAttributes)
    
    With tMatrixColor
        If PictureColor <> CLR_NONE Then
            If (PictureColor And &H80000000) Then PictureColor = GetSysColor(PictureColor And &HFF&)
            Dim R As Byte, G As Byte, b As Byte

            b = ((PictureColor \ &H10000) And &HFF)
            G = ((PictureColor \ &H100) And &HFF)
            R = (PictureColor And &HFF)

            .M(0, 0) = R / 255
            .M(1, 0) = G / 255
            .M(2, 0) = b / 255
            .M(0, 4) = R / 255
            .M(1, 4) = G / 255
            .M(2, 4) = b / 255
        Else
            .M(0, 0) = 1
            .M(1, 1) = 1
            .M(2, 2) = 1
        End If
        .M(3, 3) = 1
        .M(4, 4) = 1

        If Disabled Then
            .M(0, 0) = 0.299
            .M(1, 0) = 0.299
            .M(2, 0) = 0.299
            .M(0, 1) = 0.587
            .M(1, 1) = 0.587
            .M(2, 1) = 0.587
            .M(0, 2) = 0.114
            .M(1, 2) = 0.114
            .M(2, 2) = 0.114
        End If
    End With
    
    GdipCreateFromHDC hdc, hGraphics
    Call GdipSetInterpolationMode(hGraphics, 7&)  '//InterpolationModeHighQualityBicubic
    Call GdipSetPixelOffsetMode(hGraphics, 4&)
    GdipSetImageAttributesColorMatrix hAttributes, &H0, True, tMatrixColor, tMatrixGray, &H0
    ImageListDraw = GdipDrawImageRectRectI(hGraphics, hImage, Left, Top, Width, Height, 0, 0, Width, Height, UnitPixel, hAttributes) = 0&
    Call GdipDeleteGraphics(hGraphics)
    Call GdipDisposeImageAttributes(hAttributes)
End Function

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Public Function GetWindowsDPI() As Double
    Const LOGPIXELSX        As Long = 88
    Dim hdc As Long, lPx  As Double
    Static DPI As Double
    If DPI = 0 Then
        hdc = GetDC(0)
        lPx = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
        ReleaseDC 0, hdc
    
        If (lPx = 0) Then
            DPI = 1#
        Else
            DPI = lPx / 96#
        End If
    End If
    
    GetWindowsDPI = DPI
    
End Function

Public Sub ExpandAllNodes(Optional ByVal Node As Node)

    Dim N As Node
    
    If Node Is Nothing Then
        Set Node = cNodes
    End If

    Node.Expanded = True
    For Each N In Node.Childs
        ExpandAllNodes N
    Next
    
End Sub


Public Sub CollapseAllNodes(Optional ByVal Node As Node)

    Dim N As Node
    
    If Node Is Nothing Then
        Set Node = cNodes
    End If

    Node.Expanded = False
    For Each N In Node.Childs
        ExpandAllNodes N
    Next
    
End Sub


Public Function LoadFromXML(ByVal sPath As String) As Boolean

    Dim xmlDoc As Object
    Dim childNode As Object, xmlNode As Object, xNode As Object
    Dim i As Long

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    
    xmlDoc.Load sPath
    
    If xmlDoc.parseError.errorCode <> 0 Then
        MsgBox "Error loading XML file: " & xmlDoc.parseError.reason, vbCritical, "XML Load Error"
        Exit Function
    End If
    
    cNodes.Clear

    Set xmlNode = xmlDoc.documentElement
    For i = 0 To xmlNode.childNodes.Length - 1
        Set childNode = xmlNode.childNodes.Item(i)
        Select Case childNode.nodename
        
            Case "images"
                CargarImagenes childNode
            Case "nodes"
                For Each xNode In childNode.childNodes
                    RecorrerNodos xNode, cNodes
                Next
        End Select
   
    Next i
    
    Call Update
End Function

'*--

Private Sub CargarImagenes(imagesNode As Object)
    'On Error Resume Next
    Dim imageNode As Object
    Dim xmlAttr As Object
    Dim ImgWidth As Long, ImgHeight As Long
    
    ImgWidth = 16
    ImgHeight = 16
    
    If Not imagesNode Is Nothing Then

        If Not imagesNode.Attributes Is Nothing Then
            With Me
                For Each xmlAttr In imagesNode.Attributes
                    Select Case xmlAttr.nodename
                        Case "Width": ImgWidth = xmlAttr.Text
                        Case "Height": ImgHeight = xmlAttr.Text
                        Case "Radius": m_ImagesRadius = xmlAttr.Text
                    End Select
                Next xmlAttr
                ImageListInitialize ImgWidth, ImgHeight
            End With
        End If
    
        For Each imageNode In imagesNode.childNodes
 
            If imageNode.nodename = "image" Then
                Dim key As Variant
                If Not imageNode.Attributes.getNamedItem("key") Is Nothing Then
                    key = imageNode.Attributes.getNamedItem("key").Text
                    ImageListAddImage imageNode.Attributes.getNamedItem("src").Text, key
                Else
                    ImageListAddImage imageNode.Attributes.getNamedItem("src").Text
                End If
            End If
        Next imageNode
    End If
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim NewString As String
        Dim bCancel As Boolean
        
        NewString = Text1.Text
    
        RaiseEvent AfterLabelEdit(m_SelItem, bCancel, NewString)
        
        If Not bCancel Then
            m_SelItem.Text = NewString
            Text1.Visible = False
            m_SelItem.Height = 0
            Update
            KeyAscii = 0
            Text1.Text = vbNullString
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        Text1.Visible = False
        KeyAscii = 0
    End If
End Sub


Private Sub Text1_LostFocus()
    Text1.Visible = False
End Sub

Private Sub ucScrollbar1_ContainerMouseEnter()
    RaiseEvent MouseEnter
End Sub

Private Sub ucScrollbar1_ContainerMouseLeave()
    If Not m_HotItem Is Nothing Then
        RaiseEvent NodeMouseLeave(m_HotItem)
        Set m_HotItem = Nothing
        Draw
    End If
    
    RaiseEvent MouseLeave
End Sub

Private Sub UserControl_Click()
    Dim pt As POINTAPI
    Dim x As Single, Y As Single
    
    GetCursorPos pt
    ScreenToClient UserControl.hwnd, pt
    x = pt.x
    Y = pt.Y

    Rect.Top = lMargin - ucScrollbar1.Value
    RecurCalcNodeCheckBox x, Y, cNodes, 0, True, False
End Sub

Private Sub UserControl_DblClick()
    Dim pt As POINTAPI
    Dim x As Single, Y As Single
    
    GetCursorPos pt
    ScreenToClient UserControl.hwnd, pt
    x = pt.x
    Y = pt.Y
    
    Rect.Top = lMargin - ucScrollbar1.Value
    RecurCalcNodeCheckBox x, Y, cNodes, 0, False, True
End Sub

Private Sub UserControl_InitProperties()
    Set m_Font = Ambient.Font
    Set m_SubTextFont = Ambient.Font
    m_SubTextFont.Size = m_SubTextFont.Size - 1
    m_SubTextColor = &HCCCCCC
    m_PlusMinusVisible = True
    m_ContainerBkColor = Ambient.BackColor
    UserControl.BackColor = m_ContainerBkColor
    m_BackColor = vbWindowBackground
    m_BorderColor = vbActiveBorder
    m_SelectionColor = &HFFE8CC
    m_NodeSpacing = 16
    m_ToggleIconSize = 10
    m_MinNodeHeight = 32
    m_Margin = 5
    m_BorderRadius = 5
    m_TreeLinesColor = &HCCCCCC
    m_PlusMinusColor = vbBlack
    m_PlusMinusHotColor = &HFF6363
    m_CheckBoxColor = &HFF6363
    m_BorderWidth = 1
    m_FullRowSelect = True
 
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    ' Calcular el tiempo desde la última pulsación de tecla
    If Timer - lastKeyPressTime > searchResetTime Then
        ' Si ha pasado más tiempo del límite, reiniciar la búsqueda
        searchText = ""
    End If

    ' Acumular el texto de búsqueda
    searchText = searchText & Chr(KeyAscii)

    ' Actualizar el tiempo de la última pulsación
    lastKeyPressTime = Timer

    ' Buscar el siguiente ítem que coincida con el texto, desde la selección actual
    If Not FindNextNode(cNodes, False) Then
        ' Si no se encuentra ningún nodo en la primera fase, buscar desde el principio
        Call FindNextNode(cNodes, True)
    End If
    Rect.Top = lMargin
    EnsureVisible cNodes
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Function FindNextNode(TheNode As Node, bNext As Boolean) As Boolean
    Dim N As Node

    For Each N In TheNode.Childs
        If bNext Then
            If Not N.Disabled Then
                If LCase(N.Text) Like LCase(searchText) & "*" Then
                    Set m_SelItem = N
                    RaiseEvent SelNodeChange(m_SelItem)
                    Draw

                    ' Nodo encontrado, regresar verdadero
                    FindNextNode = True
                    Exit Function
                End If
            End If
        End If

        ' Marcar cuando llegamos al nodo seleccionado, solo en la primera fase
        If N Is m_SelItem Then
            bNext = True
        End If

        ' Si el nodo está expandido, buscar en sus hijos
        If N.Expanded Then
            If FindNextNode(N, bNext) Then
                FindNextNode = True
                Exit Function
            End If
        End If
    Next
    
    ' Si no se encuentra ningún nodo, regresar falso
    
End Function


Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, UserControl.ScaleX(x, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition))
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, UserControl.ScaleX(x, vbPixels, vbContainerPosition), UserControl.ScaleY(Y, vbPixels, vbContainerPosition), State)
End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub ucScrollbar1_Change()
    Text1.Visible = False
    Draw

    RaiseEvent ScrollChange(ucScrollbar1.Value, ucScrollbar1.Max)
End Sub

Private Sub ucScrollbar1_Scroll()
    Draw
End Sub

Private Sub ucScrollbar2_Change()
    Text1.Visible = False
    Draw
End Sub

Private Sub ucScrollbar2_Scroll()
    Draw
End Sub


Private Sub UserControl_Initialize()
    DpiFactor = GetWindowsDPI

    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
  
    Set cImageList = New Collection
    m_Redraw = True

    searchText = ""

    searchResetTime = 1  ' 1 segundo para restablecer la búsqueda
    lastKeyPressTime = 0
    
    Update
End Sub



'*1
Private Sub Draw()
    Dim hPath As Long, hBrush As Long, hPen As Long
    
    BW = m_BorderWidth * DpiFactor
    NSpacing = m_NodeSpacing * DpiFactor
    MNHeight = m_MinNodeHeight * DpiFactor
    lToggleIconSize = m_ToggleIconSize * DpiFactor
    CheckSize = 14 * DpiFactor
    lMBottom = m_MarginBottom * DpiFactor
    lMargin = m_Margin * DpiFactor
    If m_Redraw = False Then Exit Sub
    Rect.Top = BW + lMargin - ucScrollbar1.Value
    Cls
    GdipCreateFromHDC hdc, mhGraphics
    GdipSetSmoothingMode mhGraphics, SmoothingModeAntiAlias
    
    hPath = CreateRoundPath(BW \ 2, BW \ 2, UserControl.ScaleWidth - BW, UserControl.ScaleHeight - BW, m_BorderRadius * DpiFactor)
    
    GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
    GdipFillPath mhGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    
    bAlternate = False
    
    RecursiveDraw cNodes, 0
    
    If BW > 0 Then
        GdipCreatePen1 RGBtoARGB(m_BorderColor, 100), BW, UnitPixel, hPen
        GdipDrawPath mhGraphics, hPen, hPath
        GdipDeletePen hPen
    End If
    
    GdipDeletePath hPath
    Call GdipDeleteGraphics(mhGraphics)
    UserControl.Refresh
End Sub

Private Sub RecursiveDraw(TheNode As Node, Ident As Long)
    Dim N As Node
    Dim H As Long
    Dim TIS As Long
    Dim oParent As Node
    Dim TempIdent As Long
    Dim lLeft As Long, lWidth As Long
    Dim lTop As Long, lHeight As Long
    Dim hPen As Long
    Dim hBrush As Long, hPath As Long
    Dim bkColor  As Long
    Dim SelColor As Long
    
    TIS = lToggleIconSize

    For Each N In TheNode.Childs
        H = IIf(N.Height < MNHeight, MNHeight, N.Height)
        N.Right = -1: N.Left = -1
        N.Bottom = -1: N.Top = -1
        
        If Rect.Top + H > 0 Then

        '---------Alternar colores en nodos-------------
        If m_AlternateNodeColors Then
            If bAlternate = True Then
                If IsDarkColor(m_BackColor) Then
                    bkColor = ShiftColor(m_BackColor, vbWhite, 200)
                Else
                    bkColor = ShiftColor(m_BackColor, vbBlack, 250)
                End If
                GdipCreateSolidFill RGBtoARGB(bkColor, 100), hBrush
                GdipFillRectangleI mhGraphics, hBrush, 0, Rect.Top, UserControl.ScaleWidth, N.Height
                GdipDeleteBrush hBrush
                bAlternate = False
            Else
                bAlternate = True
                bkColor = m_BackColor
            End If
        Else
            bkColor = m_BackColor
        End If
        '------------------------------------------

        Rect.Left = -ucScrollbar2.Value + lMargin + Ident * NSpacing + BW + NSpacing / 2
        If m_ShowTreeLines Or m_PlusMinusVisible And Not m_PlusMinusAlignRight Then
            Rect.Left = Rect.Left + TIS + lMargin
        End If
        Rect.Bottom = Rect.Top + H - DpiFactor
        Rect.Right = Rect.Left + lMargin
        If N.TextWidth > 0 Then Rect.Right = Rect.Right + N.TextWidth + lMargin
        If (m_CheckBoxes And Not N.HideCheck) Then
            Rect.Right = Rect.Right + CheckSize + lMargin '* 2
        End If
        If Not VarType(N.Image) = vbError Then
            If N.Image > 0 Then Rect.Right = Rect.Right + ImgLstWidth + lMargin '* 2
        End If
        
        If m_FullRowSelect Then
            lLeft = -ucScrollbar2.Value + BW + lMargin
            lWidth = m_ClientWidth - lMargin * 2 - BW * 2.5
            If lWidth < Rect.Right Or ucScrollbar2.Value > 0 Then
                lWidth = Rect.Right - lLeft
            End If
            If lLeft + lWidth < m_ClientWidth - lMargin * 2 - BW * 2.5 Then
                lWidth = m_ClientWidth + ucScrollbar2.Value - lMargin * 2 - BW * 2.5
            End If
            
            lTop = Rect.Top
            lHeight = Rect.Bottom - Rect.Top
            
        Else
            lLeft = Rect.Left
            lTop = Rect.Top
            lWidth = Rect.Right - Rect.Left
            lHeight = Rect.Bottom - Rect.Top
        End If


        If (N.BackColor <> -1) Or (N.BorderWidth > 0) Or (N.Selected) Or (m_HotItem Is N) Or (m_SelItem Is N) Then
            hPath = CreateRoundPath(lLeft, lTop, lWidth, lHeight, 4 * DpiFactor)
           
            If (N.BackColor <> -1) Then
                If m_FullRowSelect Then bkColor = N.BackColor
                GdipCreateSolidFill RGBtoARGB(N.BackColor, 100), hBrush
                GdipFillPath mhGraphics, hBrush, hPath
                GdipDeleteBrush hBrush
            End If
            
            If (N.BorderWidth > 0) Then
                GdipCreatePen1 RGBtoARGB(N.BorderColor, 100), N.BorderWidth * DpiFactor, UnitPixel, hPen
                GdipDrawPath mhGraphics, hPen, hPath
                GdipDeletePen hPen
            End If
            
            If N.Selected Or (m_SelItem Is N) Then
                If m_FullRowSelect Then bkColor = ShiftColor(bkColor, m_SelectionColor, 50)
                GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 90), hBrush
                GdipFillPath mhGraphics, hBrush, hPath
                GdipDeleteBrush hBrush
                If (m_HotItem Is N) Then
                    GdipCreatePen1 RGBtoARGB(m_SelectionColor, 100), DpiFactor, UnitPixel, hPen
                    GdipDrawPath mhGraphics, hPen, hPath
                    GdipDeletePen hPen
                End If
            ElseIf (m_HotItem Is N) Then
                If m_FullRowSelect Then bkColor = ShiftColor(bkColor, m_SelectionColor, 60)
                GdipCreateSolidFill RGBtoARGB(m_SelectionColor, 70), hBrush
                GdipFillPath mhGraphics, hBrush, hPath
                GdipDeleteBrush hBrush
            End If

            GdipDeletePath hPath
        End If

        '---------------------- LINEAS-----------------
        If m_ShowTreeLines Then
            GdipCreatePen1 RGBtoARGB(m_TreeLinesColor, 100), DpiFactor, UnitPixel, hPen
            GdipSetPenDashStyle hPen, DashStyleDot

            lLeft = Rect.Left - TIS / 2 - lMargin
            GdipDrawLine mhGraphics, hPen, lLeft, Rect.Top + H / 2, Rect.Left, Rect.Top + H / 2
 
            lLeft = Rect.Left - TIS / 2 - lMargin
            
            If (N Is N.Parent.Childs(1)) And Not (N Is Me.Nodes.Childs(1)) Then
                GdipDrawLine mhGraphics, hPen, lLeft, Rect.Top - lMBottom, lLeft, Rect.Top + H / 2
            End If
            
            If N Is TheNode.Childs(TheNode.Childs.Count) And Not N Is TheNode.Childs(1) Then
                GdipDrawLine mhGraphics, hPen, lLeft, Rect.Top - lMBottom, lLeft, Rect.Top + H / 2
            Else
                If Not N Is TheNode.Childs(1) Then
                    GdipDrawLine mhGraphics, hPen, lLeft, Rect.Top - lMBottom, lLeft, Rect.Top + H + DpiFactor
                Else
                    If Not N Is TheNode.Childs(TheNode.Childs.Count) Then
                        GdipDrawLine mhGraphics, hPen, lLeft, Rect.Top + H / 2 - lMBottom / 2, lLeft, Rect.Top + H
                    End If
                End If
            End If
            
            Set oParent = N.Parent
            TempIdent = Ident

            Do
                TempIdent = TempIdent - 1
      
                If TempIdent > -1 And Not oParent.Parent Is Nothing Then
                    If Not oParent Is oParent.Parent.Childs(oParent.Parent.Childs.Count) Then
                        lLeft = -ucScrollbar2.Value + BW + lMargin + TempIdent * NSpacing + TIS / 2 + NSpacing / 2
                        GdipDrawLine mhGraphics, hPen, lLeft, Rect.Top - lMBottom, lLeft, Rect.Top + H
                    End If
                End If

                Set oParent = oParent.Parent
                If oParent Is Nothing Then Exit Do
            Loop
            GdipDeletePen hPen
        End If
        '----------------------LINEAS-----------------
        
        If m_PlusMinusVisible Then
            If N.Childs.Count > 0 Or N.PlusButton = True Then
                If m_HotToggle Is N Then
                    SelColor = RGBtoARGB(m_PlusMinusHotColor, 100)
                Else
                    SelColor = RGBtoARGB(m_PlusMinusColor, 100)
                End If
                
                If m_PlusMinusAlignRight Then
                    lLeft = m_ClientWidth - TIS - lMargin - NSpacing / 2 - BW
                Else
                    lLeft = Rect.Left - TIS - lMargin
                End If
                
                If m_PlusMinusStyle = TVP_Arrow Then
                    DrawArrow1 mhGraphics, N.Expanded, bkColor, SelColor, lLeft, Rect.Top + H / 2 - TIS / 2, TIS
                Else
                    DrawToggleButton mhGraphics, N.Expanded, SelColor, lLeft, Rect.Top + H / 2 - TIS / 2, TIS
                End If
            End If
        End If
        
        lLeft = Rect.Left
        If (m_CheckBoxes And Not N.HideCheck) Then
            Dim lColor As Long
            
            If N.Disabled Then
                lColor = &HCCCCCC
            Else
                lColor = m_CheckBoxColor
            End If

            lLeft = lLeft + lMargin
            If TheNode.RadioStyle Then
                DrawRadio lColor, lLeft, Rect.Top + H / 2 - CheckSize / 2, CheckSize, CheckSize, N.Checked
            Else
                DrawCheckBox lColor, lLeft, Rect.Top + H / 2 - CheckSize / 2, CheckSize, CheckSize, N.Checked
                'DrawSwitch mhGraphics, lColor, lLeft, Rect.Top + H / 2 - CheckSize / 2, CheckSize, CheckSize, N.Checked
            End If
            lLeft = lLeft + CheckSize '+ lMargin
        End If
        
        If Not VarType(N.Image) = vbError Then
            lLeft = lLeft + lMargin
            If N.Expanded And Not VarType(N.SelectedImage) = vbError Then
                 ImageListDraw hdc, N.SelectedImage, lLeft, Rect.Top + N.Height / 2 - ImgLstHeight / 2, ImgLstWidth, ImgLstHeight
            Else
                ImageListDraw hdc, N.Image, lLeft, Rect.Top + N.Height / 2 - ImgLstHeight / 2, ImgLstWidth, ImgLstHeight
            End If
            lLeft = lLeft + ImgLstWidth '+ lMargin
        End If
        

        If N.TextColor <> -1 Then
            UserControl.ForeColor = N.TextColor
        Else
            If N.Disabled Then
                UserControl.ForeColor = &HCCCCCC
            Else
                If m_SelItem Is N Then
                    If IsDarkColor(m_SelectionColor) And IsDarkColor(m_TextColor) Then
                       UserControl.ForeColor = vbWhite
                    Else
                        UserControl.ForeColor = m_TextColor
                    End If
                Else
                    UserControl.ForeColor = m_TextColor
                End If
            End If
        End If

        
        If N.FontBold Then
            UserControl.FontBold = N.FontBold
        Else
            UserControl.FontBold = False
        End If
        
        
        If N.FontUnderline Or (N.LinkStyle = True And N Is m_HotItem) Then
            UserControl.FontUnderline = True
        Else
            UserControl.FontUnderline = False
        End If
        
        
        lLeft = lLeft + lMargin
       
        If Len(N.SubText) Then
            Dim R As Rect
            R = Rect
            R.Left = lLeft
            R.Bottom = R.Top + N.Height / 2 - lMargin / 2
            If Not m_Font Is Nothing Then Set UserControl.Font = m_Font
            DrawText hdc, StrPtr(N.Text), -1, R, DT_LEFT Or DT_BOTTOM Or DT_SINGLELINE

            UserControl.ForeColor = m_SubTextColor
            R.Top = R.Bottom + lMargin / 2
            R.Bottom = R.Top + N.Height / 2
            If Not m_SubTextFont Is Nothing Then Set UserControl.Font = m_SubTextFont
            DrawText hdc, StrPtr(N.SubText), -1, R, DT_LEFT Or DT_TOP Or DT_SINGLELINE

            If Not m_Font Is Nothing Then Set UserControl.Font = m_Font
            
        Else
            R = Rect
            R.Left = lLeft
            
            DrawText hdc, StrPtr(N.Text), -1, R, DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
        End If

        If Len(N.Badge) Then
            R.Left = 0: R.Top = 0
           
            DrawText hdc, StrPtr(N.Badge), -1, R, DT_CALCRECT
            R.Top = Rect.Top + (Rect.Bottom - Rect.Top) / 2 - R.Bottom / 2 - lMargin / 2
            R.Bottom = R.Top + R.Bottom + lMargin
            R.Left = m_ClientWidth - lMargin * 2 - NSpacing / 2 - BW - R.Right
            R.Right = R.Left + R.Right + lMargin * 2
            
            If N.BadgeBackColor <> -1 Then
                hPath = CreateRoundPath(R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top, 4 * DpiFactor)
                GdipCreateSolidFill RGBtoARGB(N.BadgeBackColor, 100), hBrush
                GdipFillPath mhGraphics, hBrush, hPath
                GdipDeleteBrush hBrush
                GdipDeletePath hPath
                If IsDarkColor(N.BadgeBackColor) Then
                    UserControl.ForeColor = vbWhite
                Else
                    UserControl.ForeColor = vbBlack
                End If
            End If
            
            DrawText hdc, StrPtr(N.Badge), -1, R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
            
        End If
        

        
        N.Left = Rect.Left
        N.Top = Rect.Top
        N.Right = Rect.Right
        N.Bottom = Rect.Bottom
        
        RaiseEvent NodePostPaint(N, hdc, mhGraphics)
        
        End If
        Rect.Top = Rect.Top + H + lMBottom
        If Rect.Top > UserControl.ScaleHeight Then Exit Sub

        If N.Expanded Then
            RecursiveDraw N, Ident + 1
        End If
    Next
    
    
End Sub

Public Function DrawSwitch(hGraphics, Color As Long, Left As Long, Top As Long, Width As Long, Height As Long, Value As Boolean)
    Dim hPath As Long, PenWidth As Long
    Dim hBrush As Long, hPen As Long
    PenWidth = 2 * DpiFactor
    hPath = CreateRoundPath(Left, Top + Height / 4, Width, Height / 2, Height / 2)
    GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    
    GdipCreatePen1 RGBtoARGB(Color, 100), PenWidth, UnitPixel, hPen
    GdipDrawPath hGraphics, hPen, hPath
    GdipDeletePen hPen
    
    
    GdipCreateSolidFill RGBtoARGB(Color, 100), hBrush
    Dim x As Long
    If Value Then
        x = Left + Width - 3 * DpiFactor - Height / 3
    Else
        x = Left + 3 * DpiFactor
    End If
    
    GdipFillEllipse hGraphics, hBrush, x, Top + Height / 2 - Height / 6, Width / 3, Height / 3
    GdipDeleteBrush hBrush

    GdipDeletePath hPath
End Function

Private Sub DrawRadio(Color As Long, Left As Long, Top As Long, Width As Long, Height As Long, Value As Boolean)
    Dim PenWidth As Double
    Dim hBrush As Long, hPen As Long
    
    PenWidth = 2 * DpiFactor
    
    GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
    GdipFillEllipse mhGraphics, hBrush, Left, Top, Width, Height
    GdipDeleteBrush hBrush

    GdipCreatePen1 RGBtoARGB(Color, 100), PenWidth, UnitPixel, hPen
    GdipDrawEllipse mhGraphics, hPen, Left, Top, Width, Height
    GdipDeletePen hPen
    
    If Value Then
        PenWidth = PenWidth * 1.5
        GdipCreateSolidFill RGBtoARGB(Color, 100), hBrush
        GdipFillEllipse mhGraphics, hBrush, Left + PenWidth, Top + PenWidth, Width - PenWidth * 2, Height - PenWidth * 2
        GdipDeleteBrush hBrush
    End If
        
End Sub

Private Sub DrawCheckBox(Color As Long, Left As Long, Top As Long, Width As Long, Height As Long, Value As Boolean)
    Dim PenWidth As Double
    Dim hPath As Long
    Dim hBrush As Long, hPen As Long
    
    PenWidth = 2 * DpiFactor
    
    
    hPath = CreateRoundPath(Left, Top, Width, Height, 2)
    GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
    GdipFillPath mhGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipCreatePen1 RGBtoARGB(Color, 100), PenWidth, UnitPixel, hPen
    GdipDrawPath mhGraphics, hPen, hPath
    GdipDeletePen hPen
    GdipDeletePath hPath
    
    If Value Then
        GdipCreateSolidFill RGBtoARGB(Color, 100), hBrush
        Call GdipCreatePath(&H0, hPath)
        AddPathChekFigure hPath, Left + PenWidth, Top + PenWidth, Width - PenWidth * 2
        GdipFillPath mhGraphics, hBrush, hPath
        GdipDeletePath hPath
        GdipDeleteBrush hBrush
    End If
End Sub

Private Sub AddPathChekFigure(hPath As Long, Left As Long, Top As Long, SizeBox As Single)
    Dim SW As Single, SH As Single
    Dim L As Long, t As Long
    Dim pt(6) As POINTF
    
    SW = SizeBox
    SH = SizeBox * 0.8318841

    L = Left
    t = Top + SizeBox / 2 - SH / 2

    pt(0).x = L + 0.828 * SW: pt(0).Y = t + 0
    pt(1).x = L + 0.329 * SW: pt(1).Y = t + 0.598 * SH
    pt(2).x = L + 0.163 * SW: pt(2).Y = t + 0.397 * SH
    pt(3).x = L + 0 * SW: pt(3).Y = t + 0.594 * SH
    pt(4).x = L + 0.34 * SW: pt(4).Y = t + 1 * SH
    pt(5).x = L + 1 * SW: pt(5).Y = t + 0.205 * SH
    pt(6).x = L + 0.828 * SW: pt(6).Y = t + 0 * SH
    
    GdipAddPathPolygon hPath, pt(0), 7
End Sub

Private Function CreateRoundPath(ByVal Left As Single, ByVal Top As Single, ByVal Width As Single, ByVal Height As Single, ByVal Radius As Single) As Long
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

'*2
Private Sub RecurCalcTreeWidth(TheNode As Node, Ident As Long, ParentExpanded As Boolean)
    Dim N As Node
    Dim lLen As Long

    For Each N In TheNode.Childs
        Dim R As Rect
        R.Left = 0
        If N.Height = 0 Then
            UserControl.FontBold = N.FontBold
            DrawText hdc, StrPtr(N.Text), -1, R, DT_CALCRECT
       
            N.Height = R.Bottom - R.Top
            N.TextWidth = R.Right
            
            If Len(N.SubText) Then
                Set UserControl.Font = m_SubTextFont
                DrawText hdc, StrPtr(N.SubText), -1, R, DT_CALCRECT
                N.Height = N.Height + R.Bottom - R.Top + lMargin
                If R.Right > N.TextWidth Then N.TextWidth = R.Right
                Set UserControl.Font = m_Font
            End If
            
        End If
        
        If N.Height < MNHeight Then N.Height = MNHeight
        lLen = BW + lMargin + (Ident + 1) * NSpacing + NSpacing / 2 + N.TextWidth + lMargin
        If Not VarType(N.Image) = vbError Then
            lLen = lLen + ImgLstWidth + lMargin
            If N.Height < ImgLstHeight + lMargin * 2 Then N.Height = ImgLstHeight + lMargin * 2
        End If
        If m_CheckBoxes Then lLen = lLen + CheckSize + lMargin
        lLen = lLen + lMargin
        If lLen > TreeWidth Then TreeWidth = lLen
    
        If N.Expanded Then RecurCalcTreeWidth N, Ident + 1, N.Expanded
        If ParentExpanded Then TreeHeight = TreeHeight + N.Height + lMBottom
     
    Next
End Sub

Private Function RecurCalcNodeToggle(x As Single, Y As Single, TheNode As Node, Ident As Long) As Boolean
    Dim N As Node

    For Each N In TheNode.Childs
        If m_PlusMinusAlignRight Then
            Rect.Left = m_ClientWidth - lToggleIconSize - lMargin - NSpacing / 2 - BW
        Else
            Rect.Left = -ucScrollbar2.Value + BW + Ident * NSpacing + NSpacing / 2  '- lToggleIconSize / 2
        End If
        
        Rect.Right = Rect.Left + lToggleIconSize + lMargin * 2
        Rect.Bottom = Rect.Top + N.Height
      
        If x >= Rect.Left And x < Rect.Right And Y >= Rect.Top And Y < Rect.Bottom Then
            N.Expanded = Not N.Expanded
            If N.Expanded Then
                RaiseEvent NodeExpand(N)
            Else
                RaiseEvent NodeCollapse(N)
            End If
            Update
            RecurCalcNodeToggle = True
            Exit Function
        End If
        Rect.Top = Rect.Top + N.Height + lMBottom
        If Rect.Top > UserControl.ScaleHeight Then Exit Function
        If N.Expanded Then
            RecurCalcNodeToggle = RecurCalcNodeToggle(x, Y, N, Ident + 1)
        End If
    Next
End Function


Private Sub RecurCalcNodeCheckBox(x As Single, Y As Single, TheNode As Node, Ident As Long, Click As Boolean, DblClick As Boolean)
    Dim N As Node
    Dim i As Long
    Dim lMid As Long
    Dim bInArea As Boolean
    
    For Each N In TheNode.Childs
        If (m_CheckBoxes And Not N.HideCheck) And Not Click Then
            lMid = (N.Bottom - N.Top) / 2
            If x >= N.Left + lMargin And x < N.Left + lMargin + CheckSize And _
                Y >= N.Top + lMid - CheckSize / 2 And Y <= N.Top + lMid + CheckSize / 2 Then
                
                If N.Disabled Then Exit Sub
                
                If TheNode.RadioStyle Then
                    For i = 1 To TheNode.Childs.Count
                         TheNode.Childs(i).Checked = CBool(N Is TheNode.Childs(i))
                    Next
                Else
                    N.Checked = Not N.Checked
                    
                End If
                
                Draw
                RaiseEvent NodeCheck(N)
                
                Exit Sub
            End If
        End If
        
        If m_FullRowSelect Then
            If x >= lMargin And x < m_ClientWidth And Y >= N.Top And Y < N.Bottom Then
                bInArea = True
            End If
        Else
            If x >= N.Left And x < N.Right And Y >= N.Top And Y < N.Bottom Then
                bInArea = True
            End If
        End If
        
        If bInArea Then
            If Not N.Disabled Then

                Set m_SelItem = N
                Draw
                If Click Then
                    RaiseEvent NodeClick(N)
                End If
                
                If DblClick Then
                    RaiseEvent NodeDblClick(N)
                    If N.Childs.Count > 0 Then
                        N.Expanded = Not N.Expanded
                        Update
                    End If
                End If

            End If
            Exit Sub
        End If
        
        Rect.Top = Rect.Top + N.Height + lMBottom
        If Rect.Top > UserControl.ScaleHeight Then Exit Sub
        
        If N.Expanded Then
            RecurCalcNodeCheckBox x, Y, N, Ident + 1, Click, DblClick
        End If
    Next
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    
    RaiseEvent KeyDown(KeyCode, Shift)
    
    Select Case KeyCode
    
        Case vbKeyDown
            SelNextItem cNodes, False
            Rect.Top = lMargin
            EnsureVisible cNodes
        Case vbKeyUp
            SelPrevItem cNodes, Nothing
            Rect.Top = lMargin
            EnsureVisible cNodes
        Case vbKeyPageDown
            Rect.Top = lMargin
            bExit = False
            RecuPageDown cNodes, False
            If bExit = False Then
                RecuSelLast cNodes
            End If
            Rect.Top = lMargin
            EnsureVisible cNodes
            
        Case vbKeyPageUp
            Rect.Top = 0
            
            Do
                SelPrevItem cNodes, Nothing
                If m_SelItem Is cNodes.Childs(1) Then Exit Do
                If Not m_SelItem Is Nothing Then
                    Rect.Top = Rect.Top + m_SelItem.Height
                    
                    If Rect.Top > UserControl.ScaleHeight Then Exit Do
                End If
            Loop
            Rect.Top = lMargin
            EnsureVisible cNodes

        Case vbKeyHome
            If cNodes.Childs.Count > 0 Then
                Set m_SelItem = cNodes.Childs(1)
                If m_SelItem.Disabled Then
                    SelNextItem cNodes, False
                    If m_SelItem.Disabled Then Set m_SelItem = Nothing
                End If
                Rect.Top = lMargin
                EnsureVisible cNodes
            End If
            
        Case vbKeyEnd
            RecuSelLast cNodes
            Rect.Top = lMargin
            EnsureVisible cNodes
            
        Case vbKeyAdd
            m_SelItem.Expanded = True
            RaiseEvent NodeExpand(m_SelItem)
            Update
            
        Case vbKeySubtract
            m_SelItem.Expanded = False
            RaiseEvent NodeCollapse(m_SelItem)
            Update
            
    End Select
End Sub


Private Sub RecuSelLast(TheNode As Node)
   
    If TheNode.Childs.Count > 0 Then
        If TheNode.Childs(TheNode.Childs.Count).Expanded Then
            RecuSelLast TheNode.Childs(TheNode.Childs.Count)
        Else
            Set m_SelItem = TheNode.Childs(TheNode.Childs.Count)
            If m_SelItem.Disabled Then
                SelPrevItem cNodes, Nothing
                If m_SelItem.Disabled Then Set m_SelItem = Nothing
            End If
        End If
    Else
        Set m_SelItem = TheNode
        RaiseEvent SelNodeChange(m_SelItem)
    End If
End Sub


Private Sub RecuPageDown(TheNode As Node, bNext As Boolean)
    Dim N As Node

    For Each N In TheNode.Childs
        If bExit Then Exit Sub
        If bNext Then
            Rect.Top = Rect.Top + N.Height
            
            If Rect.Top >= UserControl.ScaleHeight Then
                If Not N.Disabled Then
                    Set m_SelItem = N
                    RaiseEvent SelNodeChange(m_SelItem)
                    bExit = True
                    Exit Sub
                End If
            End If

        End If
    
        If N Is m_SelItem Then
            bNext = True
        End If
        
        If N.Expanded Then
            RecuPageDown N, bNext
        End If
    Next
End Sub


Private Sub SelNextItem(TheNode As Node, bNext As Boolean)
    Dim N As Node

    For Each N In TheNode.Childs
        If bNext Then
            If Not N.Disabled Then
                Set m_SelItem = N
                RaiseEvent SelNodeChange(m_SelItem)
                bNext = False
                Exit Sub
            End If
        End If
    
    
        If N Is m_SelItem Then
            bNext = True
        End If
        
        If N.Expanded Then
            SelNextItem N, bNext
        End If
    Next
End Sub


Private Sub SelPrevItem(TheNode As Node, PrevNode As Node)
    Dim N As Node

    For Each N In TheNode.Childs
        
        If N Is m_SelItem Then
            If PrevNode Is Nothing Then Exit Sub
            Set m_SelItem = PrevNode
            RaiseEvent SelNodeChange(m_SelItem)
            Exit Sub
        Else
            If Not N.Disabled Then Set PrevNode = N
        End If
        
        If N.Expanded Then
            SelPrevItem N, PrevNode
        End If
    Next
 
End Sub



Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Text1.Visible = False
    
    Rect.Top = lMargin - ucScrollbar1.Value
    RecurCalcNodeCheckBox x, Y, cNodes, 0, False, False
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim oHI As Node
    Dim oHT As Node
    
    Set oHI = m_HotItem
    Set oHT = m_HotToggle
    Set m_HotItem = Nothing
    Set m_HotToggle = Nothing
    
    Rect.Top = lMargin - ucScrollbar1.Value
    RecurHotItem x, Y, cNodes, 0
    
    
    If m_FullRowSelect Then
        If Not m_HotToggle Is Nothing Then
            Set m_HotItem = m_HotToggle
        End If
    End If
    
    If Not m_HotToggle Is oHT Then
        Draw
        'Exit Sub
    End If
    
    If Not m_HotItem Is Nothing Then
        If hCur Then
            If Not m_HotItem.Disabled Then
                SetCursor hCur
                If UserControl.MousePointer <> vbCustom Then UserControl.MousePointer = vbCustom
            Else
                If UserControl.MousePointer <> vbDefault Then UserControl.MousePointer = vbDefault
            End If
        End If
    Else
        If UserControl.MousePointer <> vbDefault Then UserControl.MousePointer = vbDefault
    End If
    
    If Not m_HotItem Is oHI Then
        'If Not m_HotItem Is Nothing Then Debug.Print m_HotItem.Text
        If Not m_HotItem Is Nothing Then
           RaiseEvent NodeMouseEnter(m_HotItem)
         
           If m_HotItem.LinkStyle = True Then
                UserControl.MouseIcon = GetSystemHandCursor
                SetCursor hCur
           Else
                UserControl.MouseIcon = Nothing
                If hCur Then DestroyCursor hCur: hCur = 0
                UserControl.MousePointer = vbCustom
           End If
           
           If Len(m_HotItem.ToolTipText) Then
                UserControl.Extender.ToolTipText = m_HotItem.ToolTipText
           Else
                If m_HotItem.Right > m_ClientWidth Then
                    UserControl.Extender.ToolTipText = m_HotItem.Text
                Else
                     UserControl.Extender.ToolTipText = ""
                End If
            End If
        Else
            RaiseEvent NodeMouseLeave(oHI)
            UserControl.Extender.ToolTipText = ""
        End If
        Draw
    End If
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub RecurHotItem(x As Single, Y As Single, TheNode As Node, Ident As Long)
    Dim N As Node
    Dim H As Long
    Dim lLeft As Long
    For Each N In TheNode.Childs
        H = IIf(N.Height < MNHeight, MNHeight, N.Height)
        
        If Rect.Top + H > 0 Then

            If m_PlusMinusAlignRight Then
                lLeft = m_ClientWidth - lToggleIconSize - lMargin - NSpacing / 2 - BW
            Else
                lLeft = N.Left - lToggleIconSize - lMargin
            End If
            
            If x > lLeft And x < lLeft + lMargin + lToggleIconSize And Y > N.Top And Y < N.Bottom Then
                Set m_HotToggle = N
                Exit Sub
            End If
        
            If m_FullRowSelect Then
                If Y >= N.Top And Y < N.Bottom Then
                    If Not N.Disabled Then
                        Set m_HotItem = N
                    End If
                    Exit Sub
                End If
            Else
                If x >= N.Left And x < N.Right And Y >= N.Top And Y < N.Bottom Then
                    If Not N.Disabled Then
                        Set m_HotItem = N
                    End If
                    Exit Sub
                End If
            End If

        End If
                    
        Rect.Top = Rect.Top + N.Height + lMBottom
        If Rect.Top > UserControl.ScaleHeight Then Exit Sub
        
        If N.Expanded Then
            RecurHotItem x, Y, N, Ident + 1
        End If
        
        If Not m_HotItem Is Nothing Then Exit Sub

    Next

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If m_PlusMinusVisible Then
        Rect.Top = lMargin - ucScrollbar1.Value
        If RecurCalcNodeToggle(x, Y, cNodes, 0) Then
            Exit Sub 'si se hizo click en el PlusMinus que no llame al evento click
        End If
    End If
    
    RaiseEvent MouseUp(Button, Shift, x, Y)
         
    Static lasTimeClick As Single
    Static N As Node
    If (Timer - lasTimeClick) > 1 And (Timer - lasTimeClick) < 3 Then
        If (N Is m_HotItem) And Not m_HotItem Is Nothing Then
            If m_AllowEdit Then
                Dim bCancel As Boolean
                
                Dim lMid As Long
                If (m_CheckBoxes And Not N.HideCheck) Then
                    lMid = (N.Bottom - N.Top) / 2
                    If x >= N.Left + lMargin And x < N.Left + lMargin + CheckSize And _
                        Y >= N.Top + lMid - CheckSize / 2 And Y <= N.Top + lMid + CheckSize / 2 Then
                        Exit Sub
                    End If
                End If
                
                
                If N.Disabled Then Exit Sub

                RaiseEvent BeforeLabelEdit(N, bCancel)
                If Not bCancel Then
                    Call StartEditNode(N)
                End If
            End If
        End If
    End If
    Set N = m_HotItem

    lasTimeClick = Timer
End Sub



Private Function LoadPictureEx(SrcImg As Variant, Optional ByVal Width As Long, Optional ByVal Height As Long) As Long
    Dim hImage1 As Long, hImage2 As Long, hGraphics As Long
    Dim DataArr() As Byte
    Dim lPictureRealWidth As Long, lPictureRealHeight As Long
    Dim x As Long, Y As Long, cx As Long, cy As Long
    Dim sngRatio1 As Single, sngRatio2 As Single
    Dim mpath As Long, Radius As Long, hPen As Long
    
    Select Case VarType(SrcImg)
        Case vbString
            If PathIsURL(SrcImg) Then
            
                If Left$(LCase(SrcImg), 5) = "data:" Then
                    Base64Decode Split(SrcImg, ",")(1), DataArr
                    Call LoadImageFromArray(DataArr, hImage1)
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
                        Call LoadImageFromArray(DataArr, hImage1)
                    End If
                End If
            Else
                Call GdipLoadImageFromFile(ByVal StrPtr(SrcImg), hImage1)
            End If
        Case vbLong
            Const OBJ_BITMAP As Long = 7
            
            If GetObjectType(SrcImg) = OBJ_BITMAP Then
                Call GdipCreateBitmapFromHBITMAP(SrcImg, 0, hImage1)
            Else
                'Call GdipCreateBitmapFromHICON(SrcImg, hImage1)
                Dim hBmp As Long
                If IsAlphaIcon(SrcImg) Then
                   hImage1 = AlphaIconToBmp(SrcImg, hBmp)
                   
                Else
                    GdipCreateBitmapFromHICON SrcImg, hImage1
                End If
                
                
            End If
            
        Case vbDataObject
            Call GdipLoadImageFromStream(SrcImg, hImage1)
            
        Case (vbArray Or vbByte)
            DataArr() = SrcImg
            Call LoadImageFromArray(DataArr, hImage1)
    End Select

    If hImage1 <> 0 Then
        GdipGetImageWidth hImage1, lPictureRealWidth
        GdipGetImageHeight hImage1, lPictureRealHeight
        If Width = 0 Then Width = lPictureRealWidth
        If Height = 0 Then Height = lPictureRealHeight
        
        sngRatio1 = Width / lPictureRealWidth
        sngRatio2 = Height / lPictureRealHeight
        If sngRatio1 > sngRatio2 Then sngRatio1 = sngRatio2
        cx = lPictureRealWidth * sngRatio1: cy = lPictureRealHeight * sngRatio1
        x = (Width - cx) \ 2: Y = (Height - cy) \ 2

        GdipCreateBitmapFromScan0 Width, Height, 0&, PixelFormat32bppPARGB, ByVal 0&, hImage2
        GdipGetImageGraphicsContext hImage2, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias
        Call GdipSetInterpolationMode(hGraphics, 7&)  '//InterpolationModeHighQualityBicubic
        Call GdipSetPixelOffsetMode(hGraphics, 4&)
        If m_ImagesRadius > 0 Then
            Radius = m_ImagesRadius * DpiFactor
            If GdipCreatePath(&H0, mpath) = 0 Then
                x = x + DpiFactor: Y = Y + DpiFactor
                cx = cx - DpiFactor * 2: cy = cy - DpiFactor * 2
                GdipAddPathArcI mpath, x, Y, Radius, Radius, 180, 90
                GdipAddPathArcI mpath, x + cx - Radius, Y, Radius, Radius, 270, 90
                GdipAddPathArcI mpath, x + cx - Radius, Y + cy - Radius, Radius, Radius, 0, 90
                GdipAddPathArcI mpath, x, Y + cy - Radius, Radius, Radius, 90, 90
                GdipClosePathFigure mpath
                GdipSetClipPath hGraphics, mpath, CombineModeIntersect
            End If
        End If
        
        GdipDrawImageRectRectI hGraphics, hImage1, x, Y, cx, cy, 0, 0, lPictureRealWidth, lPictureRealHeight, UnitPixel, 0&, 0&, 0&
        
        If mpath Then
            GdipCreatePen1 RGBtoARGB(m_BackColor, 100), 1 * DpiFactor, &H2, hPen
            GdipResetClip hGraphics
            GdipDrawPath hGraphics, hPen, mpath
            GdipDeletePen hPen
            GdipDeletePath mpath
        End If
        
        GdipDeleteGraphics hGraphics
        GdipDisposeImage hImage1
        LoadPictureEx = hImage2
    End If
    
    If hBmp <> 0 Then GdipDisposeImage hBmp
End Function

Private Function AlphaIconToBmp(ByVal IconHandle As Long, ByRef RefBmp As Long) As Long
    
    Dim tRECT As Rect
    Dim tICONINFO As ICONINFO
    Dim tBitmapData As BitmapData
    Dim lPixelFormat As Long
    Dim sngWidth As Single
    Dim sngHeight As Single

    If GetIconInfo(IconHandle, tICONINFO) <> 0 Then

        If GdipCreateBitmapFromHBITMAP(tICONINFO.hbmColor, 0&, RefBmp) = 0 Then

            If GdipGetImagePixelFormat(RefBmp, lPixelFormat) = 0 Then

                If GdipGetImageDimension(RefBmp, sngWidth, sngHeight) = 0 Then

                    With tRECT
                        .Right = CLng(sngWidth)
                        .Bottom = CLng(sngHeight)
                    End With

                    If GdipBitmapLockBits(RefBmp, tRECT, ImageLockModeRead, lPixelFormat, tBitmapData) = 0 Then
                        
                        Call GdipCreateBitmapFromScan0(tRECT.Right, tRECT.Bottom, tBitmapData.Stride, PixelFormat32bppARGB, ByVal tBitmapData.Scan0, AlphaIconToBmp)

                        Call GdipBitmapUnlockBits(RefBmp, tBitmapData)
                            
                    End If
                    
                End If
                
            End If
            
        End If

        Call DeleteObject(tICONINFO.hbmMask)
        Call DeleteObject(tICONINFO.hbmColor)
        
    End If
    
End Function

Private Function IsAlphaIcon(ByVal IconHandle As Long) As Boolean

    Dim tARGB() As ARGB
    Dim tRECT As Rect
    Dim tICONINFO As ICONINFO
    Dim tBitmapData As BitmapData
    Dim lPixelFormat As Long
    Dim lngX As Long
    Dim lngY As Long
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim lngArgbBmp As Long
    Dim lngColorBmp As Long
    Dim bolRet As Boolean

    If GetIconInfo(IconHandle, tICONINFO) <> 0 Then
        
        If tICONINFO.hbmColor <> 0 Then

            If GdipCreateBitmapFromHBITMAP(tICONINFO.hbmColor, 0&, lngColorBmp) = 0 Then

                If GdipGetImagePixelFormat(lngColorBmp, lPixelFormat) = 0 Then

                    If lPixelFormat <> PixelFormat32bppRGB Then

                        bolRet = False
                        
                    Else

                        If GdipGetImageDimension(lngColorBmp, sngWidth, sngHeight) = 0 Then
                            
                            With tRECT
                                .Right = CLng(sngWidth)
                                .Bottom = CLng(sngHeight)
                            End With

                            ReDim tARGB(tRECT.Right - 1&, tRECT.Bottom - 1&)
                            
                            With tBitmapData
                                .Scan0 = VarPtr(tARGB(0&, 0&))
                                .Stride = 4& * tRECT.Right
                            End With
                            
                            If GdipBitmapLockBits(lngColorBmp, tRECT, ImageLockModeRead Or ImageLockModeUserInputBuf, lPixelFormat, tBitmapData) = 0 Then

                                For lngY = 0 To tBitmapData.Height - 1
                                    For lngX = 0 To tBitmapData.Width - 1
                                    
                                        If tARGB(lngX, lngY).Alpha > 0 Then
                                            If tARGB(lngX, lngY).Alpha < 255 Then
                                            
                                                bolRet = True
                                                Exit For
                                                
                                            End If
                                        End If
                                        
                                    Next lngX

                                    If bolRet Then Exit For
                                    
                                Next lngY
                                
                                Call GdipDisposeImage(lngArgbBmp)

                                Call GdipBitmapUnlockBits(lngColorBmp, tBitmapData)
                                    
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                
                Call GdipDisposeImage(lngColorBmp)
                
            End If
            
            Call DeleteObject(tICONINFO.hbmColor)
        
        End If
        
        Call DeleteObject(tICONINFO.hbmMask)
        
    Else
    
        bolRet = False
        
    End If

    IsAlphaIcon = bolRet
    
End Function


Private Function LoadImageFromArray(ByRef bvData() As Byte, ByRef hImage As Long) As Boolean
    On Local Error GoTo LoadImageFromArray_Error
    Dim IStream     As IUnknown
    If Not IsArrayDim(VarPtrArray(bvData)) Then Exit Function
    
    Call CreateStreamOnHGlobal(bvData(0), 0&, IStream)
    If Not IStream Is Nothing Then
        If GdipLoadImageFromStream(IStream, hImage) = 0 Then
            LoadImageFromArray = True
        End If
    End If
    Set IStream = Nothing
    
LoadImageFromArray_Error:
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

Private Function IsArrayDim(ByVal lpArray As Long) As Boolean
    Dim lAddress As Long
    Call CopyMemory(lAddress, ByVal lpArray, &H4)
    IsArrayDim = Not (lAddress = 0)
End Function

Private Function RGBtoARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long 'By LaVople
    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    RGBtoARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        RGBtoARGB = RGBtoARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        RGBtoARGB = RGBtoARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
End Function
'*-

'Funcion para combinar dos colores
Private Function ShiftColor(ByVal clrFirst As Long, ByVal clrSecond As Long, ByVal lAlpha As Long) As Long
  
    Dim clrFore(3)         As Byte
    Dim clrBack(3)         As Byte
    
    If (clrFirst And &H80000000) Then clrFirst = GetSysColor(clrFirst And &HFF&)
    If (clrSecond And &H80000000) Then clrSecond = GetSysColor(clrSecond And &HFF&)
  
    CopyMemory clrFore(0), clrFirst, 4
    CopyMemory clrBack(0), clrSecond, 4
  
    clrFore(0) = (clrFore(0) * lAlpha + clrBack(0) * (255 - lAlpha)) / 255
    clrFore(1) = (clrFore(1) * lAlpha + clrBack(1) * (255 - lAlpha)) / 255
    clrFore(2) = (clrFore(2) * lAlpha + clrBack(2) * (255 - lAlpha)) / 255
  
    CopyMemory ShiftColor, clrFore(0), 4
  
End Function

Private Function IsDarkColor(ByVal Color As Long) As Boolean
    Dim BGRA(0 To 3) As Byte
    If (Color And &H80000000) Then Color = GetSysColor(Color And &HFF&)
    CopyMemory BGRA(0), Color, 4&
    IsDarkColor = ((CLng(BGRA(0)) + (CLng(BGRA(1) * 3)) + CLng(BGRA(2))) / 2) < 382
End Function

'Private Sub TextBox_SelStart(HWnd As Long, ByVal New_Value As Long)
'    Call SendMessage(HWnd, EM_SETSEL, New_Value, ByVal New_Value)
'End Sub
'
'Private Function TextBox_SetText(HWnd As Long, sText As String) As Long
'    TextBox_SetText = SendMessageW(HWnd, WM_SETTEXT, 0&, ByVal StrPtr(sText))
'End Function
'
'Private Function TextBox_GetText(HWnd As Long) As String
'    Dim sText As String
'    Dim lLength As Long
'
'    lLength = SendMessage(HWnd, WM_GETTEXTLENGTH, 0, ByVal 0&)
'    sText = Space$(lLength + 1)
'
'    Call SendMessageW(HWnd, WM_GETTEXT, lLength + 1, ByVal StrPtr(sText))
'    TextBox_GetText = Left$(sText, lLength)
'End Function

Public Property Get RunMode() As Boolean
On Error Resume Next
    RunMode = True
    RunMode = Ambient.UserMode
    RunMode = Extender.Parent.RunMode
End Property



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ucScrollbar1.Visible = False
    ucScrollbar1.SmoothScrollFactor = 1
    ucScrollbar1.ShowButtons = False
    ucScrollbar1.Style = sGoogle
    ucScrollbar2.ShowButtons = False
    ucScrollbar2.Style = sGoogle
    ucScrollbar2.SmoothScrollFactor = 1

    With PropBag
    
        
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        m_ContainerBkColor = .ReadProperty("ContainerBkColor", Ambient.BackColor)
        m_BackColor = .ReadProperty("BackColor", vbWindowBackground)
        m_TextColor = .ReadProperty("TextColor", Ambient.ForeColor)
        m_BorderRadius = .ReadProperty("BorderRadius", 5)
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_BorderColor = .ReadProperty("BorderColor", vbActiveBorder)
        m_SelectionColor = .ReadProperty("SelectionColor", &HFFE8CC)
        m_FullRowSelect = .ReadProperty("FullRowSelect", True)
        m_CheckBoxes = .ReadProperty("CheckBoxes", False)
        m_ShowTreeLines = .ReadProperty("ShowTreeLines", False)
        m_PlusMinusVisible = .ReadProperty("PlusMinusVisible", True)
        m_PlusMinusStyle = .ReadProperty("PlusMinusStyle", TVP_Arrow)
        m_SubTextColor = .ReadProperty("SubTextColor", &HCCCCCC)
        Set m_SubTextFont = .ReadProperty("SubTextFont", Ambient.Font)
        m_TreeLinesColor = .ReadProperty("TreeLinesColor", &HCCCCCC)
        m_AlternateNodeColors = .ReadProperty("AlternateNodeColors", False)
        m_PlusMinusColor = .ReadProperty("PlusMinusColor", vbBlack)
        m_PlusMinusHotColor = .ReadProperty("PlusMinusHotColor", &HFF6363)
        m_PlusMinusAlignRight = .ReadProperty("PlusMinusAlignRight", False)
        m_CheckBoxColor = .ReadProperty("CheckBoxColor", &HFF6363)
        m_NodeSpacing = .ReadProperty("NodeSpacing", 16)
        m_MinNodeHeight = .ReadProperty("MinNodeHeight", 32)
        m_Margin = .ReadProperty("Margin", 5)
        m_ToggleIconSize = .ReadProperty("ToggleIconSize", 10)
        m_MousePointerHands = .ReadProperty("MousePointerHands", False)
        m_AllowEdit = .ReadProperty("AllowEdit", False)
        Me.OLEDropMode = .ReadProperty("OLEDropMode", vbOLEDropNone)
        m_MarginBottom = .ReadProperty("MarginBottom", 0)
        Set Text1.Font = Me.Font
        Text1.Visible = False
        
        UserControl.BackColor = m_ContainerBkColor
        Set UserControl.Font = m_Font
        ucScrollbar1.StyleBackColor = m_BackColor
        ucScrollbar2.StyleBackColor = m_BackColor
        
        If m_MousePointerHands Then
            If Ambient.UserMode Then
                UserControl.MouseIcon = GetSystemHandCursor
            End If
        Else
            UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
            UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        End If
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Font", m_Font, Ambient.Font
        .WriteProperty "ContainerBkColor", m_ContainerBkColor, Ambient.BackColor
        .WriteProperty "BackColor", m_BackColor, vbWindowBackground
        .WriteProperty "TextColor", m_TextColor, Ambient.ForeColor
        .WriteProperty "BorderRadius", m_BorderRadius, 5
        .WriteProperty "BorderWidth", m_BorderWidth, 1
        .WriteProperty "BorderColor", m_BorderColor, vbActiveBorder
        .WriteProperty "SelectionColor", m_SelectionColor, &HFFE8CC
        .WriteProperty "FullRowSelect", m_FullRowSelect, True
        .WriteProperty "CheckBoxes", m_CheckBoxes, False
        .WriteProperty "ShowTreeLines", m_ShowTreeLines, False
        .WriteProperty "PlusMinusVisible", m_PlusMinusVisible, True
        .WriteProperty "PlusMinusStyle", m_PlusMinusStyle, TVP_Arrow
        .WriteProperty "SubTextColor", m_SubTextColor, &HCCCCCC
        .WriteProperty "SubTextFont", m_SubTextFont, Ambient.Font
        .WriteProperty "TreeLinesColor", m_TreeLinesColor, &HCCCCCC
        .WriteProperty "AlternateNodeColors", m_AlternateNodeColors, False
        .WriteProperty "PlusMinusColor", m_PlusMinusColor, &HFF6363
        .WriteProperty "PlusMinusHotColor", m_PlusMinusHotColor, vbBlack
        .WriteProperty "PlusMinusAlignRight", m_PlusMinusAlignRight, False
        .WriteProperty "CheckBoxColor", m_CheckBoxColor, &HFF6363
        .WriteProperty "NodeSpacing", m_NodeSpacing, 16
        .WriteProperty "MinNodeHeight", m_MinNodeHeight, 32
        .WriteProperty "Margin", m_Margin, 5
        .WriteProperty "ToggleIconSize", m_ToggleIconSize, 10
        .WriteProperty "MousePointerHands", m_MousePointerHands, False
        .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "AllowEdit", m_AllowEdit, False
        .WriteProperty "OLEDropMode", Me.OLEDropMode, vbOLEDropNone
        .WriteProperty "MarginBottom", m_MarginBottom, 0
    End With
End Sub

Private Function GetSystemHandCursor() As Picture
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

Private Sub AjustScrolls()
    On Error Resume Next
    Dim BW As Long, BR As Long
    BW = m_BorderWidth * DpiFactor
    BR = m_BorderRadius * DpiFactor / 1.5
    ucScrollbar1.Move ScaleWidth - ucScrollbar1.Width - BW, BW + BR, _
                    ucScrollbar1.Width, ScaleHeight - BW * 2 - BR * 2 - IIf(ucScrollbar2.Visible, ucScrollbar2.Height, 0)
    ucScrollbar2.Move BW + BR, ScaleHeight - ucScrollbar2.Height - BW, _
                    ScaleWidth - BW * 2 - BR * 2 - IIf(ucScrollbar1.Visible, ucScrollbar1.Width, 0)

End Sub
Private Sub UserControl_Resize()
    AjustScrolls
    Update
End Sub

Private Sub DrawToggleButton(hGraphics As Long, Value As Boolean, Color As Long, ByVal Left As Single, ByVal Top As Single, ByVal SizeBox As Single)
    Dim PenWidth As Double
    Dim hPath As Long
    Dim hBrush As Long, hPen As Long
    
    PenWidth = 1 * DpiFactor
    
    GdipCreatePen1 Color, PenWidth, UnitPixel, hPen
    GdipCreateSolidFill RGBtoARGB(m_BackColor, 100), hBrush
    hPath = CreateRoundPath(Left, Top, SizeBox, SizeBox, 2)
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDrawPath mhGraphics, hPen, hPath
    GdipDeletePath hPath

    GdipDrawLine hGraphics, hPen, Left + PenWidth * 2, Top + SizeBox / 2, Left + SizeBox - PenWidth * 2, Top + SizeBox / 2
    
    If Value = False Then
        GdipDrawLine hGraphics, hPen, Left + SizeBox / 2, Top + PenWidth * 2, Left + SizeBox / 2, Top + SizeBox - PenWidth * 2
    End If
    
    GdipDeletePen hPen
End Sub

Private Sub DrawArrow1(hGraphics As Long, Value As Boolean, bkColor As Long, Color As Long, ByVal Left As Single, ByVal Top As Single, ByVal SizeBox As Single)
    Dim SW As Single, SH As Single
    Dim L As Single, t As Single
    Dim hPath As Long, hBrush As Long
    Dim mAngle As Long
    Dim pt(6) As POINTF

    SW = SizeBox
    SH = SizeBox * 1.1

    If m_ShowTreeLines Then
        GdipCreateSolidFill RGBtoARGB(bkColor, 100), hBrush
        hPath = CreateRoundPath(Left - SizeBox / 4, Top - SizeBox / 4, SizeBox * 1.5, SizeBox * 1.5, SizeBox * 1.5)
        GdipFillPath mhGraphics, hBrush, hPath
    End If

    L = Left - SizeBox / 16
    t = Top

    pt(0).x = L: pt(0).Y = t + SW / 2
    pt(1).x = L + SW / 2: pt(1).Y = t
    pt(2).x = L + SW: pt(2).Y = t + SW / 2
    pt(3).x = L + SW - 0.1 * SH: pt(3).Y = t + SW / 2 + 0.1 * SH
    pt(4).x = L + 0.5 * SW: pt(4).Y = t + 0.2 * SH
    pt(5).x = L + 0.1 * SH: pt(5).Y = t + SW / 2 + 0.1 * SH
    pt(6).x = L: pt(6).Y = t + SW / 2
    
    
    If Value Then
        mAngle = 180
    Else
        mAngle = 90
    End If

    GdipCreatePath 0&, hPath
    GdipAddPathPolygon hPath, pt(0), 7
    
    Dim tMatrix   As Long
    Dim tRECT As RECTL
    Dim cx As Single, cy As Single
    GdipGetPathWorldBoundsI hPath, tRECT, 0&, 0
    cx = L + SizeBox / 2
    cy = t + SizeBox / 2
    GdipCreateMatrix tMatrix
    GdipTranslateMatrix tMatrix, -cx, -cy, MatrixOrderAppend
    GdipRotateMatrix tMatrix, mAngle, MatrixOrderAppend
    GdipTranslateMatrix tMatrix, cx, cy, MatrixOrderAppend
    GdipTransformPath hPath, tMatrix
    
    GdipCreateSolidFill Color, hBrush
    GdipFillPath hGraphics, hBrush, hPath
    GdipDeleteBrush hBrush
    GdipDeletePath hPath
End Sub

Private Sub UserControl_Show()
    Update
End Sub

Private Sub UserControl_Terminate()
    Dim i As Long
    If Not cImageList Is Nothing Then
        For i = 1 To cImageList.Count
            GdipDisposeImage CLng(cImageList(i))
        Next
        
        Set cImageList = Nothing
    End If
 
    Call GdiplusShutdown(GdipToken)
End Sub


Private Sub RecorrerNodos(xmlNode As Object, ByVal parentNode As Node)
    Dim childNode As Object
    Dim newNode As Node
    Dim xmlAttr As Object
    Dim i As Long
    
    Set newNode = parentNode.Add(xmlNode.nodename)
    
    If Not xmlNode.Attributes Is Nothing Then
        For Each xmlAttr In xmlNode.Attributes
            Select Case LCase(xmlAttr.nodename)
                Case "text": newNode.Text = xmlAttr.Text
                Case "subtext": newNode.SubText = xmlAttr.Text
                Case "image": newNode.Image = IIf(IsNumeric(xmlAttr.Text), Val(xmlAttr.Text), xmlAttr.Text)
                Case "selectedimage": newNode.SelectedImage = IIf(IsNumeric(xmlAttr.Text), Val(xmlAttr.Text), xmlAttr.Text)
                Case "checked": newNode.Checked = CBool(xmlAttr.Text)
                Case "plusbutton": newNode.PlusButton = CBool(xmlAttr.Text)
                Case "expanded": newNode.Expanded = CBool(xmlAttr.Text)
                Case "fontbold": newNode.FontBold = CBool(xmlAttr.Text)
                Case "fontunderline": newNode.FontUnderline = CBool(xmlAttr.Text)
                Case "textcolor": newNode.TextColor = HexToColor(xmlAttr.Text)
                Case "backcolor": newNode.BackColor = HexToColor(xmlAttr.Text)
                Case "radiostyle": newNode.RadioStyle = CBool(xmlAttr.Text)
                Case "hidecheck": newNode.HideCheck = CBool(xmlAttr.Text)
                Case "disabled": newNode.Disabled = CBool(xmlAttr.Text)
                Case "selected": newNode.Selected = CBool(xmlAttr.Text)
                Case "key": newNode.key = xmlAttr.Text
                Case "data": newNode.Data = xmlAttr.Text
                Case "linkstyle": newNode.LinkStyle = CBool(xmlAttr.Text)
                Case "badge": newNode.Badge = xmlAttr.Text
                Case "badgebackcolor": newNode.BadgeBackColor = HexToColor(xmlAttr.Text)
                Case "borderwidth": newNode.BorderWidth = CLng(xmlAttr.Text)
                Case "bordercolor": newNode.BorderColor = HexToColor(xmlAttr.Text)
                Case "tooltiptext": newNode.ToolTipText = xmlAttr.Text
            End Select
        Next xmlAttr
    End If
    
    For i = 0 To xmlNode.childNodes.Length - 1
        Set childNode = xmlNode.childNodes.Item(i)
        RecorrerNodos childNode, newNode
    Next i
    
End Sub


Private Function HexToColor(ByVal HexColor As String) As Long
    If Left(HexColor, 1) = "#" Then
        HexColor = Mid(HexColor, 2)
    Else
        HexToColor = Val(HexColor)
        Exit Function
    End If
    
    If Len(HexColor) <> 6 Then
        MsgBox "the value hexadecimal is not correct."
        Exit Function
    End If
    
    Dim Red As Long, Green As Long, Blue As Long
    
    Red = Val("&H" & Mid(HexColor, 1, 2))
    Green = Val("&H" & Mid(HexColor, 3, 2))
    Blue = Val("&H" & Mid(HexColor, 5, 2))
    
    HexToColor = RGB(Red, Green, Blue)
End Function
