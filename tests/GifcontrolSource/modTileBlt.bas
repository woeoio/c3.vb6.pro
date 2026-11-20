Attribute VB_Name = "modTileBlt"
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

' APIs used to create transparent Bitmap (if applicable)
Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, ByRef lpBits As Any) As Long
Private Declare Function GetMapMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetMapMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetTextColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

' API to convert vbSysColor to standard RGB color
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Const COLOR_BTNFACE As Long = 15

' APIs use to draw our tiles
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetBkColor Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' APIs used to determine what type an image is when image is a handle only
Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Declare Function GetIconInfo Lib "user32.dll" (ByVal hIcon As Long, ByRef piconinfo As ICONINFO) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetGDIObject Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Sub TileBltRectEx(destDC As Long, destRect As RECT, _
    sourcePicObjectOrHandle As Variant, _
    Optional bStaggered As Boolean, _
    Optional bNoEraseBkg As Boolean = False, _
    Optional ByVal TransparentColor As Long = -1)
    
' This is the extended version which allows memory handles, and optional
' staggered offsets & background preservation. See how simple the original
' version looked. Its at end of this module

' PARAMETERS
' destDC :: the DC where the tileBlt will occur
' destRect :: a rectangle structure, in pixels, where in the DC the tileBlt will occur
' sourcePicObjectOrHandle :: can be a picture object or an image handle
'   - Picture Object examples: Image1.Picture, Picture1.Picture, Picture1.Image, stdPicture
'   - Handle :: Image1.Picture1.Handle, memory image handle
'   Supported image types include bitmap, jpg, gif, icon & cursor
' Following parameters used primarily for transparent tiles
' bStaggered :: will offset alternating rows by 1/2 the image width
' bNoEraseBkg :: option to preserve destination DC & tile over it.
'   This option should only be used only for transparent tile images (icons,cursors,gifs)
'   - If False, then the destDC backcolor will be used for base of TileBlt
'   - If True, then the destDC will be copied & routine will TileBlt over it
'     This will be somewhat slower, since the tiled image needs to be converted
'     to a transparent bitmap then overlayed on the destination DC
' TransparentColor :: only used if bNoEraseBkg is True
'   - The transparent color of the source image. -1 will use the top left pixel color

' //////////// OVERVIEW COMMENTS \\\\\\\\\\\
' pretty fast Tile Bltter - compatible with Win95
' Why not just use CreatePatternBrush, right?
' Wrong... that API can only use a maximum of an 8x8 image in Win95.
' Ok, don't care about Win95, what about speed comparisons anyway?
' Using following test scenario & begin timing after offscreen DC & patternbrush created:
' -- Using CreatePatternBrush 1000 times on full screen bitmap: amazing avg of 20 secs
' -- Using my routine 1000 times on a full screen bitmap: avg of 6 secs
' The above test was done using a 48x48 bitmap

' Typical tile bltters will tile an image n number times the image width until the
' image surpasses the destination width & then blt the row vertically in similar fashion.
' Tested using a destination area of 1,000 x 1,000 & image of 16x16, typical routines:
'   63 blts across (1000/16= 62.5) & 62 blts down (don't need to blt over 1st row)
'   for a total of 125 loops to completely fill the area. But...
' This routine, however, does incremental blts, using the image for the 1st blt only,
' then recursing over the previously bltted parts to blt even larger areas. Using
' the same dimensions above, the incremental bltter used only 14 loops.

' 1st loop: 16 pixels blt total
' 2nd loop: 32 pixels blt total
' 3rd loop: 64 pixels blt total
' 4th loop: 128 pixels blt total
' 5th loop: 256 pixels blt total
' 6th loop: 512 pixels blt total
' 7th loop: finished the remaining 488 pixels for total of 1000 x 16
' 8th thru 14th loop: using the entire top row & same algorithm going vertically down
' 14 loops compared to 125 with typical routines

' NOW TO THE ROUTINE

' sanity checks
If destDC = 0 Then Exit Sub
If destRect.Right <= destRect.Left Then Exit Sub
If destRect.Bottom <= destRect.Top Then Exit Sub

Dim tPic As StdPicture, bIsIcon As Boolean
Dim iWidth As Long, iHeight As Long

' get width & height of our tile in pixels
If TypeOf sourcePicObjectOrHandle Is StdPicture Then
    ' need hard reference to use Picture object's Render function
    Set tPic = sourcePicObjectOrHandle
    iWidth = ConvertHimetrix2Pixels(tPic.Width, True)
    iHeight = ConvertHimetrix2Pixels(tPic.Height, False)
    
Else    ' not a Picture object -- test to see what it might be
    Dim tBitmapInfo As BITMAP, tIconInfo As ICONINFO

    ' see if this is not a bitmap handle (jpgs, bmps & gifs return true)
    If GetGDIObject(sourcePicObjectOrHandle, Len(tBitmapInfo), tBitmapInfo) = 0 Then
        ' probably an icon or cursor -- check
        If GetIconInfo(sourcePicObjectOrHandle, tIconInfo) Then
            ' we still need the width & height, so let's get it from the
            ' icon/cursor bitmap objects
            If tIconInfo.hbmColor Then
                GetGDIObject tIconInfo.hbmColor, Len(tBitmapInfo), tBitmapInfo
            ElseIf tIconInfo.hbmMask Then
                GetGDIObject tIconInfo.hbmMask, Len(tBitmapInfo), tBitmapInfo
                ' b&w icons/cursors have mask stacked below non-mask,
                ' so height is actually 1/2
                tBitmapInfo.bmHeight = tBitmapInfo.bmHeight / 2
            End If
        End If
        ' the GetIconInfo API returns bitmaps that need to be destroyed else leaks
        If tIconInfo.hbmColor Then DeleteObject tIconInfo.hbmColor
        If tIconInfo.hbmMask Then DeleteObject tIconInfo.hbmMask
        bIsIcon = True
    End If
    ' set the image height and width & exit if not available
    iWidth = tBitmapInfo.bmWidth
    iHeight = tBitmapInfo.bmHeight
    If iWidth = 0 Or iHeight = 0 Then Exit Sub
End If

' total area to tile
Dim dcWidth As Long, dcHeight As Long
' offscreen bitmap & DC
Dim tBmp As Long, dBmp As Long, oDC As Long
' background fill
Dim backRect As RECT, backBrush As Long
' value indicating how much area covered while in a loop
Dim bltArea As Long

' get width & height of our target area in pixels
dcWidth = destRect.Right - destRect.Left + 1
dcHeight = destRect.Bottom - destRect.Top + 1

' create an off-screen DC to do our drawing
' also this serves as an effective clipping region
oDC = CreateCompatibleDC(destDC)
dBmp = CreateCompatibleBitmap(destDC, dcWidth, dcHeight)
tBmp = SelectObject(oDC, dBmp)

' cause we can be using a transparent gif, icon or cursor, let's
' draw a blank rectangle on our offscreen DC, then draw image over that

Dim transDC As Long, hTransBmp As Long

If bNoEraseBkg Then
    ' user wants to preserve background image on target DC
    
    ' see if we need to check for backcolor
    If TransparentColor = -1 Then
        ' create DC & select tile into it
        transDC = CreateCompatibleDC(destDC)
        hTransBmp = SelectObject(transDC, sourcePicObjectOrHandle)
        ' let's get the top left pixel color & unselect the tile
        TransparentColor = GetPixel(transDC, 0, 0)
        SelectObject transDC, hTransBmp
        ' if we get an invalid number, we will use button face
        If TransparentColor < 0 Then TransparentColor = GetSysColor(COLOR_BTNFACE)
        
    ElseIf TransparentColor < 0 Then
        ' not -1 & still negative (must be a vbSysColor). Convert to standard RGB
        TransparentColor = GetSysColor(TransparentColor And &HFF&)
        
    End If
Else
    ' standard tiling -- over complete area. Use target DC's backcolor
    TransparentColor = GetBkColor(destDC)
End If

' create rectangle of image size & then create brush
' account for staggering also
SetRect backRect, 0, 0, iWidth, iHeight - (bStaggered * iHeight)
backBrush = CreateSolidBrush(TransparentColor)
' draw the blank rectangle
FillRect oDC, backRect, backBrush
DeleteObject backBrush

' blt the first tile on our off-screen DC
If tPic Is Nothing Then
    ' no stdPic
    If bIsIcon Then
        ' use this API & no other DCs needed
        DrawIconEx oDC, 0, 0, sourcePicObjectOrHandle, 0, 0, 0, 0, &H3
        If bStaggered Then
            DrawIconEx oDC, -iWidth / 2, iHeight, sourcePicObjectOrHandle, 0, 0, 0, 0, &H3
        End If
    Else
        ' create another DC if needed, select source img & blt it from there
        If transDC = 0 Then transDC = CreateCompatibleDC(destDC)
        hTransBmp = SelectObject(transDC, sourcePicObjectOrHandle)
        BitBlt oDC, 0, 0, iWidth, iHeight, transDC, 0, 0, vbSrcCopy
        If bStaggered Then
            BitBlt oDC, 0, iHeight, iWidth / 2, iHeight, transDC, iWidth / 2, 0, vbSrcCopy
        End If
        ' unselect the image -- clean up a little later
        SelectObject transDC, hTransBmp
    End If
Else
    ' Picture object's Render is a pretty powerful tool. We will use it here to
    ' blt directly to an offscreen DC without needing to select the image into
    ' a secondary DC as is common practice.
    With tPic
        .Render oDC + 0&, 0&, 0&, iWidth + 0&, iHeight + 0&, 0&, .Height, .Width, -.Height, ByVal 0&
        If bStaggered Then
            .Render oDC + 0&, 0&, iHeight + 0&, iWidth / 2, iHeight + 0&, .Width / 2, .Height, .Width / 2, -.Height, ByVal 0&
        End If
    End With
    Set tPic = Nothing  ' not needed any more
End If
' if a temp DC was created, then delete it
If transDC Then DeleteDC transDC

' now we will blt the rest using incremental steps
bltArea = iWidth
Do Until bltArea * 2 >= dcWidth
    BitBlt oDC, bltArea, 0, bltArea, iHeight, oDC, 0, 0, vbSrcCopy
    bltArea = bltArea * 2
Loop
If bltArea < dcWidth Then ' need IF so below calc doesn't pass negative blt width
    ' did as much as we could w/o overdrawing, now blt the remainder
    BitBlt oDC, bltArea, 0, dcWidth - bltArea, iHeight, oDC, 0, 0, vbSrcCopy
End If
    
If dcHeight > iHeight Then
    ' do the same vertically, 1st row already done
    If bStaggered Then
        BitBlt oDC, (iWidth / 2), iHeight, dcWidth - (iWidth / 2), iHeight, oDC, 0, 0, vbSrcCopy
        bltArea = iHeight * 2
    Else
        bltArea = iHeight
    End If
    
    Do Until bltArea * 2 >= dcHeight
        BitBlt oDC, 0, bltArea, dcWidth, bltArea, oDC, 0, 0, vbSrcCopy
        bltArea = bltArea * 2
    Loop
    If bltArea < dcHeight Then  ' need IF so below calc doesn't pass negative blt height
        ' did as much as we could w/o overdrawing, now blt the remainder
        BitBlt oDC, 0, bltArea, dcWidth, dcHeight - bltArea, oDC, 0, 0, vbSrcCopy
    End If
End If

' if wanting to preserve destDC, then send our tiled bitmap & destDC to
' function to overlay the two so transparency effect is applied
If bNoEraseBkg Then DrawTransparentBitmap destDC, destRect.Left, destRect.Top, dcWidth, dcHeight, oDC, TransparentColor

' copy the offscreen DC to the window DC, offsetting it on the way
BitBlt destDC, destRect.Left, destRect.Top, dcWidth, dcHeight, oDC, 0, 0, vbSrcCopy

' clean up
DeleteObject SelectObject(oDC, tBmp)
DeleteDC oDC

End Sub

Private Function ConvertHimetrix2Pixels(vHiMetrix As Long, byWidth As Boolean) As Long

' handy converter when you can't use ScaleX/ScaleY
If byWidth Then
    ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelX
Else
    ConvertHimetrix2Pixels = vHiMetrix * 1440 / 2540 / Screen.TwipsPerPixelY
End If

End Function


Private Sub DrawTransparentBitmap(lHDCdest As Long, srcX As Long, srcY As Long, _
    destWidth As Long, destHeight As Long, lHDCsrc As Long, lMaskColor As Long)
                                                    
Const DSna = &H220326 '0x00220326
' =====================================================================
' A pretty good transparent bitmap maker I use in several projects
' Modified here to remove stuff I wont use (i.e., Flipping/Rotating images)
' =====================================================================

' Excuse lack of comments; this is pretty much a standard bitmap mask creator
' Removed most comments 'cause they didn't apply after I tweaked this for this project

    ' memory bitmaps & placeholders
    Dim lBmMask As Long, lBmAndMem As Long, lBmColor As Long
    Dim lBmObjectOld As Long, lBmMemOld As Long, lBmColorOld As Long
    ' memory DCs
    Dim lHDCMem As Long, lHDCMask As Long, lHDCcolor As Long
    
    'Create some DCs & bitmaps
    lHDCMask = CreateCompatibleDC(lHDCsrc)
    lHDCMem = CreateCompatibleDC(lHDCsrc)
    lHDCcolor = CreateCompatibleDC(lHDCsrc)
    
    lBmColor = CreateCompatibleBitmap(lHDCsrc, destWidth, destHeight)
    lBmAndMem = CreateCompatibleBitmap(lHDCsrc, destWidth, destHeight)
    lBmMask = CreateBitmap(destWidth, destHeight, 1&, 1&, ByVal 0&)
    
    lBmColorOld = SelectObject(lHDCcolor, lBmColor)
    lBmMemOld = SelectObject(lHDCMem, lBmAndMem)
    lBmObjectOld = SelectObject(lHDCMask, lBmMask)
    
' ====================== Start working here ======================
    
    SetMapMode lHDCMem, GetMapMode(lHDCdest)
    
    BitBlt lHDCMem, 0&, 0&, destWidth, destHeight, lHDCdest, srcX, srcY, vbSrcCopy
    
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)
    
    BitBlt lHDCcolor, 0&, 0&, destWidth, destHeight, lHDCsrc, 0, 0, vbSrcCopy
    
    SetBkColor lHDCcolor, lMaskColor
    SetTextColor lHDCcolor, vbWhite
    
    BitBlt lHDCMask, 0&, 0&, destWidth, destHeight, lHDCcolor, 0&, 0&, vbSrcCopy
    
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    
    BitBlt lHDCcolor, 0, 0, destWidth, destHeight, lHDCMask, 0, 0, DSna
    BitBlt lHDCMem, 0, 0, destWidth, destHeight, lHDCMask, 0&, 0&, vbSrcAnd
    BitBlt lHDCMem, 0&, 0&, destWidth, destHeight, lHDCcolor, 0, 0, vbSrcPaint
    
    ' transfer result
    BitBlt lHDCsrc, 0, 0, destWidth, destHeight, lHDCMem, 0&, 0&, vbSrcCopy
    
    'clean up
    DeleteObject SelectObject(lHDCcolor, lBmColorOld)
    DeleteObject SelectObject(lHDCMask, lBmObjectOld)
    DeleteObject SelectObject(lHDCMem, lBmMemOld)
    DeleteDC lHDCMem
    DeleteDC lHDCMask
    DeleteDC lHDCcolor
End Sub



Private Sub TileBltRect(destDC As Long, destRect As RECT, srcImg As StdPicture)
' much simpler version when bitmaps are used & no fancy offsets are needed

Dim iWidth As Long, iHeight As Long
Dim dcWidth As Long, dcHeight As Long
Dim backRect As RECT, backBrush As Long

Dim tBmp As Long, dBmp As Long, oDC As Long
Dim bltArea As Long

' get width & height of our image in pixels
iWidth = ConvertHimetrix2Pixels(srcImg.Width, True)
iHeight = ConvertHimetrix2Pixels(srcImg.Height, False)

' get width & height of our client area in pixels
dcWidth = destRect.Right - destRect.Left + 1
dcHeight = destRect.Bottom - destRect.Top + 1

' create an off-screen DC to do our drawing
oDC = CreateCompatibleDC(destDC)
dBmp = CreateCompatibleBitmap(destDC, dcWidth, dcHeight)
tBmp = SelectObject(oDC, dBmp)

' create rectangle of image size & then create brush
SetRect backRect, 0, 0, iWidth, iHeight * 2
backBrush = CreateSolidBrush(GetBkColor(destDC))
' draw the blank rectangle
FillRect oDC, backRect, backBrush
DeleteObject backBrush

' blt the first tile on our off-screen DC
With srcImg
    .Render oDC + 0&, 0&, 0&, iWidth + 0&, iHeight + 0&, 0&, .Height, .Width, -.Height, ByVal 0&
End With

' now we will blt the rest using incremental steps
bltArea = iWidth
Do Until bltArea * 2 >= dcWidth
    BitBlt oDC, bltArea, 0, bltArea, iHeight, oDC, 0, 0, vbSrcCopy
    bltArea = bltArea * 2
Loop
If bltArea < dcWidth Then ' need IF so below calc doesn't pass negative blt width
    ' did as much as we could w/o overdrawing, now blt the remainder
    BitBlt oDC, bltArea, 0, dcWidth - bltArea, iHeight, oDC, 0, 0, vbSrcCopy
End If

' do the same vertically, 1st row already done
bltArea = iHeight
Do Until bltArea * 2 >= dcHeight
    BitBlt oDC, 0, bltArea, dcWidth, bltArea, oDC, 0, 0, vbSrcCopy
    bltArea = bltArea * 2
Loop
If bltArea < dcHeight Then ' need IF so below calc doesn't pass negative blt height
    ' did as much as we could w/o overdrawing, now blt the remainder
    BitBlt oDC, 0, bltArea, dcWidth, dcHeight - bltArea, oDC, 0, 0, vbSrcCopy
End If

' the above tileblt over the border area, now optional drawing of a region frame

' copy the offscreen DC to the window DC, offsetting it on the way
BitBlt destDC, destRect.Left, destRect.Top, dcWidth, dcHeight, oDC, 0, 0, vbSrcCopy
' clean up
DeleteObject SelectObject(oDC, tBmp)
DeleteDC oDC

End Sub

