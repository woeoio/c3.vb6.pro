VERSION 5.00
Begin VB.Form frmUCtest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GIF 动画播放控件Ver1.0 枕善居　http://mndsoft.com"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   StartUpPosition =   1  '所有者中心
   Begin prjNewAniGif.ucGIF ucGIF2 
      Height          =   495
      Left            =   3120
      Top             =   3960
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      fileLen         =   0
   End
   Begin prjNewAniGif.ucGIF ucGIF1 
      Height          =   495
      Left            =   2520
      Top             =   3960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      fileLen         =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "帧进度"
      Height          =   1215
      Left            =   5640
      TabIndex        =   20
      Top             =   4140
      Width           =   1335
      Begin VB.CommandButton cmdNav 
         Caption         =   "重新播放"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   ">>"
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   22
         ToolTipText     =   "下一帧"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdNav 
         Caption         =   "<<"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Previous Frame"
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   195
      TabIndex        =   17
      Top             =   4095
      Width           =   1935
      Begin VB.OptionButton optScaleMode 
         Caption         =   "发光"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optScaleMode 
         Caption         =   "阴暗"
         Height          =   255
         Index           =   1
         Left            =   915
         TabIndex        =   18
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.OptionButton optBorder 
      Caption         =   "凹陷"
      Height          =   255
      Index           =   3
      Left            =   5400
      TabIndex        =   16
      Top             =   210
      Width           =   1575
   End
   Begin VB.OptionButton optBorder 
      Caption         =   "凸起"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   210
      Width           =   975
   End
   Begin VB.OptionButton optBorder 
      Caption         =   "平面"
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   14
      Top             =   210
      Width           =   735
   End
   Begin VB.OptionButton optBorder 
      Caption         =   "无边框"
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   13
      Top             =   210
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   795
      Left            =   45
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   5580
      Width           =   6915
   End
   Begin VB.ComboBox cboMinDelay 
      Height          =   300
      ItemData        =   "frmUCtest.frx":0000
      Left            =   1455
      List            =   "frmUCtest.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4935
      Width           =   2295
   End
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "全部复原"
      Height          =   315
      Index           =   2
      Left            =   4110
      TabIndex        =   6
      ToolTipText     =   "复原到原始设置"
      Top             =   4500
      Width           =   1110
   End
   Begin VB.CommandButton cmdSpeed 
      Caption         =   ">>"
      Height          =   315
      Index           =   1
      Left            =   5220
      TabIndex        =   5
      ToolTipText     =   "Speed Up Frame +50 ms"
      Top             =   4500
      Width           =   330
   End
   Begin VB.CommandButton cmdSpeed 
      Caption         =   "<<"
      Height          =   315
      Index           =   0
      Left            =   3780
      TabIndex        =   4
      ToolTipText     =   "慢播放 -50 ms"
      Top             =   4500
      Width           =   330
   End
   Begin VB.ComboBox cboFrame 
      Height          =   300
      ItemData        =   "frmUCtest.frx":0078
      Left            =   1455
      List            =   "frmUCtest.frx":007A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4500
      Width           =   2295
   End
   Begin VB.CommandButton cmdUndo 
      Caption         =   "复位"
      Height          =   420
      Left            =   30
      TabIndex        =   2
      Top             =   3525
      Width           =   2415
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "透明"
      Height          =   420
      Left            =   30
      TabIndex        =   1
      Top             =   3105
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   2580
      ItemData        =   "frmUCtest.frx":007C
      Left            =   30
      List            =   "frmUCtest.frx":009E
      TabIndex        =   0
      Top             =   465
      Width           =   2400
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3495
      Left            =   2475
      ScaleHeight     =   3435
      ScaleWidth      =   4440
      TabIndex        =   11
      Top             =   465
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   3720
      Picture         =   "frmUCtest.frx":014A
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "每帧最小延时"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   4980
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "帧延时"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   1170
   End
   Begin VB.Label lblFrame 
      Caption         =   "帧: 0  延时:  0"
      Height          =   225
      Left            =   75
      TabIndex        =   7
      Top             =   225
      Width           =   3735
   End
End
Attribute VB_Name = "frmUCtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
' This form is not well remarked, but I would think is pretty easy to follow.

' The user control is very well remarked & it would be wise to get familiar
' with the public properties & functions so you can use it to its fullest
' capabilities.

' Strong suggestion. If you are thinking of using this for dozens of animated
' gifs, I would compile the control into a separate OCX vs compiling into your
' application.  Speed would be the difference.

' You will not that the usercontrol has a few public events too and they
' are used in some of the examples here


' FILES REQUIRED FOR THE USERCONTROL:

' 1. The ucGIF .CTL File
' 2. tHE ClsGIFframe .CLS File
' 3. The ppGIFpreview .PAG File

' All other files are solely for this example project.

Option Explicit

Private Sub cboMinDelay_Click()
    ucGIF1.gifMinFrameDelay = Val(cboMinDelay.Text)
End Sub

Private Sub cmdDo_Click()

    If cmdUndo.Enabled Then Exit Sub
    
    ' variables used for the various examples
    Dim newWindowCx As Long, newWindowCy As Long
    Dim newScale As Long, frameNr As Long, custDC As Long
    Dim gSrcX As Long, gSrcY As Long
    Dim tileRect As RECT
    
    With ucGIF1
    
        If List1.ListIndex <> 4 Then .Visible = False
        .Animate = False
    
        Select Case List1.ListIndex
        Case 0: 'reduced-scale
            .AutoSize = False ' remove window size lock
            .Stretch = ScaleFrames    ' ensure frames are scaled
            newScale = CLng(Rnd * 30 + 20)  ' reduce window randomly
            newWindowCx = (.gifLogicalWindowCx - (.gifLogicalWindowCx * (newScale / 100))) \ Screen.TwipsPerPixelX
            newWindowCy = (.gifLogicalWindowCy - (.gifLogicalWindowCy * (newScale / 100))) \ Screen.TwipsPerPixelY
            .ScaleLogWinToSize newWindowCx, newWindowCy, False
            .Move .Left, .Top, newWindowCx * Screen.TwipsPerPixelX, newWindowCy * Screen.TwipsPerPixelY
        
        Case 1: 'reduced-stretch
            .AutoSize = False  ' remove window size lock
            .Stretch = StretchFrames  ' ensure frames are stretched
            .Width = .Width \ 2 ' reduce window
            .Height = .Height - .Height * 0.1
        
        Case 2: ' enlarged-scale
            .AutoSize = False  ' remove window size lock
            .Stretch = ScaleFrames    ' ensure frames are scaled
            newScale = CLng(Rnd * 30 + 20)  ' enlarge window randomly
            newWindowCx = (.gifLogicalWindowCx + (.gifLogicalWindowCx * (newScale / 100))) \ Screen.TwipsPerPixelX
            newWindowCy = (.gifLogicalWindowCy + (.gifLogicalWindowCy * (newScale / 100))) \ Screen.TwipsPerPixelY
            .ScaleLogWinToSize newWindowCx, newWindowCy, False
            .Move .Left, .Top, newWindowCx * Screen.TwipsPerPixelX, newWindowCy * Screen.TwipsPerPixelY
        
        Case 3: 'enlarged-stretch
            .AutoSize = False  ' remove window size lock
            .Stretch = StretchFrames  ' ensure frames are stretched
            .Width = .Height * 2 ' enlarge window
        
        Case 4: ' looping
            MsgBox "The number of loops will be changed to 10 from 1000." & vbCrLf & _
                "Animation will stop after 10 more loops" & vbCrLf & vbCrLf & _
                "Tip: Look at label at top of listbox for the countdown.", vbInformation + vbOKOnly
            ' yes, we said 10, but let's add 1 more sore current loop will finish
            ' and our info label will show Remaing Loops as 10 vs 9
            .gifMaxLoops = 11
        
        Case 5: ' forced transparency
            For frameNr = 1 To .gifFrameCount
                If .Frame(frameNr).isTransparent = False Then .Frame(frameNr).ForceTransparency = True
            Next
            ucGIF2.Animate = False
            For frameNr = 1 To ucGIF2.gifFrameCount
                If ucGIF2.Frame(frameNr).isTransparent = False Then ucGIF2.Frame(frameNr).ForceTransparency = True
            Next
            ucGIF2.Animate = True
        
        Case 6: ' overlays. Long remarks due to trying to explain how stdPic.Render works
        
            ' I'm using the StdPicture.Render function below. You could easily do some
            ' DC/BltBlt stuff. But Render doesn't require DCs from a StdPicture object.
            
            ' In this particular example I cheated a little when blting. I'm making the gif control
            ' invisible at top of module, tile blting to picbox, then Rendering to gif control's
            ' special DC. Because the gif control is invisible, no problems blting from picbox;
            ' but if it was visible, then chances are we'd get garbage in the blt since control
            ' overlaps the picbox.
            ' If I didn't want to hide the control & still wanted to blt from the underlying
            ' control, I'd need to bring the underlying picbox's image into a DC & blt from
            ' there, or in the case of the tile blting routine, I'd tile blt onto the
            ' special DC directly, starting at a negative top/left coordinate so that it
            ' would perfectly align over the picture box's image.
            ' An example is shown below. Just change the True to False
            
            
            ' just fluff -- tileblt a bitmap across the background DC
            tileRect.Right = Picture1.ScaleWidth \ Screen.TwipsPerPixelX
            tileRect.Bottom = Picture1.ScaleHeight \ Screen.TwipsPerPixelY
            TileBltRectEx Picture1.hdc, tileRect, Image1.Picture
            
            ' 1. Calculate the how far left & top gif control is from the left/top of picbox.
            ' 2. Convert to pixels for subtraction within the Render function
            '       also pixels needed if using BitBlt
            ' 3. Ensure you include the sunken border offsets for the picture box
            ' 4. Also include the border style for the usercontrol if any
            
            ' this function will return the bordersize*2
            .ScaleLogWinToSize 1, 1, , gSrcY
            ' now convet the left/top differences between the two windows, in pixels
            gSrcX = ScaleX((.Left - Picture1.Left), Me.ScaleMode, vbPixels) - 2 + gSrcY / 2
            gSrcY = ScaleY((.Top - Picture1.Top), Me.ScaleMode, vbPixels) - 2 + gSrcY / 2
                    
            ' call usercontrol function to return DC & Bitmap to paint to
            custDC = .GetCustomBkgDC()
            
            If True Then    ' option to blt using Render. Change to False to direct blt
        
                With Picture1.Image ' note that the .Image is a StdPicture class
                
                    ' when using Render for a portion of a picture vs whole thing, need to
                    ' do some funky calcs cause the stdPic is stored upside down
                    ' The 1st 5 params are self-explanatory, they are always Long data types,
                    '   but if variables, then add +0 or errors more times than not
                    ' The next 4 are not & will discuss each one now.
                    ' ScaleX(gSrcX, vbPixels, vbHimetric):: left position to start blting
                    '   no issue here, but we need to pass vbHimetric
                    ' ScaleY(tileRect.Bottom - gSrcY, vbPixels, vbHimetric)::top position to start blting
                    '   upside down image, so we need to subtract the full height of the image from the
                    '   starting point we want to blt. tileRect.Bottom is the full height of the image
                    ' ScaleX(ucGIF1.GifLogicalWindowCx, vbPixels, vbHimetric)::width of the blt
                    '   no issue here, we need to pass vbHimetric
                    ' -ScaleY(ucGIF1.GifLogicalWindowCy, vbPixels, vbHimetric)::height of blt
                    '   upside down, so pass the negative value; else image is blt'd upside down
                    
                    .Render custDC + 0, 0, 0, ucGIF1.gifLogicalWindowCx + 0, ucGIF1.gifLogicalWindowCy + 0, _
                    ScaleX(gSrcX, vbPixels, vbHimetric), ScaleY(tileRect.Bottom - gSrcY, vbPixels, vbHimetric), _
                    ScaleX(ucGIF1.gifLogicalWindowCx, vbPixels, vbHimetric), -ScaleY(ucGIF1.gifLogicalWindowCy, vbPixels, vbHimetric), ByVal 0&
                End With
            Else
                ' option tileblts directly to the DC. This would be a similar method
                ' you would use if you wanted to BitBlt from a memory DC.
                tileRect.Left = -gSrcX
                tileRect.Top = -gSrcY
                tileRect.Right = .gifLogicalWindowCx
                tileRect.Bottom = .gifLogicalWindowCy
                TileBltRectEx custDC, tileRect, Image1
            End If
            .SetCustomBkgDC
        
        Case 7 ' simple background. In this example, I am tileblting an icon
         ' call usercontrol function to return DC & Bitmap to paint to
         custDC = .GetCustomBkgDC()
         ' just fluff -- tileblt a bitmap across the background DC
         tileRect.Right = .gifLogicalWindowCx
         tileRect.Bottom = .gifLogicalWindowCy
         TileBltRectEx custDC, tileRect, Me.Icon.Handle, True
         .SetCustomBkgDC
        
        Case 8 ' frame removal
            ' simply remove these frames to force animation to start/end
            ' in the limited form size we have. Always delete in reverse
            For frameNr = 21 To 8 Step -1
                .RemoveFrame frameNr
            Next
            ' update our frames combo box
            ' While were at it, this fish moves too slow for me -- speed it up
            cboFrame.Clear
            For frameNr = 1 To ucGIF1.gifFrameCount
                cboFrame.AddItem "帧 " & frameNr
                .Frame(frameNr).frameDelay = 300
            Next
            cboFrame.AddItem "所有帧", 0
            cboFrame.ListIndex = 0
            
        Case 9
            ' Two ways of adjusting either top and/or left of the frames
            ' 1. is to do it directly. However, you may adjust the frame so
            '    that it now falls outside the logical window dimensions.
            ' 2. is to use the MoveFrame function. This will automatically
            '    resize the window, if necessary, after adjusting the frame position
            ' change True to False & vice versa to toggle the options
            If True Then    ' option 1
                .Frame(4).FrameTop = .Frame(4).FrameTop + 6
            Else            ' option 2
                .MoveFrame 4, .Frame(4).FrameLeft, .Frame(4).FrameTop + 6
            End If
        End Select
        
    .Animate = True
    .Visible = True
    
    End With
    
    cmdUndo.Enabled = True
End Sub

Private Sub cmdNav_Click(Index As Integer)
    
    With cboFrame
        If Index = 2 Then
            ucGIF1.Animate = True
        Else
            If ucGIF1.Animate Then
                ucGIF1.Animate = False
                .ListIndex = 0
            End If
            If Index = 1 Then ' next
                If .ListIndex = .ListCount - 1 Then
                    .ListIndex = 1
                Else
                    .ListIndex = .ListIndex + 1
                End If
            Else
                If .ListIndex < 2 Then
                    .ListIndex = .ListCount - 1
                Else
                    .ListIndex = .ListIndex - 1
                End If
            End If
            lblFrame.Caption = "帧: " & .ListIndex & " 延时: " & ucGIF1.Frame(.ListIndex).frameDelay & "毫秒"
            ucGIF1.ShowFrame .ListIndex
        End If
    End With
    
End Sub

Private Sub cmdSpeed_Click(Index As Integer)
    Dim adjDelay As Long, nrFrames As Long, beginFrame As Long, fNr As Long
    
    If Index = 0 Then adjDelay = 50 Else adjDelay = -50
    
    If cboFrame.ListIndex = 0 Or Index = 2 Then
        nrFrames = cboFrame.ListCount - 1
        beginFrame = 1
    Else
        beginFrame = cboFrame.ListIndex
        nrFrames = beginFrame
    End If
    For fNr = beginFrame To nrFrames
        If Index < 2 Then
            ucGIF1.Frame(fNr).frameDelay = ucGIF1.Frame(fNr).frameDelay + adjDelay
        Else
            ucGIF1.Frame(fNr).ResetFrameAttributes fattrDelay
        End If
    Next
End Sub

Private Sub cmdUndo_Click()
    ucGIF1.RemoveCustomBkgDC
    List1.Tag = -1
    Call List1_Click
End Sub

Private Sub Form_Load()
    If ValidateSampleFiles = False Then
        Unload Me
    Else
        List1.ListIndex = 6
    End If
End Sub

Private Sub LoadSampleGif(Index As Integer)

    ' always reset some stuff before doing a different example
    
    ' the Force Transparency example
    If ucGIF2.Visible Then
        ucGIF2.Visible = False
        ucGIF2.Animate = False
    End If
    Picture1.Visible = False
    ucGIF1.Move Picture1.Left, Picture1.Top
    
    ' if the overlay or background example was previously show, then
    ' if we don't tell gif control to remove bacground, it will continue to be used
    ucGIF1.RemoveCustomBkgDC
    
    Select Case Index
    Case 0, 1, 2, 3:
        ucGIF1.LoadGIF App.Path & "\Fox_.gif"
    Case 4
        ucGIF1.LoadGIF App.Path & "\cheetah.gif"
    Case 5
        ucGIF1.LoadGIF App.Path & "\dragon_right.gif"
        ucGIF2.Move ucGIF1.Left, ucGIF1.Top + ucGIF1.Height + 30
        ucGIF2.LoadGIF App.Path & "\Line_of_flames_3.gif"
        ucGIF2.Animate = True
        ucGIF2.Visible = True
    Case 6, 7
        ucGIF1.LoadGIF App.Path & "\dove.gif"
        ucGIF1.gifMinFrameDelay = 30
    Case 8
        ucGIF1.LoadGIF App.Path & "\fish04.gif"
    Case 9
        ucGIF1.LoadGIF App.Path & "\GooseAni3.gif"
    End Select
    
    cmdUndo.Enabled = False
    
    cboFrame.Clear
    Dim fNr As Long
    For fNr = 1 To ucGIF1.gifFrameCount
        cboFrame.AddItem "帧 " & fNr
    Next
    cboFrame.AddItem "所有帧", 0
    cboFrame.ListIndex = 0
    lblFrame.Caption = "帧: 1 延时: " & ucGIF1.Frame(1).frameDelay & "毫秒"
    
    For fNr = 0 To cboMinDelay.ListCount - 1
        If Val(cboMinDelay.List(fNr)) = ucGIF1.gifMinFrameDelay Then Exit For
    Next
    If fNr = cboMinDelay.ListCount Then
        cboMinDelay.AddItem ucGIF1.gifMinFrameDelay & " 毫秒"
        cboMinDelay.ListIndex = cboMinDelay.ListCount - 1
    Else
        cboMinDelay.ListIndex = fNr
    End If
    
    ucGIF1.Animate = True

End Sub

Private Sub List1_Click()
    If Val(List1.Tag) <> List1.ListIndex Or cmdUndo.Enabled = False Then
        
        ucGIF1.AutoSize = True
        ucGIF1.gifMinFrameDelay = 50
        
        LoadSampleGif List1.ListIndex
        List1.Tag = List1.ListIndex
    
        Select Case List1.ListIndex
        Case 0, 2: ' reduce-scale/stretch
            Text1 = "Shrinking animated GIFs. Two methods to calculate the resizing can be used. The Light version " & _
                "simply call one VB function. However, when reducing the image, the Heavy method may provide a " & _
                "better image quality. That methods converts frames to bitmaps & StretchBlts the images"
        Case 1, 3: ' enlarge-scale/stretch
            Text1 = "Enlarging animated GIFS. Two methods to calculate the resizing can be used. The Light version " & _
                "simply call one VB function. The source quality will determine the quality of the stretched image."
        Case 4: ' looping
            Text1 = "The cheetah you are looking at will stop after 1,000 loops. But by setting a property, you can " & _
                "decrease/increase the number of loops. You can also set the number of loops to zero which means infinite."
        Case 5 ' force transparency
            Text1 = "Not all GIFs are transparent. Those that are 100% non-transparent can generally be tweaked to appear " & _
                "transparent. This will not work on all GIFs. The dragon above has a white background & the fire has black, " & _
                "but both are 100% not transparent."
        Case 6 ' overlay example
            Dim tileRect As RECT
            Picture1.Visible = True
            Picture1.ZOrder 1
            ' Let's get a random top/left position for the GIF control
            ' Random so you know I'm not pulling a fast one on you
            ucGIF1.Left = CLng(Rnd * (Picture1.Width - ucGIF1.Width - 60)) + Picture1.Left + 30
            ucGIF1.Top = CLng(Rnd * (Picture1.Height - ucGIF1.Height - 60)) + Picture1.Top + 30
            
            ' just fluff -- tileblt a bitmap across the background DC
            tileRect.Right = Picture1.ScaleWidth \ Screen.TwipsPerPixelX
            tileRect.Bottom = Picture1.ScaleHeight \ Screen.TwipsPerPixelY
            TileBltRectEx Picture1.hdc, tileRect, Image1.Picture
        
            Text1 = "The user control is a solid object. To make it appear transparent, you simply need to draw the underlying " & _
                "background onto the user control. There are two different examples in the code that do this. Once you accurately " & _
                "clone the background, the control appears to be transparent."
                
        Case 7 ' backgrounds
            Text1 = "The user control can be made to have any background color, image or custom drawing. " & _
                "There are 3 functions which expose the user control's DC: GetCustomBkgDC, SetCustomBkgDC & RemoveCustomBkgDC"
        
        Case 8 ' removals
            Text1 = "The fish above has an animation window that extends way past our form's right edge. " & _
                "You can remove frames with a simple call. Once you click 'Show Me', 14 frames will be " & _
                "removed so that the animation starts/stops on our visible window."
                
        Case 9  ' repositioning frames
            Text1 = "Look at the neck of the goose above. It jumps a lot between the last & first frames. " & _
                "You can adjust the top/left position of any frame. When you click 'Show Me', the last frame " & _
                "will be adjusted so the neck & head of the goose stay pretty close to same position from frame to frame."
        End Select
    End If
End Sub

Private Sub optBorder_Click(Index As Integer)
    ucGIF1.PauseAnimation False
    ucGIF1.BorderStyle = Index
    ucGIF1.PauseAnimation True
End Sub

Private Sub optScaleMode_Click(Index As Integer)
    If cmdUndo.Enabled Then ucGIF1.Animate = False
    ucGIF1.StretchMethod = Index
    If cmdUndo.Enabled Then ucGIF1.Animate = True
End Sub

Private Sub ucGIF1_AnimationLoopComplete(ByVal loopsRemaining As Long)
    If List1.ListIndex = 4 Then
        lblFrame.Caption = "Loops remaining: " & loopsRemaining
    End If
End Sub

Private Sub ucGIF1_AnimationLoopExpired()
    If List1.ListIndex = 4 Then
        lblFrame.Caption = "Example Finished: Click Reset"
        ucGIF1.ShowFrame 4
    End If
End Sub

Private Sub ucGIF1_AnimationProgress(ByVal currentFrame As Long)
    If List1.ListIndex <> 4 Then
        lblFrame.Caption = "帧: " & currentFrame & " 延时: " & ucGIF1.Frame(currentFrame).frameDelay & "毫秒"
    End If
End Sub

Private Function ValidateSampleFiles() As Boolean
    Dim I As Integer, sFile As String, errMsg As String
    For I = 1 To 7
        Select Case I
        Case 1: sFile = "Fox_.gif"
        Case 2: sFile = "cheetah.gif"
        Case 3: sFile = "dragon_right.gif"
        Case 4: sFile = "Line_of_flames_3.gif"
        Case 5: sFile = "dove.gif"
        Case 6: sFile = "fish04.gif"
        Case 7: sFile = "GooseAni3.gif"
        End Select
        If Len(Dir(App.Path & "\" & sFile)) = 0 Then
            errMsg = errMsg & vbCrLf & sFile
        End If
    Next
    If Len(errMsg) Then
        MsgBox "Please download the missing GIF files included in your zip to same directory as this sample project" & vbCrLf & vbCrLf & _
            "The sample project needs them in order to show you the test form" & vbCrLf & errMsg, vbInformation + vbOKOnly
    Else
        ValidateSampleFiles = True
    End If
End Function
