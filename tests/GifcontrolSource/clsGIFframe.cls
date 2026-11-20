VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGIFframe"
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
' Class to simply store information about each frame in the animated GIF.

' Note that I am attempting to prevent the user from changing certain properties.
' This is needed 'cause if you don't compile this as a separate control, then the
' Friend properties will be available to you.  And those are the properties that
' I don't want users to modify. Therefore, developed a simple way to preventing
' users from modifying certain properties even though they can see them.

Public Enum FrameAttrEnum
    [_minAttr] = 1
    fattrAll = 0
    fattrDelay = 1
    fattrTransColor = 2
    fattrDisposal = 4
    fattrLeft = 8
    fattrTop = 16
    fattrForceTrans = 32
    [_maxAttr] = 32
End Enum
Private Enum FrameAttrEnumX
    fattrWidth = 64
    fattrHeight = 128
    fattrOffsetA = 256
    fattrOffsetZ = 512
    fattrTransparent = 1024
End Enum


' 0=user setting, 1=original setting
Private cTop(0 To 1) As Integer          ' image left.top.width.height
Private cLeft(0 To 1) As Integer
Private cWidth As Integer                ' read only
Private cHeight As Integer               ' read only
Private cDelay(0 To 1) As Integer        ' image delay
Private cDisposal(0 To 1) As Byte        ' image disposal rule
Private cTransColor(0 To 1) As Long      ' if so, the transparency color
Private cOffset(0 To 1) As Long          ' image offsets in GIF file (read only)
Private cTransparent As Boolean          ' is frame transparent (read only)
Private cForceTransparency As Boolean    ' option to override transparency values
Private cInitiState As Long              ' flags indicating if settings are due to initalization

Public Property Let FrameLeft(ByVal vValue As Long)
    cLeft(0) = vValue
    If (cInitiState Or fattrLeft) <> cInitiState Then
        cInitiState = cInitiState Or fattrLeft
        cLeft(1) = vValue
    End If
End Property
Public Property Get FrameLeft() As Long
    FrameLeft = cLeft(0)
End Property

Public Property Let FrameTop(ByVal vValue As Long)
    cTop(0) = vValue
    If (cInitiState Or fattrTop) <> cInitiState Then
        cInitiState = cInitiState Or fattrTop
        cTop(1) = vValue
    End If
End Property
Public Property Get FrameTop() As Long
    FrameTop = cTop(0)
End Property

Public Property Let frameDelay(ByVal vValue As Integer)
   If vValue < 0 Then vValue = 0
    cDelay(0) = vValue
    If (cInitiState Or fattrDelay) <> cInitiState Then
        cInitiState = cInitiState Or fattrDelay
        cDelay(1) = vValue
    End If
End Property
Public Property Get frameDelay() As Integer
    frameDelay = cDelay(0)
End Property

Public Property Let FrameDisposal(vValue As Byte)
    cDisposal(0) = vValue
    If (cInitiState Or fattrDelay) <> cInitiState Then
        cInitiState = cInitiState Or fattrDisposal
        cDisposal(1) = vValue
    End If
End Property
Public Property Get FrameDisposal() As Byte
    FrameDisposal = cDisposal(0)
End Property

Public Property Let FrameTransparentColor(vValue As Long)
    cTransColor(0) = vValue
    If (cInitiState Or fattrTransColor) = cInitiState Then
        ' user changing transparent color
        If cTransparent = False Then
            ' if wasn't transparent, then we need to force transparency to allow
            ' usercontrol's bitmap routines to process GIF file vs VB's Render
            If (cInitiState Or fattrTransparent) = cInitiState Then ForceTransparency = True
        End If
    Else
        ' GIF parser finished reading transparency options. Set flags
        cInitiState = cInitiState Or fattrTransColor
        cInitiState = cInitiState Or fattrTransparent
        cTransColor(1) = vValue
    End If
End Property
Public Property Get FrameTransparentColor() As Long
    FrameTransparentColor = cTransColor(0)
End Property

Public Property Let ForceTransparency(vValue As Boolean)
    cForceTransparency = vValue
    If vValue = False Then cTransColor(0) = cTransColor(1)
End Property
Public Property Get ForceTransparency() As Boolean
    ForceTransparency = cForceTransparency
End Property

Public Function ResetFrameAttributes(ByVal fAttribute As FrameAttrEnum) As Long

' function allows user to reset attributes to original settings
' By passing the negative value of the fAttribute parameter, user can
' see the original setting. Not all attributes are exposed, by design

Dim fAttr As Long, rtnVal As Long

If fAttribute < fattrAll Then
    Select Case Abs(fAttribute)
    Case fattrTransColor: rtnVal = cTransColor(1)
    Case fattrDelay: rtnVal = cDelay(1)
    Case fattrDisposal: rtnVal = cDisposal(1)
    Case fattrTop: rtnVal = cTop(1)
    Case fattrLeft: rtnVal = cLeft(1)
    Case fattrForceTrans = False
    End Select
Else
    On Error GoTo ExitRoutine
    fAttr = [_minAttr]
    Do While fAttr < [_maxAttr]
        If (fAttribute And fAttr) = fAttr Then
            Select Case fAttr
            Case fattrDelay: cDelay(0) = cDelay(1)
            Case fattrDisposal: cDisposal(0) = cDisposal(1)
            Case fattrLeft: cLeft(0) = cLeft(1)
            Case fattrTop: cTop(0) = cTop(1)
            Case fattrForceTrans
                If cTransColor(1) = -1 Then cTransColor(0) = -1
                cForceTransparency = False
            Case fattrTransColor
                If cTransColor(1) = -1 Then cForceTransparency = False
                cTransColor(0) = cTransColor(1)
            End Select
        End If
        fAttr = fAttr * 2
        rtnVal = rtnVal Or fAttr
    Loop
End If

ResetFrameAttributes = rtnVal

ExitRoutine:
End Function

Private Sub Class_Initialize()
cTransColor(0) = -1
cTransColor(1) = -1
End Sub

' READ ONLY BY USERS.  DO NOT MODIFY THESE PROPERTIES

Friend Property Let FrameWidth(ByVal vValue As Long)
' can only be changed by the user control
    If (cInitiState Or fattrWidth) <> cInitiState Then
        cWidth = vValue
        cInitiState = cInitiState Or fattrWidth
    End If
End Property
Public Property Get FrameWidth() As Long
    FrameWidth = cWidth
End Property

Friend Property Let FrameHeight(ByVal vValue As Long)
' can only be changed by the user control
    If (cInitiState Or fattrHeight) <> cInitiState Then
        cHeight = vValue
        cInitiState = cInitiState Or fattrHeight
    End If
End Property
Public Property Get FrameHeight() As Long
    FrameHeight = cHeight
End Property

Public Property Get isTransparent() As Boolean
    isTransparent = cTransparent
End Property
Friend Property Let isTransparent(vValue As Boolean)
    cTransparent = vValue
End Property

Friend Property Let byteOffset(OffsetA As Boolean, bytePos As Long)
' can only be changed by the user control
    If OffsetA Then
        If (cInitiState Or fattrOffsetA) <> cInitiState Then
            cInitiState = cInitiState Or fattrOffsetA
            cOffset(0) = bytePos
        End If
    Else
        If (cInitiState Or fattrOffsetZ) <> cInitiState Then
            cInitiState = cInitiState Or fattrOffsetZ
            cOffset(1) = bytePos
        End If
    End If
End Property
Friend Property Get byteOffset(OffsetA As Boolean) As Long
    byteOffset = cOffset(OffsetA + 1)
End Property

