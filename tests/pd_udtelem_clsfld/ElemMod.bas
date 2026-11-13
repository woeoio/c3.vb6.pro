Attribute VB_Name = "ElemMod"
Option Explicit

' 整编 pdGIF 形状: Private Type PD_GifFrame 里 frameDIB/backupFrameDIB 是工程类字段,
' 容器是 UDT 数组 ⇒ `Set m_frames(i).DIB = …` 的目标是 VB6_SA_AT(<UDT>, …).<类字段>。
Public Type ElemRec
    FrameTime As Long
    BackupDIB As ElemObj
    DIB As ElemObj
End Type

Private m_frames() As ElemRec

Public Sub Main()
    ReDim m_frames(0 To 1)
    Set m_frames(0).DIB = New ElemObj
    Set m_frames(0).BackupDIB = m_frames(0).DIB
    Set m_frames(0).DIB = Nothing
    If (Not m_frames(0).BackupDIB Is Nothing) Then Debug.Print "kept"
    If m_frames(0).DIB Is Nothing Then Debug.Print "gone"
End Sub
