VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Frames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 探针夹具 (§38 第二站): 类私有 UDT 动态数组的**元素**再带动态数组字段。
' PhotoDemon pdGIF 的形状：`m_allFrames(frameIndex).pixelData(0, 0)`。
' 两站都要覆盖：成员访问站 (m_allFrames(i).pixelData) 与 With 站 (.pixelData)。
' 负控: 类自己的 2 参方法 Pick 必须照旧发成调用 (`Pick(1, 2)`)。

Private Type Frame
    pixelData() As Byte
End Type

Private m_allFrames() As Frame

Public Sub Fill(ByVal frameIndex As Long)
    ReDim m_allFrames(0 To 3)
    ReDim m_allFrames(frameIndex).pixelData(0 To 7, 0 To 7) As Byte
    m_allFrames(frameIndex).pixelData(0, 0) = 1
End Sub

Public Function ReadOne(ByVal frameIndex As Long) As Byte
    ReadOne = m_allFrames(frameIndex).pixelData(0, 0)
End Function

Public Function ReadWith(ByVal frameIndex As Long) As Byte
    With m_allFrames(frameIndex)
        ReadWith = .pixelData(1, 1)
    End With
End Function

Public Function Pick(ByVal a As Long, ByVal b As Long) As Long
    Pick = a + b
End Function

Public Function UsePick() As Long
    UsePick = Me.Pick(1, 2)
End Function
