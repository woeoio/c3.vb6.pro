VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B169Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Tag As Long
Private m_S As String
Private m_A(1 To 5) As Long

Public Property Get Tag() As Long
    Tag = m_Tag
End Property

Public Property Let Tag(ByVal v As Long)
    m_Tag = v
End Property

Public Property Get S() As String
    S = m_S
End Property

Public Property Let S(ByVal v As String)
    m_S = v
End Property

' 带索引的 (与 With 合用那一格)
Public Property Get Slot(ByVal i As Long) As Long
    Slot = m_A(i)
End Property

Public Property Let Slot(ByVal i As Long, ByVal v As Long)
    m_A(i) = v
End Property

Public Sub Inc(ByRef x As Long)
    x = x + 1
End Sub

' 型不一致那一格的证人: Long 属性接 Integer 形参 (本刀仍按值, 在册)
Public Sub IncI(ByRef x As Integer)
    x = x + 1
End Sub

Public Sub App(ByRef t As String)
    t = t & "!"
End Sub
