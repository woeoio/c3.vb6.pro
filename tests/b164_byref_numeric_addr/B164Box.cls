VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B164Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public N As Long
Public I As Integer
Public D As Double
Public S As Single
Public B As Boolean
Public Y As Byte

Public Property Get Tag() As Long
    Tag = m_T
End Property

Public Property Let Tag(ByVal v As Long)
    m_T = v
End Property

Private m_T As Long

Public Sub Inc(ByRef x As Long)
    x = x + 1
End Sub

Public Sub IncI(ByRef x As Integer)
    x = x + 1
End Sub

Public Sub IncD(ByRef x As Double)
    x = x + 1.5
End Sub

Public Sub IncS(ByRef x As Single)
    x = x + 0.5
End Sub

Public Sub FlipB(ByRef x As Boolean)
    x = True
End Sub

Public Sub SetY(ByRef x As Byte)
    x = 7
End Sub
