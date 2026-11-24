VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B161Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_I As Integer
Private m_Y As Byte
Private m_L As Long

Public Property Get I() As Integer
    I = m_I
End Property
Public Property Let I(ByVal x As Integer)
    m_I = x
End Property

Public Property Get Y() As Byte
    Y = m_Y
End Property
Public Property Let Y(ByVal x As Byte)
    m_Y = x
End Property

Public Property Get L() As Long
    L = m_L
End Property
Public Property Let L(ByVal x As Long)
    m_L = x
End Property
