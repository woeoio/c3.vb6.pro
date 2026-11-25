VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B160Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_S As String
Private m_D As Double
Private m_I As Integer
Private m_B As Boolean

Public Property Get S() As String
    S = m_S
End Property
Public Property Let S(ByVal x As String)
    m_S = x
End Property

Public Property Get D() As Double
    D = m_D
End Property
Public Property Let D(ByVal x As Double)
    m_D = x
End Property

Public Property Get I() As Integer
    I = m_I
End Property
Public Property Let I(ByVal x As Integer)
    m_I = x
End Property

Public Property Get B() As Boolean
    B = m_B
End Property
Public Property Let B(ByVal x As Boolean)
    m_B = x
End Property
