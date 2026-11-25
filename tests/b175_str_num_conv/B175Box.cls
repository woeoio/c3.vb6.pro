VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B175Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_N As Long
Private m_D As Double
Private m_S As String

Public Property Get N() As Long
    N = m_N
End Property

Public Property Let N(ByVal v As Long)
    m_N = v
End Property

Public Property Get D() As Double
    D = m_D
End Property

Public Property Let D(ByVal v As Double)
    m_D = v
End Property

Public Property Get S() As String
    S = m_S
End Property

Public Property Let S(ByVal v As String)
    m_S = v
End Property
