VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B176Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_S1 As String
Private m_N As Long

Public Property Get S1() As String
    S1 = m_S1
End Property

Public Property Let S1(ByVal v As String)
    m_S1 = v
End Property

Public Property Get N() As Long
    N = m_N
End Property

Public Property Let N(ByVal v As Long)
    m_N = v
End Property
