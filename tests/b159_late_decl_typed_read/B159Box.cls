VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B159Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_S As String
Private m_D As Double
Private m_O As Object
Private m_V As Variant

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

Public Property Get O() As Object
    Set O = m_O
End Property
Public Property Set O(ByVal x As Object)
    Set m_O = x
End Property

Public Property Get V() As Variant
    V = m_V
End Property
Public Property Let V(ByVal x As Variant)
    m_V = x
End Property

Public Property Get Tag2() As String
    Tag2 = "101"
End Property
