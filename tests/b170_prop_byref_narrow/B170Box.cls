VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B170Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_L As Long
Private m_I As Integer
Private m_B As Byte
Private m_D As Double
Private m_S As Single
Private m_C As Currency

Public Property Get L1() As Long
    L1 = m_L
End Property

Public Property Let L1(ByVal v As Long)
    m_L = v
End Property

Public Property Get I1() As Integer
    I1 = m_I
End Property

Public Property Let I1(ByVal v As Integer)
    m_I = v
End Property

Public Property Get B1() As Byte
    B1 = m_B
End Property

Public Property Let B1(ByVal v As Byte)
    m_B = v
End Property

Public Property Get D1() As Double
    D1 = m_D
End Property

Public Property Let D1(ByVal v As Double)
    m_D = v
End Property

Public Property Get S1() As Single
    S1 = m_S
End Property

Public Property Let S1(ByVal v As Single)
    m_S = v
End Property

' 货币 (证人: 本刀仍按值, 在册 —— 与货币数组元素同一张欠项)
Public Property Get C1() As Currency
    C1 = m_C
End Property

Public Property Let C1(ByVal v As Currency)
    m_C = v
End Property

Public Sub IncI(ByRef x As Integer)
    x = x + 1
End Sub

Public Sub IncL(ByRef x As Long)
    x = x + 1
End Sub

Public Sub IncB(ByRef x As Byte)
    x = x + 1
End Sub

Public Sub IncD(ByRef x As Double)
    x = x + 1
End Sub

Public Sub IncC(ByRef x As Currency)
    x = x + 1
End Sub
