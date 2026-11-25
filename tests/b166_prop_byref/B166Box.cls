VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B166Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_T As Long
Private m_Name As String
Private m_I As Integer
Private m_D As Double
Private m_Y As Byte
Private m_F As Boolean
Private m_Tm As Date

Public Property Get Tag() As Long
    Tag = m_T
End Property

Public Property Let Tag(ByVal v As Long)
    m_T = v
End Property

Public Property Get Label2() As String
    Label2 = m_Name
End Property

Public Property Let Label2(ByVal v As String)
    m_Name = v
End Property

Public Property Get INum() As Integer
    INum = m_I
End Property

Public Property Let INum(ByVal v As Integer)
    m_I = v
End Property

Public Property Get DNum() As Double
    DNum = m_D
End Property

Public Property Let DNum(ByVal v As Double)
    m_D = v
End Property

Public Property Get YNum() As Byte
    YNum = m_Y
End Property

Public Property Let YNum(ByVal v As Byte)
    m_Y = v
End Property

Public Property Get FNum() As Boolean
    FNum = m_F
End Property

Public Property Let FNum(ByVal v As Boolean)
    m_F = v
End Property

Public Property Get Tm() As Date
    Tm = m_Tm
End Property

Public Property Let Tm(ByVal v As Date)
    m_Tm = v
End Property

Public Sub Inc(ByRef x As Long)
    x = x + 1
End Sub

Public Sub IncI(ByRef x As Integer)
    x = x + 1
End Sub

Public Sub IncD(ByRef x As Double)
    x = x + 1.5
End Sub

Public Sub SetY(ByRef x As Byte)
    x = 7
End Sub

Public Sub FlipB(ByRef x As Boolean)
    x = Not x
End Sub

Public Sub AddDay(ByRef x As Date, ByVal n As Long)
    x = DateAdd("d", n, x)
End Sub

Public Sub Append(ByRef x As String, ByVal tail As String)
    x = x & tail
End Sub
