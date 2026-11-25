VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B168Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_A(1 To 5) As Long
Private m_N(1 To 5) As String

' 账 #305 §11.150 (第一百六十八刀): 带索引的属性 —— 读-改-写。
' 带一个索引:
Public Property Get Slot(ByVal i As Long) As Long
    Slot = m_A(i)
End Property

Public Property Let Slot(ByVal i As Long, ByVal v As Long)
    m_A(i) = v
End Property

' 带两个索引 (Let 的取值形参在最后):
Public Property Get Pair(ByVal i As Long, ByVal j As Long) As Long
    Pair = m_A(i) + m_A(j) * 100
End Property

Public Property Let Pair(ByVal i As Long, ByVal j As Long, ByVal v As Long)
    m_A(i) = v Mod 100
    m_A(j) = v \ 100
End Property

' String 那一档的带索引属性 (临时量的 BSTR 归调用方退):
Public Property Get Tag(ByVal i As Long) As String
    Tag = m_N(i)
End Property

Public Property Let Tag(ByVal i As Long, ByVal v As String)
    m_N(i) = v
End Property

' 被调方: 只写回形参, 用来验「读完有没有写回属性」。
Public Sub Bump(ByRef x As Long)
    x = x + 1
End Sub

Public Sub App(ByRef s As String)
    s = s & "!"
End Sub

Private Sub Class_Initialize()
    m_A(1) = 1
    m_A(2) = 10
    m_A(3) = 3
    m_N(1) = "ab"
End Sub
