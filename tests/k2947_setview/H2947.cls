VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "H2947"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IK2947

Private m_own As Object
Private m_item As String

Private Property Set IK2947_Owner(ByVal o As Object)
    Set m_own = o
End Property

Private Property Get IK2947_Owner() As Object
    Set IK2947_Owner = m_own
End Property

Private Function IK2947_Has() As Long
    If m_own Is Nothing Then
        IK2947_Has = 0
    Else
        IK2947_Has = 1
    End If
End Function

Private Property Let IK2947_Item(ByVal i As Long, ByVal v As String)
    m_item = i & ":" & v
End Property

Private Property Get IK2947_Item(ByVal i As Long) As String
    IK2947_Item = m_item
End Property
Public Property Get OwnView() As IK2947
    Set OwnView = Me
End Property

Public Function ViaMe() As Long
    With Me
        ViaMe = .OwnView.Has
    End With
End Function
