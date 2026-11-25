VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B163Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public S As String
Private m_T As String

Public Sub Bump(ByRef x As String)
    x = x & "-ok"
End Sub

Public Sub Swap(ByRef x As String)
    x = "new"
End Sub

Public Property Get Tag() As String
    Tag = m_T
End Property

Public Property Let Tag(ByVal v As String)
    m_T = v
End Property
