VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "P3122"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_pic As Object

Public Property Set Pic(ByVal v As Object)
    Set m_pic = v
End Property

Public Property Get Pic() As Object
    Set Pic = m_pic
End Property

Private Function GetV() As Variant
    Dim o As P3122
    Set o = New P3122
    Set GetV = o
End Function

Public Sub LoadIt()
    Set Pic = GetV()
End Sub

Public Sub LoadRaw()
    Dim o As P3122
    Set o = New P3122
    Set Pic = o
End Sub
