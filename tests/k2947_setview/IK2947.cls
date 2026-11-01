VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "IK2947"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Set Owner(ByVal o As Object)

End Property

Public Property Get Owner() As Object

End Property

Public Function Has() As Long

End Function

' 第一百八十六刀: 多入口参数那形 —— 写在视图上要走 put_ 席 (带下标实参)。
Public Property Let Item(ByVal i As Long, ByVal v As String)

End Property

Public Property Get Item(ByVal i As Long) As String

End Property
