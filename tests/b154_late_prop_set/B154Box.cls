VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B154Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mIt As B154Item

Public Property Get Held() As Long
    If mIt Is Nothing Then
        Held = -1
    Else
        Held = mIt.Tag
    End If
End Property

Public Property Get ItemIsNothing() As Boolean
    ItemIsNothing = (mIt Is Nothing)
End Property

Public Property Set Item(ByVal it As B154Item)
    Set mIt = it
End Property

Public Property Get Item() As B154Item
    Set Item = mIt
End Property
