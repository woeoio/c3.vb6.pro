VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B162Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub Bump(ByRef x As String)
    x = x & "-ok"
End Sub

Public Sub Swap(ByRef x As String)
    x = "new"
End Sub

Public Sub Clear(ByRef x As String)
    x = ""
End Sub

Public Sub NoChange(ByRef x As String)
    x = x
End Sub

Public Sub BumpVal(ByVal x As String)
    x = x & "-byval"
End Sub

Public Sub BumpL(ByRef x As Long)
    x = x + 1
End Sub
