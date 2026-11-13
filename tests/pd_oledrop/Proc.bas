Attribute VB_Name = "Proc"
Option Explicit

Public Const vbOLEDropNone = 0
Public Const vbOLEDropManual = 1

Public Sub EnableDrop(ByVal enable As Boolean)
    If enable Then
        FormMain.OLEDropMode = vbOLEDropManual
    Else
        FormMain.OLEDropMode = vbOLEDropNone
    End If
End Sub
