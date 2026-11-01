VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "H3138B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IK3138

Private Sub Class_Initialize()
    mdl3138.g_initB = mdl3138.g_initB + 1
End Sub

Private Sub Class_Terminate()
    mdl3138.g_termB = mdl3138.g_termB + 1
End Sub

Private Function IK3138_Twice(ByVal n As Long) As Long
    IK3138_Twice = 2 * n + 200
End Function

Private Property Get IK3138_Kind() As Long
    IK3138_Kind = 1200
End Property
