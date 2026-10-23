VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K177Beta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IK177

Private Sub IK177_Ping()
    mdlK177.Beta = mdlK177.Beta + 10
End Sub

Private Function IK177_Twice(ByVal n As Long) As Long
    IK177_Twice = 3 * n
End Function
