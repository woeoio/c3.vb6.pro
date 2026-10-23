VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K175Alpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IK175

Private Sub IK175_Ping()
    mdlK175.Alpha = mdlK175.Alpha + 1
End Sub

Private Function IK175_Twice(ByVal n As Long) As Long
    IK175_Twice = 2 * n
End Function
