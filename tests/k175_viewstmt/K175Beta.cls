VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K175Beta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IK175

Private Property Get IK175_Num() As Long
    IK175_Num = 1200
End Property

Private Sub IK175_Ping()
    mdlK175.Beta = mdlK175.Beta + 10
End Sub

Private Function IK175_Twice(ByVal n As Long) As Long
    IK175_Twice = 3 * n
End Function
