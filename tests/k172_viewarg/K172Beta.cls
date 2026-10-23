VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K172Beta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IK172

Private Function IK172_Value() As Long
    IK172_Value = 202
End Function

Public Function PushInto() As Long
    PushInto = mdlK172.ViewValue(Me)
End Function

Public Function BoxInto() As Long
    Dim v As IK172
    Set v = Me
    BoxInto = mdlK172.BoxValue(v)
End Function
