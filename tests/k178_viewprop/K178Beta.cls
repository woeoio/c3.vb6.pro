VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K178Beta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IK178

Private Property Get IK178_Label() As String
    mdlK178.Beta = mdlK178.Beta + 1
    IK178_Label = "B" & mdlK178.Beta
End Property

Private Property Let IK178_Label(ByVal v As String)
    mdlK178.LetHits = mdlK178.LetHits + 10
End Property

Private Property Get IK178_Num() As Long
    mdlK178.BetaN = mdlK178.BetaN + 1
    IK178_Num = 500 + mdlK178.BetaN
End Property

Private Function IK178_Twice(ByVal n As Long) As Long
    IK178_Twice = 3 * n
End Function
