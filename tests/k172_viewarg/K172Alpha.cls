VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K172Alpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IK172

Private Function IK172_Value() As Long
    IK172_Value = 101
End Function

' 站B: 把 Me (vb6_cls_K172Alpha*) 递进一枚视图形参 —— Ray.vbp 那四发的同形。
Public Function PushInto() As Long
    PushInto = mdlK172.ViewValue(Me)
End Function

' 站E: 一枚视图槽递进 ByRef Variant 形参 —— RayTrace.c(158) 的同形。
Public Function BoxInto() As Long
    Dim v As IK172
    Set v = Me
    BoxInto = mdlK172.BoxValue(v)
End Function
