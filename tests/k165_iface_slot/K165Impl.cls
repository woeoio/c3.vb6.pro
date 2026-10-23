VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K165Impl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements K165Iface

Private Function K165Iface_Bar(ByVal n As Long) As Long
    K165Iface_Bar = n * 3
End Function

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
