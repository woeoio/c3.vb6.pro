VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IK3138"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' 账 #305 第二百零三刀靶子: 工程内接口的**字段槽** (类字段)。两枚实现者让
' 「视图槽打给谁」可判 —— 槽里存的是 &实例->__iv_IK3138, 打错人 = 另一枚计数器涨。
Public Function Twice(ByVal n As Long) As Long

End Function

Public Property Get Kind() As Long

End Property
