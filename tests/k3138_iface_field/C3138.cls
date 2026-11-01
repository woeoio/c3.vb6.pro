VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C3138"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' 持有者: 字段槽装的是**接口视图**。本刀之前它落进工程类那位主人, 发的是
' vb6_cls_IK3138_AddRef/_Release —— 接口档压根没有计数出口 = LNK2019。
Private m_iv As IK3138

Public Sub FromNew()
    Set m_iv = New H3138
End Sub

Public Sub FromNewB()
    Set m_iv = New H3138B
End Sub

Public Sub FromImpl()
    Set m_iv = mdl3138.g_impl
End Sub

Public Sub FromGlobalView()
    Set m_iv = mdl3138.g_iv
End Sub

Public Sub ClearIt()
    Set m_iv = Nothing
End Sub

Public Function TwiceViaField(ByVal n As Long) As Long
    TwiceViaField = m_iv.Twice(n)
End Function

Public Function KindViaField() As Long
    KindViaField = m_iv.Kind
End Function
