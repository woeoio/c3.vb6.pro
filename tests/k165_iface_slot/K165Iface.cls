VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "K165Iface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' <vbeclipse> 第一百六十五刀的接口档: 抽象那一形 (只有声明、没有体),
' 且显式 VB_Creatable = True => 进不了 ivref 登记表, 走 legacy 的 knownIfaceVars_
' 名册 => 槽的 C 值是值结构体 vb6_iface_K165Iface { vtbl, obj }。
Public Function Bar(ByVal n As Long) As Long

End Function
