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

' <vbeclipse> 账 #305 第二百一十五刀: 一枚**字段槽**装的是一枚 legacy 接口视图
' (`Private m_lv As IK175` 的 C 字段型 = 值结构体 vb6_iface_IK175 { vtbl, obj })。
' 三格改前各撞一张脸: 工厂初始化发 `me->m_lv = NULL` = cl C2440 / Set 装填裸拷贝实例
' 指针 = cl C2440 / 属性读把整枚结构体递进 vb6_ComGetIntProp 首槽 = cl C2172。
' BindSelf 与 BindBeta 各填 1100 / 1200 —— 同一格字段换一枚实现者读数就换，
' 认错 wrap 时这两条必有一条不是它自己的数。
Private m_lv As IK175

Public Sub BindSelf()
    Set m_lv = New K175Alpha
End Sub

Public Sub BindBeta()
    Set m_lv = New K175Beta
End Sub

' 第二百一十七刀: 字段槽的**置空** —— 那一格装的是值结构体 vb6_iface_IK175，没有 NULL 这一形。
' 改前发 `vb6_ReleaseObject((void**)&me->m_lv)` = 把整枚结构体的地址交给一枚 IUnknown** 槽；
' 置空之后立刻重绑再读 (VB6 里读一枚已置空的接口变量报 91，C3 今天对那一形没有任何口径，
' 1421 份语料里也 0 处可达 ⇒ 本刀不为它发明错误档，只让「清完还能装回去」有读数)。
Public Sub ClearAndRebind()
    Set m_lv = Nothing
    Set m_lv = New K175Alpha
End Sub

Public Function NumViaField() As Long
    NumViaField = m_lv.Num
End Function

Private Property Get IK175_Num() As Long
    IK175_Num = 1100
End Property

Private Sub IK175_Ping()
    mdlK175.Alpha = mdlK175.Alpha + 1
End Sub

Private Function IK175_Twice(ByVal n As Long) As Long
    IK175_Twice = 2 * n
End Function
