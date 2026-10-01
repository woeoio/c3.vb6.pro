VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VbView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Fix <vbeclipse> rev16 回归夹具: 真工程 VbEclipse/Classes/View.cls 的最小形态
' (Private m_ViewId As String + Private m_View As Object + 两个属性对)。
' 用途: 复现 "工程类实例 → List.Add(ByRef Item As Variant) 装箱 → Set l = Nothing
' → List.Item 取回 → 读属性" 这条链。真工程 play78 崩在这条链的下游:
' ucPerspective.CreateFolder 里 `l_ucFolder.AddView m_Views.Item(ViewId), False`
' 拿到 Nothing → ucFolder.View.ViewId 对 NULL 解引用 (av read 0x4)。

Private m_ViewId As String
Private m_View As Object

Public Property Get ViewId() As String
    ViewId = m_ViewId
End Property

Public Property Let ViewId(ByVal NewViewId As String)
    m_ViewId = NewViewId
End Property

Public Property Get View() As Object
    Set View = m_View
End Property

Public Property Set View(ByRef NewView As Object)
    Set m_View = NewView
End Property
