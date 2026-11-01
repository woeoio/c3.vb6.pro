Attribute VB_Name = "mdlK175"
Option Explicit

' <vbeclipse> fixture driver (knife 175 / account 471 station D):
' a zero-arg method call written in STATEMENT position on a legacy
' interface VIEW. Before the fix this fell to vb6_ComCall(view, ...)
' = the whole struct into a void* slot = cl C2172 (24 sites in the
' external corpus).

Public Alpha As Long
Public Beta As Long

Public Sub Main()
    Dim x As IK175
    Set x = New K175Alpha
    x.Ping
    x.Ping
    Set x = New K175Beta
    x.Ping
    Dim t As Long
    t = x.Twice(5)
    ' 第二百一十五刀: 字段槽那一档的视图读 —— 同一格 m_lv 先装自家 (1100)、再装另一枚
    ' 实现类 (1200); 读数随**装进去的那一枚**变, 就是「按实现者选 wrap」那一步没走错。
    Dim h As K175Alpha
    Set h = New K175Alpha
    h.BindSelf
    Dim f1 As Long
    f1 = h.NumViaField
    h.BindBeta
    Dim f2 As Long
    f2 = h.NumViaField
    ' 第二百一十七刀: 同一格字段先**置空**、再重绑、再读 —— 读数回到 1100 才说明清零那一发
    ' 没把槽搞成装不回去的形状 (视图是借来的，置空不退旧，所以 Alpha 那枚计数不在这条判据里)。
    h.ClearAndRebind
    Dim f3 As Long
    f3 = h.NumViaField
    Debug.Print "K175-alpha="; Alpha
    Debug.Print "K175-beta="; Beta
    Debug.Print "K175-t="; t
    Debug.Print "K175-f="; f1
    Debug.Print "K175-fb="; f2
    Debug.Print "K175-fc="; f3
    Debug.Print "K175-DONE"
End Sub
