' test_late_varslot.bas - C30-11 第一百零六刀 (§11.84): 「Variant 槽的赋值位」
' 真机读数 (K4 那一格, 真 VB6 量到的):
'   v = 111 : Err.Clear : v = o.NoSuchProperty  = v=111  err=438  ← 目标一个字都不写
' 与 I 族 (Long 槽) / F 族 (String 槽) 同一条规则: 命中即整句中止, 目标保持原值;
' 差别只在「聚合体的落点」—— 标量走 `_t` 临时量, Variant 目标走 `_av` (Empty 起手 + 浅搬)。
' 控制格: 存在的属性读必须照样写进目标 (L2/L4), 否则守卫会把正常赋值也挡掉。
Option Explicit

Public gN As Long

Public Sub VarSlotAssignMissing()
    Dim o As Object
    Dim v As Variant
    Set o = CreateObject("Scripting.Dictionary")
    v = 111
    v = o.NoSuchProperty
End Sub

Public Sub VarSlotAssignPresent()
    Dim o As Object
    Dim v As Variant
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"
    v = 111
    v = o.Count
End Sub

Public Sub VarSlotAssignSuccessThenMissing()
    Dim o As Object
    Dim v As Variant
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"
    v = 111
    v = o.Count
    v = o.NoSuchProperty
End Sub

Public Sub VarSlotAssignString()
    Dim o As Object
    Dim v As Variant
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"
    v = 111
    v = o.Item("k")
End Sub
