Attribute VB_Name = "TestVbaIntQual"
Option Explicit

' 账 #478: `VBA.Int` / `VBA.Fix` 是内建函数的**限定写法**（真 VB6 6.00.9782 探针:
' 两者都编得过、取值正确；`Math.Int` / `Math.Fix` 则是编译期「未找到方法或数据成员」，
' 负控在 test_vbamath_members_bad.bas）。
'
' 修前 `VBA.Int` 发 `vb6_vb6_Int` ⇒ LNK2019：`Int` 是 C 关键字，cIdent 先转成 `vb6_Int`，
' 发码点又拼一次 `vb6_` 前缀。`VBA.Fix` 只是恰好因为 "Fix" 不是关键字才名对。
'
' 取值口径（同一支真 VB6 探针, .temp/vbaint/vb6_readings.txt）:
'   Int(1.7)=1   Int(-1.7)=-2   ← floor, 向 -∞ 取整
'   Fix(1.7)=1   Fix(-1.7)=-1   ← trunc, 向 0 取整
' 注意**正数方向 Int 与 Fix 相同**（都是 1）; 两者只在负数上分界, 所以分界断言必须用负例。

Public Sub Main()
    Debug.Print "VI-POS=" & VBA.Int(1.7)
    Debug.Print "VI-NEG=" & VBA.Int(-1.7)
    Debug.Print "VF-POS=" & VBA.Fix(1.7)
    Debug.Print "VF-NEG=" & VBA.Fix(-1.7)
    ' 裸名与限定名必须同值（同一份 RTL 实现）。
    Debug.Print "BI-NEG=" & Int(-1.7)
    Debug.Print "BF-NEG=" & Fix(-1.7)
    ' 正向对照: 同一前缀下的其他内建照旧走折叠。
    Debug.Print "VA-ABS=" & VBA.Abs(-3)
    Debug.Print "VA-PASS"
End Sub
