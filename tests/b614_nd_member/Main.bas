Attribute VB_Name = "Main"
Option Explicit

' 账 #614 (第352刀): 台账 §B190 剩下的两格
'   ① 多维类成员的元素访问 `o.M2(1, 2)` —— 修前发成 `o->M2(1, 2)` = 调用 vb6_SafeArray1D*
'      这个指针 ⇒ cl C2064, 整台 exe 编不出来 (UDT owner 那一臂早就对了)。
'   ② UDT 成员的整数组读 `w = h.Items` 的装箱档 —— 修前发 `vb6_VariantLong(h.Items)`
'      = 把描述符指针截进 int32 ⇒ VarType 答 3 而不是 8195 (数组)。
' 三枚对照组钉「没改过头」: UDT owner 的多维、类实例的整数组读、模块级数组。

Private Type TH2
    Items() As Long
    Bigs() As Long
End Type

Dim g() As Long

Sub Main()
    Dim o As K614
    Dim h As TH2
    Dim v As Variant
    Dim w As Variant

    Set o = New K614

    ' ① 多维: 类实例 owner (修前编不出来)
    ReDim o.M2(1 To 2, 1 To 3)
    o.M2(1, 2) = 42
    Debug.Print "614-cls-nd=[" & o.M2(1, 2) & "]"

    ' 对照组 A: 多维的 UDT owner (改前就对)
    ReDim h.Bigs(1 To 2, 1 To 3)
    h.Bigs(1, 2) = 43
    Debug.Print "614-udt-nd=[" & h.Bigs(1, 2) & "]"

    ' ② 整数组读进 Variant: UDT owner (修前 VarType 答 3)
    ReDim h.Items(1 To 2)
    h.Items(2) = 44
    w = h.Items
    Debug.Print "614-udt-arr-vt=[" & VarType(w) & "]"

    ' 对照组 B: 整数组读进 Variant 的类实例 owner (第351刀已修对)
    ReDim o.L(1 To 2)
    o.L(2) = 45
    v = o.L
    Debug.Print "614-cls-arr-vt=[" & VarType(v) & "]"

    ' 对照组 C: 一维类成员的元素访问 + 模块级数组
    Debug.Print "614-cls-1d=[" & o.L(2) & "]"
    ReDim g(1 To 2)
    g(2) = 46
    Debug.Print "614-ctl-mod=[" & VarType(g) & "]"

    Debug.Print "614-DONE"
End Sub
