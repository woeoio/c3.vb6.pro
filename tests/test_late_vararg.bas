' test_late_vararg.bas - C30-11 第一百零五刀 (§11.83): 「Variant 槽的实参位」
' 真机读数 (K 族探针, 真 VB6 量到的):
'   K1 Call GVar(o.NoSuchProperty)          = gHit=0        err=438  ← 被调方不跑
'   K2 Call GVar2(GFn(), o.NoSuchProperty)  = gHit=0 gN=1   err=438  ← 前面的实参照常求值
'   K5 Call GVarRef(o.NoSuchProperty)       = gHit=0        err=438  ← ByRef 槽同一条规则
' 与 Long 槽（I 族 / 第一百零二刀）同一条规则: 命中即整句中止。
' 聚合体落不进「标量临时量」那一档，但 C 里 vb6_VARIANT 是可声明的 ——
' 起手 vb6_VariantEmpty()、深拷贝 vb6_VariantCopy 拿自己那份所有权、调用后 vb6_VariantClear 归还。
Option Explicit

Public gHit As Long
Public gN As Long

Public Function GFn() As Long
    gN = gN + 1
    GFn = 5
End Function

Public Sub GVar(ByVal x As Variant)
    gHit = 1
End Sub

Public Sub GVar2(ByVal a As Long, ByVal b As Variant)
    gHit = 1
End Sub

Public Sub GVarRef(ByRef x As Variant)
    gHit = 1
End Sub

Public Sub VarArgByVal()
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    gHit = 0
    Call GVar(o.NoSuchProperty)
End Sub

Public Sub VarArgAfterEval()
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    gHit = 0: gN = 0
    Call GVar2(GFn(), o.NoSuchProperty)
End Sub

Public Sub VarArgByRef()
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    gHit = 0
    Call GVarRef(o.NoSuchProperty)
End Sub

Public Sub VarArgControl()
    gHit = 0: gN = 0
    Call GVar2(GFn(), 5)
End Sub
