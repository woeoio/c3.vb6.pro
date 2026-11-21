' test_late_arg.bas - C30-11 第一百零二刀 (§11.80): 「实参位上的晚绑定缺成员读」
' 真机读数 (I 族探针, 真 VB6 量到的):
'   I1 Call GTouch2(o.NoSuchProperty, GFn()) = gHit=0 gN=0   ← 坏实参之后的实参不再求值
'   I2 Call GTouch2(GFn(), o.NoSuchProperty) = gHit=0 gN=1   ← 坏实参之前的照常求值
'   I3 Call o.NoSuchMethod(GFn())            = gN=1         ← 实参先于成员解析
' 这份夹具只负责让这些形状出现在发射里（形状针钉发码形、运行时读数在 CI 的 VBP 面）。
Option Explicit

Public gHit As Long
Public gN As Long

Public Function GFn() As Long
    gN = gN + 1
    GFn = 5
End Function

Public Sub GTouch(ByVal x As Long)
    gHit = 1
End Sub

Public Sub GTouch2(ByVal a As Long, ByVal b As Long)
    gHit = 1
End Sub

Sub Main()
    Dim o As Object
    Dim n As Long
    Set o = CreateObject("Scripting.Dictionary")
    Call GTouch(o.NoSuchProperty)
    Call GTouch2(o.NoSuchProperty, GFn())
    Call GTouch2(GFn(), o.NoSuchProperty)
    Call o.NoSuchMethod(GFn())
    n = o.NoSuchProperty(GFn())
End Sub
