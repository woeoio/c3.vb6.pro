' test_late_barecond.bas - C30-11 第一百零四刀 (§11.82): 「裸成员做循环条件」
' 真机读数 (J 族探针, 真 VB6 量到的):
'   J9  Do Until o.NoSuchProperty          = n=4  err=438  ← 报错 ⇒ 前测按假 ⇒ 进体
'   J10 Do ... Loop Until o.NoSuchProperty = n=1  err=438  ← 报错 ⇒ 出环
' 修前这一形会被**下一条语句**收尾: `gN = gN + 1` 静默变成 `o.NoSuchProperty = gN + 1`,
' `GTouch 1` 变成 `o.NoSuchProperty(1)` —— 读变写、语句消失。
' 这份夹具只负责让四站的发码形出现在发射里（形状针钉形, 运行时读数在 VBP 面）。
Option Explicit

Public gHit As Long
Public gN As Long

Public Sub GTouch(ByVal x As Long)
    gHit = gHit + 1
End Sub

Public Sub BareDoUntil()
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    gN = 0
    Do Until o.NoSuchProperty
        gN = gN + 1
        If gN > 3 Then Exit Sub
    Loop
End Sub

Public Sub BareWhileWend()
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    gN = 0
    While o.NoSuchProperty
        gN = gN + 1
        If gN > 3 Then Exit Sub
    Wend
End Sub

Public Sub BareLoopUntil()
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    gN = 0
    Do
        gN = gN + 1
        If gN > 3 Then Exit Sub
    Loop Until o.NoSuchProperty
End Sub

Public Sub BareLoopWhile()
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    gN = 0
    Do
        gN = gN + 1
        If gN > 3 Then Exit Sub
    Loop While o.NoSuchProperty
End Sub
