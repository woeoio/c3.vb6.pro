Attribute VB_Name = "pdUse"
Option Explicit

Public Sub Main()
    Dim d As RectF, e As RectF
    ' 模块函数 → 类实例 → 零参方法返回 UDT: 整条链不写括号。
    If pdGeom.TakeRect(d, pdHolder.GetImg.GetBoundaryRectF) Then Debug.Print "ok"
    ' 跨模块 `Mod.Fn()` 带 Optional 形参、调用点省略全部实参。
    If pdGeom.TakeRect(d, Tools_Pencil.GetRect()) Then Debug.Print "ok2"
End Sub

' With 目标是"返回 UDT 的零参方法"(无括号): 右值必须先落临时再取址。
Public Sub TestWith(ByVal srcImg As pdImg)
    Dim r As RectF
    With srcImg.GetBoundaryRectF
        r.Left = .Left
        r.Top = .Top
    End With
    Debug.Print CStr(r.Left)
End Sub
