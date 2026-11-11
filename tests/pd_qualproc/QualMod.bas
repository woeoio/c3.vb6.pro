Attribute VB_Name = "QualMod"
Option Explicit

Public Sub Main()
    Dim r1 As RectF, dst As RectF
    Dim ok As Boolean
    r1.Left = 1
    ok = GDI_Plus.IntersectRectF(dst, r1, Tools_Pencil.GetModifiedUpdateRectF())
    ok = GDI_Plus.IntersectRectF(dst, r1, Tools_Clone.GetModifiedUpdateRectF())
    If ok Then Debug.Print "qualproc ok " & CStr(dst.Left)
End Sub
