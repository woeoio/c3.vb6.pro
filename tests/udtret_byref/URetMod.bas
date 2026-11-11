Attribute VB_Name = "URetMod"
Option Explicit

Public Type URect
    a As Single
    b As Single
End Type

' 修前: inferUdtTypeOfExpr 不认 `obj.Method()` 的 UDT 返回 ⇒ intptr_t 兜底
' `intptr_t _vb6_anytmp0 = (intptr_t)(vb6_UPath_GetRect(p))` C2440。
Public Sub Main()
    Dim b As UBrush
    Dim p As UPath
    Set b = New UBrush
    Set p = New UPath
    b.SetRect p.GetRect()
End Sub
