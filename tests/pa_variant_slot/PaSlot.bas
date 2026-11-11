Attribute VB_Name = "PaSlot"
Option Explicit

' §56.28: ParamArray 元素回写 (`vals(i) = ...`) 走 `vb6_PA_SetVariant`
' 那一站。承接临时必须是 vb6_VARIANT —— boxToVariant 的发码
' `vb6_VariantFromValue(...)` / `vb6_VariantBool(...)` 都交 vb6_VARIANT,
' 声明成 Win32 VARIANT 就是 `initializing` C2440 (PhotoDemon PDMath.c 825…).
Public Sub Bump(ParamArray vals() As Variant)
    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        vals(i) = CInt(vals(i)) + 1
    Next i
End Sub

Public Sub Main()
    Dim a As Long, b As Long
    a = 1
    b = 2
    Bump a, b
    Debug.Print a, b
End Sub
