Attribute VB_Name = "CxUse"
Option Explicit

Public Sub Main()
    Dim a As Cpx, b As Cpx, s As Cpx
    s = CxMod.c_mul(CxMod.c_add(a, b), b)
    s = CxMod.c_add(b, CxMod.c_mul(a, b))
    Debug.Print s.c_real
End Sub
