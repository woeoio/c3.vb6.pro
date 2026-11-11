Attribute VB_Name = "Use"
Option Explicit

Public Sub Main()
    Dim a As Cpx, b As Cpx, s As Cpx
    s = c_mul(c_add(a, b), b)
    Debug.Print s.c_real
End Sub
