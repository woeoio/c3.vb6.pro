Attribute VB_Name = "CxMod"
Option Explicit

Public Type Cpx
    c_real As Double
    c_imag As Double
End Type

Public Function c_add(ByRef w As Cpx, ByRef z As Cpx) As Cpx
    c_add.c_real = w.c_real + z.c_real
    c_add.c_imag = w.c_imag + z.c_imag
End Function

Public Function c_mul(ByRef w As Cpx, ByRef z As Cpx) As Cpx
    c_mul.c_real = w.c_real * z.c_real - w.c_imag * z.c_imag
    c_mul.c_imag = w.c_real * z.c_imag + w.c_imag * z.c_real
End Function
