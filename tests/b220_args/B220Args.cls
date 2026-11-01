VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B220Args"
Option Explicit

' Ledger 305 B219 (knife 220). The inbound IDispatch leg used to answer "which variant type is
' this?" in two half-places at once: the server adapter coerced only a hard-coded list
' (I2/I1/UI1/UI2/BOOL/EMPTY -> I4, R4 -> R8), and the generated argument readers then took the
' union member implied by the DECLARED parameter type without ever looking at the label.
' Measured on x64 and x86 with tests\tools\disp_probe.c (20-shot matrix, readings in ai/022):
'   String slot <- VT_I4 40 / VT_R8 / VT_BOOL -> dereferences a scalar as a BSTR -> native AV,
'                                                 the CLIENT process dies, no error at all;
'   Long slot   <- VT_BSTR "12" -> answered 1833297976 (the low 32 bits of the BSTR pointer),
'                                 silently; <- VT_R8 12.75 -> answered 0, silently.
' Now the server normalises every declared-by-value slot to its declared type through one owner
' (vb6_ComArgNormalize), and the readers keep reading the union member they always read.
' Unconvertible -> DISP_E_TYPEMISMATCH (0x80020005), which C3's own client already folds back
' into VB error 13.
'
' Readings come from the VTARG=<member>:<kind>[,<kind>...] arms of the probe:
'   i=VT_I4(40)  t=VT_I2(9)  u=VT_UI1(7)  f=VT_R4(12.75)  r=VT_R8(12.75)  b=VT_BOOL(True)
'   s=VT_BSTR("12")  x=VT_BSTR("abc")  e=VT_EMPTY  n=VT_NULL
' EchoVar's declared type is Variant, which is deliberately NOT normalised (the whole-slot arm
' is owned elsewhere) -- VarType() is the witness that those slots keep the caller's own label.

Public Function EchoLng(ByVal n As Long) As Long
    EchoLng = n
End Function

Public Function EchoStr(ByVal s As String) As String
    EchoStr = s
End Function

Public Function EchoDbl(ByVal d As Double) As Double
    EchoDbl = d
End Function

Public Function EchoBool(ByVal b As Boolean) As Long
    EchoBool = b
End Function

Public Function EchoVar(ByVal v As Variant) As Long
    EchoVar = VarType(v)
End Function

' two slots, two different declared arms: the question is per argument, not per call
Public Function SumTwo(ByVal a As Long, ByVal b As String) As String
    SumTwo = b & ":" & a
End Function
