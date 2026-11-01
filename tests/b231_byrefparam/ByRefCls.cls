VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ByRefCls"
Option Explicit

' Knife 231 measure: the scalar ByRef slot. A VB6 client sends VT_BYREF|VT_I4 pointing at its
' own LONG. The generated read arm is `&((VARIANT*)args[0])->lVal`, i.e. it hands the callee the
' ADDRESS OF THE SLOT'S UNION -- so the body reads/writes the slot copy, never the client long.

Public Function Bump(ByRef v As Long) As Long
    v = v + 1
    Bump = v
End Function
