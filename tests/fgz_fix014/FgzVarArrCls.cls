VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FgzVarArrCls"
Option Explicit

' Account #610 knife #344: the SAFEARRAY(VARIANT) inbound slot.
'
' The client hands over a real SAFEARRAY(VARIANT) (see tests/tools/disp_probe.c,
' ARRVAR= arm). The server-side inbound path is vb6_VariantBorrowIntoSlot ->
' vb6_VariantArrayIntoSlot, which used to answer Empty for the VT_ARRAY|VT_VARIANT
' shape: VARIANT elements are not blittable, so the array-level memcpy was refused
' and the whole array was dropped.
'
' The answer string is the witness. Before the fix the slot is Empty, so v(0) reads
' back as Empty and the answer says so. After the fix the elements arrive one by one
' (a deep copy each) and the answer carries their real values.
'
' NOTE: the element read arm is vb6_VariantArrayGet(&v, 0) -- the same arm the
' upstream GetXData family needs ("read the Variant out-param, then take (0)/(1)").

Public Function TakeVarArr(ByRef v As Variant) As String
    Dim s As String
    s = "n=" & CStr(UBound(v) + 1)
    s = s & ";t0=" & TypeName(v(0)) & ";v0=" & CStr(v(0))
    s = s & ";t1=" & TypeName(v(1)) & ";v1=" & CStr(v(1))
    s = s & ";t2=" & TypeName(v(2)) & ";v2=" & CStr(v(2))
    s = s & ";t3=" & TypeName(v(3)) & ";v3=" & CStr(v(3))
    TakeVarArr = s
End Function
