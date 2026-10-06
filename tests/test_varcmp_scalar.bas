Option Explicit

' ============================================================
'  test_varcmp_scalar.bas - <vbeclipse> 238: the scalar operand of a Variant
'  comparison must be BOXED, not handed over by address.
'
'  Root cause shape: vb6_VarCmp*(&A, &B) decided "can I take this operand's
'  address" from the NAME (a bare identifier => &name).  A `Dim d As Double`
'  therefore reached the RTL as a vb6_VARIANT*, and the RTL read a VARIANT's
'  layout out of an 8-byte scalar: two equal numbers answered False (measured
'  on the pre-fix compiler, both arches), and the read itself ran past the
'  object.  The single answer now lives in CCodeGen::cmpOperandMayTakeAddr.
'
'  Two-sided on purpose: every equality judge has a near-miss partner that must
'  be False, otherwise "always True" would also pass.  The Long and String legs
'  pin the two routes that already worked (the VarCmpLong shortcut and the
'  vb6_StrCmp path) so a fix that re-routes them goes red instead of green.
' ============================================================

Dim gV As Variant
Dim gD As Double

Sub Main()
    Dim v As Variant
    Dim d As Double
    Dim sg As Single
    Dim l As Long
    Dim s As String
    Dim cur As Currency
    Dim dt As Date
    Dim a1 As Boolean
    Dim a2 As Boolean
    Dim a3 As Boolean
    Dim a4 As Boolean
    Dim a5 As Boolean
    Dim a6 As Boolean
    Dim a7 As Boolean
    Dim a8 As Boolean
    Dim a9 As Boolean
    Dim a10 As Boolean

    v = 240
    d = 240
    sg = 240
    l = 240
    s = "x"
    cur = 240
    dt = 1#
    gV = 7
    gD = 7

    ' --- the measured defect: Variant vs a scalar STORAGE (address must not be taken)
    v = 241
    a4 = ((v - 1) = cur)
    v = 240
    a1 = (v = d)
    a2 = (d = v)
    a3 = (v = sg)
    a5 = (gV = gD)
    Debug.Print "VC01-var-eq-double=" & CStr(a1)
    Debug.Print "VC02-double-eq-var=" & CStr(a2)
    Debug.Print "VC03-var-eq-single=" & CStr(a3)
    Debug.Print "VC04-var-eq-currency=" & CStr(a4)
    Debug.Print "VC05-module-var-eq-double=" & CStr(a5)

    ' --- near misses: the same shapes must answer False (kills "always True")
    d = 241
    Debug.Print "VC06-var-ne-double-true=" & CStr((v <> d))
    Debug.Print "VC07-var-eq-double-false=" & CStr(Not (v = d))
    d = 240

    ' --- relational legs, both operand orders
    v = 241
    Debug.Print "VC08-var-gt-double=" & CStr(v > d)
    Debug.Print "VC09-double-lt-var=" & CStr(d < v)
    v = 239
    Debug.Print "VC10-rel-reverse-false=" & CStr(Not ((v > d) Or (d < v)))

    ' --- the two routes that already worked, pinned so a re-route shows up
    v = 240
    Debug.Print "VC11-var-eq-long=" & CStr(v = l)
    Debug.Print "VC12-var-eq-long-false=" & CStr(Not (v = (l + 1)))
    v = "x"
    Debug.Print "VC13-var-eq-string=" & CStr(v = s)
    Debug.Print "VC14-var-eq-string-false=" & CStr(Not (v = "xy"))

    ' --- a Date scalar and an uninitialised Variant: VB6 says Empty <> 0-date here
    v = 1
    Debug.Print "VC15-var-eq-date=" & CStr(v = dt)
    Debug.Print "VC16-var-eq-date-false=" & CStr(Not (v = (dt + 1)))

    Debug.Print "VC-DONE"
End Sub
