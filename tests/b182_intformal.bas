' Ledger B182 (knife 73) regression fixture. Assigning a value outside the 16-bit range into an
' Integer variable is VB6 error 6 (Overflow), and this compiler already enforced that for
' module-level and local Integers by routing the store through vb6_ChkInt. Integer FORMALS never
' reached that check: the three formal arms fold Integer into the Long name table only, so the
' narrow table stayed empty for them and the store went out unchecked - silently truncating, and
' through a ByRef parameter, writing the truncated value back into the CALLER's variable.
' BASE readings were: P2 printed noerr (no raise at all) and P6 read back the truncated number.
' After the fix the three paths answer the same way. Readings must be identical on x64 and x86.
Option Explicit

Private Const BV_APP As String = "C3ProbeB182"

Private gi As Integer

Sub SetModuleOutOfRange()
    On Error GoTo MH
    gi = 40000
    Debug.Print BV_APP & "-P3-module-noerr"
    Exit Sub
MH:
    Debug.Print BV_APP & "-P3-module-raised=" & Err.Number
End Sub

Sub SetLocalOutOfRange()
    Dim li As Integer
    li = 1
    On Error GoTo LH
    li = 40000
    Debug.Print BV_APP & "-P1-local-noerr"
    Exit Sub
LH:
    Debug.Print BV_APP & "-P1-local-raised=" & Err.Number
End Sub

Sub SetFormalOutOfRange(ByRef fi As Integer)
    On Error GoTo FH
    fi = 40000
    Debug.Print BV_APP & "-P2-formal-noerr"
    Exit Sub
FH:
    Debug.Print BV_APP & "-P2-formal-raised=" & Err.Number
End Sub

Sub SetFormalInRange(ByRef fi As Integer)
    On Error GoTo RH
    fi = 7
    Debug.Print BV_APP & "-P5-formal-inrange-ok"
    Exit Sub
RH:
    Debug.Print BV_APP & "-P5-formal-raised=" & Err.Number
End Sub

Sub BumpByValFormal(ByVal fv As Integer)
    On Error GoTo VH
    fv = fv + 1
    Debug.Print BV_APP & "-P4-byval-inrange-ok=" & fv
    Exit Sub
VH:
    Debug.Print BV_APP & "-P4-byval-raised=" & Err.Number
End Sub

Sub SetLongFormalInRange(ByRef lg As Long)
    On Error GoTo LH9
    lg = 40000
    Debug.Print BV_APP & "-P9-long-formal-ok=" & lg
    Exit Sub
LH9:
    Debug.Print BV_APP & "-P9-long-formal-raised=" & Err.Number
End Sub

Sub Main()
    Dim sink As Integer
    Dim untouched As Integer
    sink = 1
    untouched = 2
    Dim lng As Long
    SetLocalOutOfRange
    SetFormalOutOfRange sink
    SetModuleOutOfRange
    BumpByValFormal 5
    ' P6: the out-of-range store must not land in the caller's variable. On BASE it did, because
    ' nothing checked, so this reads back a truncated number instead of the 1 that was put in.
    Debug.Print BV_APP & "-P6-sink-after-outofrange=" & sink
    SetFormalInRange sink
    Debug.Print BV_APP & "-P7-sink-after-inrange=" & sink
    SetLongFormalInRange lng
    Debug.Print BV_APP & "-P8-untouched=" & untouched
    Debug.Print BV_APP & "-DONE"
End Sub
