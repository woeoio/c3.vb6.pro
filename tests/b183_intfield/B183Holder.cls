Option Explicit

' Ledger 305 B183 (knife 74): a bare class-field name on the LEFT of an assignment never reached
' the narrow-range check. The class-member scan folds Long, Integer and Boolean together into the
' widened name table, and that table is the only thing a procedure entry restores, so inside the
' class's own code `ci = 40000` answered "this target is 32 bits wide" and went out as a plain
' store: no error 6, and the int16 field silently kept the truncated number. Byte fields and
' module-level Integers were already routed through their own narrow table, which is why the two
' witnesses below still raise. Readings must be identical on x64 and x86.
' BASE readings: P1/P2/P7 printed noerr, and the field read back -25536 (= 40000 mod 65536).

Private Const BV_APP As String = "C3ProbeB183"

Public ci As Integer
Private pf As Integer
Public cb As Byte
Public cl As Long

Public Sub StoreLegal()
    ci = 12345
    pf = 23456
End Sub

Public Sub StoreBareOutOfRange()
    On Error GoTo OE1
    ci = 40000
    Debug.Print BV_APP & "-P1-bare-noerr"
    Exit Sub
OE1:
    Debug.Print BV_APP & "-P1-bare-raised=" & Err.Number
End Sub

Public Sub StoreFromLongOutOfRange()
    Dim t As Long
    On Error GoTo OE2
    t = 40000
    ci = t
    Debug.Print BV_APP & "-P2-fromlong-noerr"
    Exit Sub
OE2:
    Debug.Print BV_APP & "-P2-fromlong-raised=" & Err.Number
End Sub

Public Sub StorePrivateOutOfRange()
    On Error GoTo OE3
    pf = 40000
    Debug.Print BV_APP & "-P7-private-noerr"
    Exit Sub
OE3:
    Debug.Print BV_APP & "-P7-private-raised=" & Err.Number
End Sub

' Witness: the Byte leg of the same scan was already registered in its own narrow table
' (ledger 123), so this one raises on BASE too. It is here to show this knife did not widen
' or move that answer.
Public Sub StoreByteOutOfRange()
    On Error GoTo OE4
    cb = 300
    Debug.Print BV_APP & "-P5-byte-noerr"
    Exit Sub
OE4:
    Debug.Print BV_APP & "-P5-byte-raised=" & Err.Number
End Sub

' Witness: a Long field is not narrow, so 40000 must still land unchecked. Answering "every
' class field is narrow" would turn this into a spurious error 6.
Public Sub StoreLongInRange()
    On Error GoTo OE5
    cl = 40000
    Debug.Print BV_APP & "-P6-long-noerr"
    Exit Sub
OE5:
    Debug.Print BV_APP & "-P6-long-raised=" & Err.Number
End Sub

Public Function ReadPf() As Integer
    ReadPf = pf
End Function
