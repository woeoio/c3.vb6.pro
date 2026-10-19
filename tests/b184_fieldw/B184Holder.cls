Option Explicit

' Ledger 305 B184 (knife 75). A class-field store whose RECEIVER is spelled out (h.ci, Me.ci,
' With h / .ci, or a class-typed formal parameter) never asked the narrow-range question: the
' member arm of the "how narrow is this lvalue" oracle only recognised UDT symbols, so all four
' spellings went out as a plain C store. 40000 landed in an int16_t field as -25536 and no error
' was raised; real VB6 raises error 6 and keeps the previous value. Knife 74 (ledger B183) had
' already fixed the bare-name spelling inside the class's own code, which is why the same value
' written as `ci = 40000` already raised while `Me.ci = 40000` did not.
' BASE readings: P1/P3/P5/P7/P13 printed noerr, ci read back -25536, cb read back 44, and
' P12 read back 6 (C truncation where VB6 rounds). The Long / Boolean witnesses are here to
' show this knife did not drag a wider tier into the new gate.
' Readings must be identical on x64 and x86.

Private Const BV As String = "C3ProbeB184"

Public ci As Integer
Public cb As Byte
Public cl As Long
Public bl As Boolean
Public cs As Single
Private m_pi As Integer

' Witness for the veto the new gate carries: this property name and the bare projection of the
' private field above (m_pi) are the same key, so "is h.Pi a data field" must be answered by the
' method/property table first. A store that reaches the property setter still round-trips.
Public Property Get Pi() As Integer
    Pi = m_pi
End Property

Public Property Let Pi(ByVal v As Integer)
    m_pi = v
End Property

Public Sub PropRoundTrip()
    Dim k As New B184Holder
    k.Pi = 7
    Debug.Print BV & "-P15-prop=" & k.Pi
End Sub


Public Sub SeedLegal()
    ci = 12345
    cb = 200
    cl = 7
    cs = 1
End Sub

Public Sub MeIntegerOutOfRange()
    On Error GoTo O1
    Me.ci = 40000
    Debug.Print BV & "-P1-me-noerr"
    Exit Sub
O1:
    Debug.Print BV & "-P1-me-raised=" & Err.Number
End Sub

Public Sub MeByteOutOfRange()
    On Error GoTo O2
    Me.cb = 300
    Debug.Print BV & "-P3-mebyte-noerr"
    Exit Sub
O2:
    Debug.Print BV & "-P3-mebyte-raised=" & Err.Number
End Sub

Public Sub MeFloatIntoInteger()
    Dim d As Double
    d = 6.73
    Me.ci = d
    Debug.Print BV & "-P12-float-round=" & Me.ci
End Sub

' Witness: a Long field is not narrow. 40000 must still land unchecked.
Public Sub MeWitnessLong()
    On Error GoTo O3
    Me.cl = 40000
    Debug.Print BV & "-P9-long-noerr"
    Exit Sub
O3:
    Debug.Print BV & "-P9-long-raised=" & Err.Number
End Sub

' Witness: Boolean is normalised, not truncated, and stays outside the gate.
Public Sub MeWitnessBoolean()
    On Error GoTo O4
    Me.bl = 300
    Debug.Print BV & "-P10-bool-noerr"
    Exit Sub
O4:
    Debug.Print BV & "-P10-bool-raised=" & Err.Number
End Sub
