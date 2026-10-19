Option Explicit

' Ledger 305 B184 (knife 75). Two of the four receiver spellings need a procedure whose formal
' parameter is the class, so they live here: `h.ci = t` (class-typed parameter, the most common
' shape in real code) and `With h / .ci = 40000` (the With station emits its own field store and
' had never asked the question at all).

Private Const BV As String = "C3ProbeB184"

Public Sub WithIntegerOutOfRange(h As B184Holder)
    On Error GoTo O1
    With h
        .ci = 40000
    End With
    Debug.Print BV & "-P5-with-noerr"
    Exit Sub
O1:
    Debug.Print BV & "-P5-with-raised=" & Err.Number
End Sub

Public Sub ParamIntegerFromLong(h As B184Holder)
    Dim t As Long
    On Error GoTo O2
    t = 40000
    h.ci = t
    Debug.Print BV & "-P7-param-noerr"
    Exit Sub
O2:
    Debug.Print BV & "-P7-param-raised=" & Err.Number
End Sub
