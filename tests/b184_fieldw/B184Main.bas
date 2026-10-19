Attribute VB_Name = "B184Main"
Option Explicit

' Ledger 305 B184 (knife 75) driver. The lesion is the four receiver spellings of a class-field
' store, so each one is driven from its own procedure (one error frame per store, so a raise
' cannot swallow the next case). The two heads of the judgment: every out-of-range store must
' raise error 6, and the field must still read back the value seeded before it. The Long,
' Boolean and property witnesses must stay exactly as they are on BASE.

Public Sub Main()
    Dim h As New B184Holder
    Dim p As New B184Peer

    h.SeedLegal
    Debug.Print "C3ProbeB184-P0-seed=" & h.ci & "|" & h.cb & "|" & h.cl

    h.MeIntegerOutOfRange
    Debug.Print "C3ProbeB184-P2-ci-kept=" & h.ci

    p.WithIntegerOutOfRange h
    Debug.Print "C3ProbeB184-P6-ci-kept2=" & h.ci

    h.MeByteOutOfRange
    Debug.Print "C3ProbeB184-P4-cb-kept=" & h.cb

    p.ParamIntegerFromLong h
    Debug.Print "C3ProbeB184-P8-ci-kept3=" & h.ci

    h.MeWitnessLong
    Debug.Print "C3ProbeB184-P9b-cl=" & h.cl

    h.MeWitnessBoolean
    h.MeFloatIntoInteger
    h.PropRoundTrip

    LocalStationOutOfRange

    Debug.Print "C3ProbeB184-DONE"
End Sub

' The receiver spelled as a plain local variable of the class type - the shape that real code
' uses most. The read-back is printed on both branches so BASE and NEW can be told apart even
' when the store was accepted.
Private Sub LocalStationOutOfRange()
    Dim k As New B184Holder
    On Error GoTo O1
    k.ci = 1111
    k.ci = 40000
    Debug.Print "C3ProbeB184-P13-local-noerr=" & k.ci
    Exit Sub
O1:
    Debug.Print "C3ProbeB184-P13-local-raised=" & Err.Number & "|" & k.ci
End Sub
