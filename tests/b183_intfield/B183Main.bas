Attribute VB_Name = "B183Main"
Option Explicit

' Ledger 305 B183 (knife 74) driver. The lesion lives in the class module's own code (a bare
' class-field name as the left side of an assignment), so the stores below are calls into
' B183Holder; this file only seeds the fields, then reads them back from outside.
' The two heads of the judgment: P1/P2/P7 must raise error 6, and P3/P4 must show that the
' rejected store left the previous legal value alone. On BASE the three stores were accepted
' and both read-backs showed -25536.

Public Sub Main()
    Dim h As New B183Holder

    h.StoreLegal
    Debug.Print "C3ProbeB183-P0-seed=" & h.ci & "|" & h.ReadPf()

    h.StoreBareOutOfRange
    h.StoreFromLongOutOfRange
    h.StorePrivateOutOfRange

    Debug.Print "C3ProbeB183-P3-ci-kept=" & h.ci
    Debug.Print "C3ProbeB183-P4-pf-kept=" & h.ReadPf()

    h.StoreByteOutOfRange
    h.StoreLongInRange
    Debug.Print "C3ProbeB183-P8-cl=" & h.cl

    Debug.Print "C3ProbeB183-DONE"
End Sub
