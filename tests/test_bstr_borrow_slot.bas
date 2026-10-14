Option Explicit

' Knife 57 (ledger 305, sec B167): boxing a BORROWED BSTR into a Variant slot used to hand
' the slot the source's own buffer (vb6_VariantString only stamps vtBSTR, it neither copies
' nor AddRefs). Then either side releasing/reassigning makes the other one read freed memory.
' Every window below writes a string into a slot, then clobbers the OTHER side, then reads
' back what must still be there. Expected values are the ORIGINAL strings.

Private Type Rec
    Name As String
    Fv As Variant
End Type

Private gModStr As String
Private gModVar As Variant
Private gStrArr(1) As String

Function LabelOf() As Variant
    LabelOf = gModStr          ' common VB6 shape: a Variant return slot takes a live string
End Function

Sub TakesParam(ByVal pstr As String)
    Dim ps As Variant
    ps = pstr                  ' 6: formal String parameter -> local slot
    pstr = "clobbered"
    Debug.Print "B167-W6=" & CStr(ps) & "|" & Len(ps)
End Sub

Sub Main()
    Dim s As String
    Dim sA As Variant
    Dim sB As Variant
    Dim rec As Rec
    Dim varArr(1) As Variant
    Dim strArr(1) As String
    Dim got As Variant
    Dim i As Long

    ' 1: local slot, then the source variable is reassigned
    s = "hello"
    sA = s
    s = "world"
    Debug.Print "B167-W1=" & CStr(sA) & "|" & Len(sA)

    ' 2: slot owns what it took, so clearing it must not kill the source
    s = "sharedbuf"
    sA = s
    sA = 1
    Debug.Print "B167-W2=" & s & "|" & Len(s)

    ' 3: module-level slot
    s = "modslot"
    gModVar = s
    s = "gone"
    Debug.Print "B167-W3=" & CStr(gModVar) & "|" & Len(gModVar)

    ' 4: UDT Variant field (this station already asked - kept as the agreeing witness)
    s = "udtfield"
    rec.Fv = s
    s = "gone"
    Debug.Print "B167-W4=" & CStr(rec.Fv) & "|" & Len(rec.Fv)

    ' 5: Variant array element (same)
    s = "arrelem"
    varArr(0) = s
    s = "gone"
    Debug.Print "B167-W5=" & CStr(varArr(0)) & "|" & Len(varArr(0))

    TakesParam "byvalstr"

    ' 7: churn - the double-ownership direction. BASE frees the same block twice per round.
    For i = 1 To 200
        s = "churn" & i
        sA = s
        sB = s
        sA = 1
        sB = 2
    Next
    Debug.Print "B167-W7=" & s & "|" & Len(s)

    ' 8/9 (knife 58): the right-hand side here is a CALL, but these two RTL exports hand
    ' back somebody else's live buffer (vb6_ErrDescription -> vb6_err.description,
    ' vb6_ErrSource -> vb6_err.source; neither copies nor frees). Before the exception
    ' list the slot took that pointer as its own, so clearing the slot freed Err's block.
    On Error Resume Next
    Err.Raise 5, "errsrcw9", "errdescw8"
    sA = Err.Description
    Debug.Print "B167-W8a=" & CStr(sA) & "|" & Len(sA)
    sA = 1
    Debug.Print "B167-W8b=" & CStr(Err.Description) & "|" & Len(Err.Description)
    sB = Err.Source
    Debug.Print "B167-W9a=" & CStr(sB) & "|" & Len(sB)
    sB = 1
    Debug.Print "B167-W9b=" & CStr(Err.Source) & "|" & Len(Err.Source)

    ' X1: UDT String field as the source
    rec.Name = "udtstr"
    sA = rec.Name
    rec.Name = "gone"
    Debug.Print "B167-X1=" & CStr(sA) & "|" & Len(sA)

    ' X2: local String array element as the source (the _Generic road)
    strArr(0) = "arrstr"
    sA = strArr(0)
    strArr(0) = "gone"
    Debug.Print "B167-X2=" & CStr(sA) & "|" & Len(sA)

    ' X3: module String array element as the source (the inline road)
    gStrArr(0) = "modarr"
    sA = gStrArr(0)
    gStrArr(0) = "gone"
    Debug.Print "B167-X3=" & CStr(sA) & "|" & Len(sA)

    ' X4: a freshly minted string must still be taken over, not copied twice
    s = "abcdefgh"
    sA = Left(s, 3)
    s = "zzz"
    Debug.Print "B167-X4=" & CStr(sA) & "|" & Len(sA)

    ' X5: slot to slot (knife 55's shape, must stay green)
    sA = "slotone"
    sB = sA
    sA = 2
    Debug.Print "B167-X5=" & CStr(sB) & "|" & Len(sB)

    ' X6: through a Variant return slot
    gModStr = "fromfunc"
    got = LabelOf()
    gModStr = "gone"
    Debug.Print "B167-X6=" & CStr(got) & "|" & Len(got)

    ' X7: literal control
    sA = "lit"
    Debug.Print "B167-X7=" & CStr(sA) & "|" & Len(sA)

    Debug.Print "B167-DONE"
End Sub
