Attribute VB_Name = "FeSlotLeak"
Option Explicit

' Acct #305 knife 55 (B164 shape 4): the For Each loop variable is a vb6_VARIANT slot.
' The station used to emit only "e = <owned value>" (a bare struct store) while every
' ordinary assignment to a Variant slot emits "clear what the slot held, then store" -
' the same decision written two ways. Consequence: each element visited orphaned the
' BSTR the previous element had handed to the slot (measured ~115 bytes per element).
'
' The judgement asks the INCREMENT per store, never an absolute number, and pins both
' arches with the same needles. Both memory reads must really succeed, otherwise the
' delta would be a fake zero (self-subtraction) and the criterion worthless.
'
' Two timed windows are kept allocation-FREE on purpose (body = n = n + 1 + a VarType
' check), so what they measure is the station and nothing else. Everything that has to
' look at a value (Len / CStr) is untimed, because those shapes allocate a temp per call
' - that is shape 5 of the same account, deliberately NOT mixed in here.

Private Type PMC
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As LongPtr
    WorkingSetSize As LongPtr
    QuotaPeakPagedPoolUsage As LongPtr
    QuotaPagedPoolUsage As LongPtr
    QuotaPeakNonPagedPoolUsage As LongPtr
    QuotaNonPagedPoolUsage As LongPtr
    PagefileUsage As LongPtr
    PeakPagefileUsage As LongPtr
End Type

Declare Function K32GetProcessMemoryInfo Lib "kernel32" (ByVal hProcess As LongPtr, ByRef ppmc As PMC, ByVal cb As Long) As Long

Private gColl As Collection
Private gDec As Collection
Private gArr() As String

Function ReadPF(ByRef pmc As PMC) As Long
    Dim hr As Long
    pmc.cb = Len(pmc)
    hr = K32GetProcessMemoryInfo(-1, pmc, Len(pmc))
    If hr <> 1 Then
        ReadPF = -1
    Else
        ReadPF = CLng(pmc.PagefileUsage)
    End If
End Function

Sub Main()
    Dim pmc As PMC
    Dim pf0 As Long, pf1 As Long
    Dim pf2 As Long, pf3 As Long
    Dim i As Long, r As Long
    Dim e As Variant
    Dim e2 As Variant
    Dim t As Variant
    Dim sum As Long
    Dim bad As Long
    Dim keep As Variant
    Dim k As Long
    Dim s As String

    Set gColl = New Collection
    For i = 1 To 40
        gColl.Add String$(i, 65), "k" & i
    Next i

    Set gDec = New Collection
    gDec.Add CDec(5), "a"
    gDec.Add CDec(7), "b"

    ReDim gArr(1 To 40)
    For i = 1 To 40
        gArr(i) = String$(i, 66)
    Next i

    ' --- timed window A: For Each over a Collection (the station under test) ---
    pf0 = ReadPF(pmc)
    For r = 1 To 500
        For Each e In gColl
            If VarType(e) <> 8 Then bad = bad + 1
        Next e
    Next r
    pf1 = ReadPF(pmc)

    ' --- timed window B: For Each over a VB array (second station, same decision?) ---
    pf2 = ReadPF(pmc)
    For r = 1 To 500
        For Each e In gArr
            If VarType(e) <> 8 Then bad = bad + 1
        Next e
    Next r
    pf3 = ReadPF(pmc)

    Debug.Print "B164-stores=" & (500 * 40)
    Debug.Print "B164-api=" & IIf(pf0 >= 0, 1, 0) & "/" & IIf(pf1 >= 0, 1, 0) & "/" & IIf(pf2 >= 0, 1, 0) & "/" & IIf(pf3 >= 0, 1, 0)
    Debug.Print "B164-delta-coll=" & (pf1 - pf0)
    Debug.Print "B164-delta-arr=" & (pf3 - pf2)
    If pf0 >= 0 And pf1 >= 0 And pf2 >= 0 And pf3 >= 0 Then
        If (pf1 - pf0) < 8 * (500 * 40) Then
            Debug.Print "B164-coll-ok=Y"
        Else
            Debug.Print "B164-coll-ok=N"
        End If
        If (pf3 - pf2) < 8 * (500 * 40) Then
            Debug.Print "B164-arr-ok=Y"
        Else
            Debug.Print "B164-arr-ok=N"
        End If
    Else
        Debug.Print "B164-coll-ok=X"
        Debug.Print "B164-arr-ok=X"
    End If

    ' --- untimed value witnesses (these shapes allocate on purpose) ---
    sum = 0
    For Each e In gColl
        sum = sum + Len(e)
    Next e
    Debug.Print "B164-sum=" & sum
    Debug.Print "B164-last-len=" & Len(e)

    sum = 0
    For Each e In gArr
        sum = sum + Len(e)
    Next e
    Debug.Print "B164-arr-sum=" & sum

    ' Decimals through the loop variable (knife 54's ownership pair, x86 keeps a pointer)
    t = 0
    For Each e2 In gDec
        t = t + e2
    Next e2
    Debug.Print "B164-dec=" & CStr(t) & "/" & VarType(e2) & "/" & TypeName(e2)

    ' a value that leaves the loop must keep living: copied into a String variable
    ' (deliberately NOT `held = e` into another Variant - that shape aliases the slot's own
    ' buffer, which is a separate open question in this account, not something this needle
    ' should confuse the reading with)
    For Each e In gColl
        s = CStr(e)
    Next e
    Debug.Print "B164-held=" & Len(s) & "/" & Len(e)
    ' Station 4's other half: a slot-to-slot assignment (keep = e) used to emit the _Generic
    ' identity = a shallow copy, so two slots shared one buffer. Adding the Clear in the For
    ' Each station turned that quiet alias into a double free (measured: test_host_decimal on
    ' x86 died 0xC0000374 in 7 of 12 runs with only the Clear added; BASE 0 of 12 because it
    ' never released the old value at all). That shape now goes through the deep copy.
    For k = 1 To 300
        For Each e In gColl
            keep = e
        Next e
    Next k
    Debug.Print "B164-alias=" & Len(keep) & "/" & Len(e)
    Debug.Print "B164-bad=" & bad
    Debug.Print "B164-DONE"
End Sub
