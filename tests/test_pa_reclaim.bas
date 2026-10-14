Attribute VB_Name = "PaReclaim"
Option Explicit

' Acct #305 knife 56 (B166): one ParamArray round trip mints two things that nobody used to
' reclaim - the SAFEARRAY each CALL SITE builds for the callee (vb6_PA_Create), and the loop-side
' slot _pa_vN that "For Each x In <ParamArray>" overwrites once per element. The judgement asks
' the INCREMENT per mint, never an absolute number, and pins both arches with the same needles.
'
' W_LONG is exact-zero by construction: Long elements carry no weight, so a non-zero reading there
' can only mean the array itself was orphaned. The loop side is pinned by a SUBTRACTION -
' (8-element window - 8-element pack-only) minus (1-element window - 1-element pack-only) cancels
' both the packing cost and the per-call tail, leaving _pa_vN alone; what remains is reported
' against a band because page commits are quantised (the previous build measured 87/61 bytes per
' element there, this one sits at page noise).
'
' The string windows still grow ON PURPOSE: the call site mints each string literal twice
' (vb6_BSTR_FromStr plus the copy vb6_PA_SetBSTR makes inside), and that residual is the next
' cell with its own measured unit cost - see ledger B166. No needle here looks at it.

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

Function CountPA(ParamArray items() As Variant) As Long
    Dim e As Variant
    Dim n As Long
    For Each e In items
        n = n + 1
    Next e
    CountPA = n
End Function

Function PackOnly(ParamArray items() As Variant) As Long
    PackOnly = UBound(items) - LBound(items) + 1
End Function

Function SumLenPA(ParamArray items() As Variant) As Long
    Dim e As Variant
    Dim s As Long
    For Each e In items
        s = s + Len(e)
    Next e
    SumLenPA = s
End Function

Function LastPA(ParamArray items() As Variant) As String
    Dim e As Variant
    For Each e In items
    Next e
    LastPA = CStr(e)
End Function

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
    Dim s0 As Long, s1 As Long
    Dim p0 As Long, p1 As Long
    Dim o0 As Long, o1 As Long
    Dim q0 As Long, q1 As Long
    Dim l0 As Long, l1 As Long
    Dim i As Long
    Dim acc As Long
    Dim loopBytes As Long
    Dim okLoop As String
    Dim apiBad As Long

    acc = CountPA("AAAAAAAAAAAAAAAAAAAA", "BBBBBBBBBBBBBBBBBBBB", "CCCCCCCCCCCCCCCCCCCC", _
                  "DDDDDDDDDDDDDDDDDDDD", "EEEEEEEEEEEEEEEEEEEE", "FFFFFFFFFFFFFFFFFFFF", _
                  "GGGGGGGGGGGGGGGGGGGG", "HHHHHHHHHHHHHHHHHHHH")

    s0 = ReadPF(pmc)
    For i = 1 To 10000
        acc = acc + CountPA("AAAAAAAAAAAAAAAAAAAA", "BBBBBBBBBBBBBBBBBBBB", "CCCCCCCCCCCCCCCCCCCC", _
                            "DDDDDDDDDDDDDDDDDDDD", "EEEEEEEEEEEEEEEEEEEE", "FFFFFFFFFFFFFFFFFFFF", _
                            "GGGGGGGGGGGGGGGGGGGG", "HHHHHHHHHHHHHHHHHHHH")
    Next i
    s1 = ReadPF(pmc)

    p0 = ReadPF(pmc)
    For i = 1 To 10000
        acc = acc + PackOnly("AAAAAAAAAAAAAAAAAAAA", "BBBBBBBBBBBBBBBBBBBB", "CCCCCCCCCCCCCCCCCCCC", _
                             "DDDDDDDDDDDDDDDDDDDD", "EEEEEEEEEEEEEEEEEEEE", "FFFFFFFFFFFFFFFFFFFF", _
                             "GGGGGGGGGGGGGGGGGGGG", "HHHHHHHHHHHHHHHHHHHH")
    Next i
    p1 = ReadPF(pmc)

    o0 = ReadPF(pmc)
    For i = 1 To 10000
        acc = acc + CountPA("AAAAAAAAAAAAAAAAAAAA")
    Next i
    o1 = ReadPF(pmc)

    q0 = ReadPF(pmc)
    For i = 1 To 10000
        acc = acc + PackOnly("AAAAAAAAAAAAAAAAAAAA")
    Next i
    q1 = ReadPF(pmc)

    l0 = ReadPF(pmc)
    For i = 1 To 10000
        acc = acc + CountPA(1, 2, 3)
    Next i
    l1 = ReadPF(pmc)

    If s0 < 0 Or p0 < 0 Or o0 < 0 Or q0 < 0 Or l0 < 0 Then apiBad = 1
    If s1 < 0 Or p1 < 0 Or o1 < 0 Or q1 < 0 Or l1 < 0 Then apiBad = apiBad + 2

    ' 70,000 extra element reads carry the loop-side signal; band = 50 bytes per element.
    loopBytes = ((s1 - s0) - (p1 - p0)) - ((o1 - o0) - (q1 - q0))
    If loopBytes < 0 Then loopBytes = 0 - loopBytes
    If loopBytes < 3500000 Then okLoop = "Y" Else okLoop = "N"

    Debug.Print "B166-sum8=" & SumLenPA("AAAAAAAAAAAAAAAAAAAA", "BBBBBBBBBBBBBBBBBBBB")
    Debug.Print "B166-count=" & CountPA(1, 2, 3)
    Debug.Print "B166-last=" & LastPA("ZZZZ", "YYYY", "XXXX")
    Debug.Print "B166-wA8=" & (s1 - s0)
    Debug.Print "B166-wP8=" & (p1 - p0)
    Debug.Print "B166-wA1=" & (o1 - o0)
    Debug.Print "B166-wP1=" & (q1 - q0)
    Debug.Print "B166-long-window=" & (l1 - l0)
    Debug.Print "B166-loop-bytes=" & loopBytes
    Debug.Print "B166-loop-ok=" & okLoop
    Debug.Print "B166-acc=" & acc
    Debug.Print "B166-apiBad=" & apiBad
    Debug.Print "B166-DONE"
End Sub
