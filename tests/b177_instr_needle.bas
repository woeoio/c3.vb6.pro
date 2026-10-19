Attribute VB_Name = "B177InstrNeedle"
' Ledger B177 (knife 68) regression fixture: InStr's needle slot has to answer TWO
' directions at once, and one C-kind string cannot carry both.
'
'   numeric actual -> BSTR   VB6 implicitly CStrs it (InStr(2, s, 7) searches for "7").
'   Variant actual -> RAW    vb6_InStr / vb6_InStrC are _Generic macros that dispatch on
'                            the NEEDED (vb6rtl_builtin.h:596); a vb6_VARIANT needle selects
'                            vb6_InStrVar, which unwraps AND frees (vb6rtl_string.c:235-247).
'
' Knife 67 answered the needle slot as "don't know" in the shapes where start was written,
' which also silenced the numeric fold: that emitted vb6_InStr(2, s, 7) -- a bare int into
' the macro's wchar_t* slot. cl does not diagnose it, and the program dies on that line
' (measured on the shipped build: x64 RC=139, no output at all).
' In the shapes where start was NOT written it did the opposite and pre-unwrapped Variant
' needles, which mints a temp nobody owns: 20000 calls of InStr(hay, vVariant) measured
' 753664 bytes against 0 for the same loop with a String needle.
'
' The readings below must match on x64 and x86. W_* are the leak windows; the fixture
' prints the difference AND a band verdict, because page commits are quantised.
Option Explicit

Private Const BV_APP As String = "C3ProbeB177"

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

Public Sub Main()
    Dim pmc As PMC
    Dim hay As String
    Dim hay2 As String
    Dim nLong As Long
    Dim vStr As Variant
    Dim vNum As Variant
    Dim vMiss As Variant
    Dim r As Long
    Dim apiBad As Long
    Dim wVar0 As Long
    Dim wVar1 As Long
    Dim wStr0 As Long
    Dim wStr1 As Long
    Dim leak As Long
    Dim okLeak As String
    Dim i As Long
    Dim sNeedle As String

    hay = "abc7def"
    hay2 = "aBcDeF"
    nLong = 7
    vStr = "7"
    vNum = 7
    vMiss = 9
    sNeedle = "7"

    ' numeric needles: the fold half. The last two are the shapes knife 67 lost.
    Debug.Print "N1-twoarg-lit=" & InStr(hay, 7)
    Debug.Print "N2-twoarg-long=" & InStr(hay, nLong)
    Debug.Print "N5-start-lit=" & InStr(2, hay, 7)
    Debug.Print "N6-start-long=" & InStr(2, hay, nLong)
    Debug.Print "N13-cmp-long=" & InStr(hay, nLong, vbTextCompare)
    Debug.Print "N14-start-cmp-long=" & InStr(2, hay, nLong, vbTextCompare)

    ' Variant needles: the raw half, and the values must not move either way.
    Debug.Print "N3-twoarg-varstr=" & InStr(hay, vStr)
    Debug.Print "N4-twoarg-varnum=" & InStr(hay, vNum)
    Debug.Print "N7-start-varstr=" & InStr(2, hay, vStr)
    Debug.Print "N8-cmp-varstr=" & InStr(hay, vStr, vbTextCompare)
    Debug.Print "N10-start-cmp-varstr=" & InStr(2, hay, vStr, vbTextCompare)

    ' witnesses that the answers are not just "anything non-zero"
    Debug.Print "N15-miss=" & InStr(hay, vMiss)
    Debug.Print "N16-bin-vs-text=" & InStr(hay2, "def") & "|" & InStr(hay2, "def", vbTextCompare)

    ' leak windows: 20000 calls with a Variant needle minus the same loop with a String
    ' needle. The shipped compiler measured 753664 bytes there (38 B per call); the fix
    ' must sit inside page noise, so the verdict is a band, not an absolute.
    For i = 1 To 2000
        r = InStr(hay, vStr)
    Next i
    wVar0 = ReadPF(pmc)
    For i = 1 To 20000
        r = InStr(hay, vStr)
    Next i
    wVar1 = ReadPF(pmc)
    wStr0 = ReadPF(pmc)
    For i = 1 To 20000
        r = InStr(hay, sNeedle)
    Next i
    wStr1 = ReadPF(pmc)

    If wVar0 < 0 Or wVar1 < 0 Or wStr0 < 0 Or wStr1 < 0 Then apiBad = 1
    leak = (wVar1 - wVar0) - (wStr1 - wStr0)
    If leak < 0 Then leak = 0 - leak
    If apiBad = 0 And leak < 400000 Then okLeak = "Y" Else okLeak = "N"
    Debug.Print "B177-apiBad=" & apiBad
    Debug.Print "B177-leak=" & leak
    Debug.Print "B177-leak-ok=" & okLeak
    Debug.Print "B177-acc=" & r
    Debug.Print BV_APP & "-DONE"
End Sub

' Never called. The four-arg Variant shapes and the haystack-side unwrap are pinned here
' so Test-EmitcShape can look at their emitted C without the run above growing a fifth
' reading per form.
Private Sub VariantWitness()
    Dim wHay As String
    Dim wVar As Variant
    Dim wNum As Long
    Dim wCmp As Variant
    wHay = "abc7def"
    wVar = "7"
    wNum = 2
    wCmp = 1
    Debug.Print InStr(wHay, wVar, wCmp)
    Debug.Print InStr(wNum, wHay, wVar, wCmp)
    Debug.Print InStr(wVar, wHay)
    Debug.Print InStr(wNum, wHay, 7)
End Sub
