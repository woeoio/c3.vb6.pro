' Ledger B180 (knife 69) regression fixture: StrPtr of a Variant used to hand out the
' address of a freshly minted BSTR and nobody reclaimed it.
'
'   before:  vb6_StrPtr(vb6_VariantToString(v))    VariantToString = vb6_CStr = vb6_Format
'            => one new BSTR per call, pointer handed to Win32, the string leaks
'   after :  vb6_StrPtrVar(v)                      vt already a string -> BORROW the slot's BSTR
'
' Measured on the shipped build (this repo's C3_k68_new.exe): 18000 calls of StrPtr(v) minus
' the same loop over a String variable = x64 1089536 / x86 909312 bytes. The window below
' prints the difference AND a band verdict, because page commits are quantised.
'
' The numeric/Empty branches still mint (VB6 coerces those to a string too) - that residue is
' the B164 shape-1 list, not this knife. Readings P1..P6 must match on x64 and x86.
Option Explicit

Private Const BV_APP As String = "C3ProbeB180"

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
    Dim s As String
    Dim e As String
    Dim vStr As Variant
    Dim vNum As Variant
    Dim vEmptyStr As Variant
    Dim p1 As LongPtr
    Dim p2 As LongPtr
    Dim w0 As Long
    Dim w1 As Long
    Dim b0 As Long
    Dim b1 As Long
    Dim leak As Long
    Dim apiBad As Long
    Dim okLeak As String
    Dim i As Long

    s = "abcdefghij"
    e = ""
    vStr = s
    vNum = 123456
    vEmptyStr = ""

    ' the three pointer answers must not move: borrow must behave like the old unwrap did
    p1 = StrPtr(vStr)
    If p1 <> 0 Then Debug.Print "P1-var-string-nonzero=Y" Else Debug.Print "P1-var-string-nonzero=N"
    p1 = StrPtr(e)
    Debug.Print "P2-empty-string=" & CLng(p1)
    p1 = StrPtr(vEmptyStr)
    Debug.Print "P3-variant-empty-string=" & CLng(p1)
    p1 = StrPtr(vNum)
    If p1 <> 0 Then Debug.Print "P4-var-numeric-nonzero=Y" Else Debug.Print "P4-var-numeric-nonzero=N"

    ' witness for "the slot's own BSTR is what we hand out": the same slot answered twice
    ' gives the same address (a per-call copy would not, and this knife is about the copy)
    p1 = StrPtr(vStr)
    p2 = StrPtr(vStr)
    If p1 = p2 Then Debug.Print "P5-stable=Y" Else Debug.Print "P5-stable=N"

    ' leak window: 20000 calls with a Variant minus the same loop with a String variable
    For i = 1 To 2000
        p1 = StrPtr(vStr)
    Next i
    w0 = ReadPF(pmc)
    For i = 1 To 20000
        p1 = StrPtr(vStr)
    Next i
    w1 = ReadPF(pmc)
    b0 = ReadPF(pmc)
    For i = 1 To 20000
        p1 = StrPtr(s)
    Next i
    b1 = ReadPF(pmc)

    If w0 < 0 Or w1 < 0 Or b0 < 0 Or b1 < 0 Then apiBad = 1
    leak = (w1 - w0) - (b1 - b0)
    If leak < 0 Then leak = 0 - leak
    If apiBad = 0 And leak < 400000 Then okLeak = "Y" Else okLeak = "N"
    Debug.Print "B180-apiBad=" & apiBad
    Debug.Print "B180-leak=" & leak
    Debug.Print "B180-leak-ok=" & okLeak
    Debug.Print "B180-acc=" & CLng(p1 <> 0)
    Debug.Print BV_APP & "-DONE"
End Sub

' Never called: pins the shapes the emit needles look at, so Test-EmitcShape can read them
' without the run above growing a needle per form. The Byte() arm is deliberately not
' exercised here: that rename (Fix 140) decides by TEXT and is not this knife's question.
Private Sub StrPtrWitness()
    Dim wHay As Variant
    Dim wTxt As String
    Dim wNum As Long
    wHay = "abc"
    wTxt = "def"
    wNum = 7
    Dim q As LongPtr
    q = StrPtr(wHay)
    q = StrPtr(wTxt)
    Debug.Print q
End Sub
