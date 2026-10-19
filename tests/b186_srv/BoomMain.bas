Attribute VB_Name = "BoomMain"
Option Explicit

' Ledger 305 B186 (knife 78): the VB-client half of the same reading. This exe does NOT reference
' the DLL -- CreateObject goes through the real registry chain, so the answer comes back as an
' HRESULT from a separate activation. Three arms, one per class:
'   BC1  Boom  -- the body raised 123 in Class_Initialize. The client must be ALIVE and must be
'                  able to read the SERVER's own number, not the blanket 429 that "some activation
'                  failed" used to mean (and not exit(123), which is what happened before B186).
'   BC2  Good  -- witness: a clean class still constructs and its member still answers 42.
'   BC3  Held  -- witness: a body owning its On Error still catches, so it constructs normally.
' On Error Resume Next, never an unguarded raise: this is a GUI-subsystem exe and an unhandled
' error pops a modal box, which turns a real reading into a run-time-out (see LateMain.bas).
Sub Main()
    On Error Resume Next

    Dim a As Object
    Err.Clear
    Set a = CreateObject("B186d.Boom")
    Debug.Print "BC1:num=" & Err.Number
    If a Is Nothing Then
        Debug.Print "BC1:null=yes"
    Else
        Debug.Print "BC1:null=no"
    End If

    Dim b As Object
    Err.Clear
    Set b = CreateObject("B186d.Good")
    If b Is Nothing Then
        Debug.Print "BC2:FAIL null"
    Else
        Debug.Print "BC2:twice=" & b.Twice(21) & " seed=" & b.seed & " num=" & Err.Number
    End If

    Dim c As Object
    Err.Clear
    Set c = CreateObject("B186d.Held")
    If c Is Nothing Then
        Debug.Print "BC3:FAIL null"
    Else
        Debug.Print "BC3:handled=" & c.handled & " twice=" & c.Twice(4)
    End If

    Debug.Print "BC-DONE"
End Sub
