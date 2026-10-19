VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B185Holder"
Option Explicit

' Ledger 305 B185 (knife 77). The IDispatch server entry point (ComObj_Invoke) used to have no
' error boundary at all: it takes an EXCEPINFO argument and never touches it, its bridge callback
' returns void, and the function ends with an unconditional "return S_OK". So any runtime error
' raised inside a VB body and not caught there ran to vb6_RaiseError, found no frame there, printed
' one line to stderr and called exit(errorNumber) -- killing the CLIENT process. Measured before
' this knife, x64 and x86 alike: writing 40000 into a Public Integer field exited 6, and a plain
' "Err.Raise 5" in a Public Sub exited 5 (so this was never specific to the narrow-store gate of
' B184 station 2 -- that station just added one more raiser into the same hole).
'
' These members are read by tests\tools\disp_probe.c through a registry-free in-proc client:
'   PUTLEGAL/GETLEGAL  a legal value still has to land (the new frame must not swallow good stores)
'   PUTRANGE/GETRANGE  out of range: the client gets DISP_E_EXCEPTION + scode, and ci keeps 12345
'   PUTLONG/GETLONG    witness: a Long field is not narrow, 40000 must land unchecked
'   BOOM               witness that the fix is not tied to the narrow gate (Err.Raise 5 in a Sub)
'   HANDLED            witness for the ordering: a body owning its On Error still catches first,
'                      so the caller sees S_OK and an empty EXCEPINFO
' Readings must be identical on x64 and x86 (the frame is setjmp/longjmp across the bridge).

Public ci As Integer
Public cl As Long

Public Sub SeedLegal()
    ci = 12345
    cl = 999
End Sub

Public Sub BoomNoHandler()
    Err.Raise 5
End Sub

Public Sub HandledBoom()
    On Error GoTo H
    Err.Raise 5
    Debug.Print "B185-handled-noerr"
    Exit Sub
H:
    Debug.Print "B185-handled-raised=" & Err.Number
End Sub
