VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B186Boom"
Option Explicit

' Ledger 305 B186 (knife 78). The class factory is the SECOND station that walks in from a COM
' entry point straight into VB code: IClassFactory::CreateInstance -> vb6_ComObject_Create ->
' desc->factoryFunc() -> Class_Initialize. Before this knife that step had no error boundary
' either, so an uncaught runtime error in a body ran to vb6_RaiseError, found no frame, printed
' one line to stderr and called exit(errorNumber) -- killing the CLIENT process (measured on
' x64 and x86 alike: a body of just "Err.Raise 123" exited 123 and the probe's CREATE line never
' printed). This channel has no EXCEPINFO at all, so the number has to travel inside the
' HRESULT the factory answers with.
'
' 123 is deliberately not 7: 7 is VB's own "Out of memory", and the reading that matters here is
' that a body's error is NOT reported as the allocation failure the old code always claimed.

Private Sub Class_Initialize()
    Err.Raise 123
End Sub

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
