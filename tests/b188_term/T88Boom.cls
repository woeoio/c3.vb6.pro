VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "T88Boom"
Option Explicit

' Ledger 305 B188, station three of "walking in from a COM entry point": the way OUT.
' IUnknown::Release -> ComObj_Release -> desc->destroyFunc -> Class_Terminate below. That station
' has no error channel of any kind (Release hands back a refcount, not an HRESULT, and there is no
' EXCEPINFO), so before the fix an uncaught error in the body reached vb6_RaiseError with no frame:
' one stderr line + ExitProcess(number). The client died while simply letting go of the object --
' here with 234, chosen so it cannot be mistaken for the usual 7/9 readings.

Private Sub Class_Terminate()
    Err.Raise 234
End Sub

Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function
