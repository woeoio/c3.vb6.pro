VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Driver"
Option Explicit

' Ledger 305 B187 second leg (boxing): inside an ActiveX (Type=DLL) project, does an
' in-project object handed to an Object slot answer a late-bound member call? The CREATE=
' arm of disp_probe calls Twice(21) on this class, so one number carries both readings:
'   early * 1000 + late  =>  42042 = both legs answer,  42000 = the slot holds a raw C
'   pointer, so the late-bound half silently answers nothing (the pre-fix reading).

Public Function Twice(ByVal n As Long) As Long
    Dim o As New B187Good
    Dim v As Object
    Dim early As Long
    Dim latev As Long
    early = o.Twice(n)
    Set v = o
    On Error Resume Next
    latev = v.Twice(n)
    Twice = early * 1000 + latev
End Function
