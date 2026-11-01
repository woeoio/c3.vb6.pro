Attribute VB_Name = "DllBlocks"
Option Explicit

' Ledger 305 SB222 (knife 230): puts ArrOutCls on the IDispatch surface so tests/tools/disp_probe.c
' can hand it a real SAFEARRAY (ARRARG=Fill:4 / ARRARGV=Fill:4).

CoClass ArrOutProbe
    [Implementation("ArrOutCls")]
    [ComCreatable(True)]
End CoClass
