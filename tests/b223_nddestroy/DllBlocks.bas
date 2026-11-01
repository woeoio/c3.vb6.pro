Attribute VB_Name = "DllBlocks"
Option Explicit

' Ledger 305 B223 (knife 223 measure): three coclasses on the IDispatch surface so tests	tools	disp_probe.c
' can create each one separately and release it (the release is the leg under measurement).

CoClass NdLongProbe
    [Implementation("NdLongCls")]
    [ComCreatable(True)]
End CoClass

CoClass NdStrProbe
    [Implementation("NdStrCls")]
    [ComCreatable(True)]
End CoClass

CoClass OneDProbe
    [Implementation("OneDCls")]
    [ComCreatable(True)]
End CoClass
