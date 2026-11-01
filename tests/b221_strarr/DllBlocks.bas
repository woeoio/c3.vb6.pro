Attribute VB_Name = "DllBlocks"
Option Explicit

' Ledger 305 B220 (knife 222): the CoClass block puts StrArrCls on the IDispatch server surface,
' which is what lets tests	ools\disp_probe.c create it, fill the fields, and then Release --
' the Release is exactly the leg this account is about.

CoClass StrArrProbe
    [Implementation("StrArrCls")]
    [ComCreatable(True)]
End CoClass
