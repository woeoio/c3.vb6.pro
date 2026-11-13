Attribute VB_Name = "DllBlocks"
Option Explicit

' Account #610 knife #344: puts FgzVarArrCls on the IDispatch surface so
' tests/tools/disp_probe.c can hand it a real SAFEARRAY(VARIANT) (ARRVAR=TakeVarArr).

CoClass FgzVarArrProbe
    [Implementation("FgzVarArrCls")]
    [ComCreatable(True)]
End CoClass
