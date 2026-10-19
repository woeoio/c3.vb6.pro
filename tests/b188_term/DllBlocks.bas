Attribute VB_Name = "DllBlocks"
Option Explicit

' Ledger 305 B188: CoClass blocks put both classes on the IDispatch server surface. Without one
' the emitted product has an empty CoClassDesc table, so Release never reaches destroyFunc and
' this station is not exercised at all.

CoClass Boom
    [Implementation("T88Boom")]
    [ComCreatable(True)]
End CoClass

CoClass Good
    [Implementation("T88Good")]
    [ComCreatable(True)]
End CoClass
