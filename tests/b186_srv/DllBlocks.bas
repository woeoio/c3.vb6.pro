Attribute VB_Name = "DllBlocks"
Option Explicit

' Ledger 305 B186 (knife 78): three CoClass blocks are what put the three classes on the IDispatch
' server surface at all (an in-project class without one gets an empty g_vb6_coclasses[] table --
' the precondition measured in B185).

CoClass Boom
    [Implementation("B186Boom")]
    [ComCreatable(True)]
End CoClass

CoClass Good
    [Implementation("B186Good")]
    [ComCreatable(True)]
End CoClass

CoClass Held
    [Implementation("B186Held")]
    [ComCreatable(True)]
End CoClass
