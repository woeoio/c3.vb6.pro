Attribute VB_Name = "DllBlocks"
Option Explicit

' Ledger 305 B185 (knife 77): the CoClass block is what puts B185Holder on the IDispatch server
' surface. Without it the class has no entry-table row at all, and an out-of-proc (or late-bound)
' write never reaches the field bridge -- see the four-line VB6 class header note in B185Holder.cls
' for the other way this table silently comes out empty.

CoClass Hold
    [Implementation("B185Holder")]
    [ComCreatable(True)]
End CoClass
