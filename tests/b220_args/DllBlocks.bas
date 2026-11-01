Attribute VB_Name = "DllBlocks"
Option Explicit

' Ledger 305 B219 (knife 220): the CoClass block is what puts B220Args on the IDispatch server
' surface (same shape as tests\b185_srv\DllBlocks.bas). Without it there is no entry-table row,
' and the inbound argument table this account is about never gets built at all.

CoClass ArgProbe
    [Implementation("B220Args")]
    [ComCreatable(True)]
End CoClass
