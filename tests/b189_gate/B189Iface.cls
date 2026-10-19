VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B189Iface"
Option Explicit

' Ledger 305 B189 tier two -- the GATE witness. This class is Implements-ed by B189Impl, so a
' variable declared As B189Iface holds a VIEW: VB6 answers the CONCRETE class of whatever object
' is behind it, never the name in the declaration. The declaration fold must not fire here.
'
' This project is an EMIT fixture (Test-EmitcShape / Test-EmitcAbsent only) and is deliberately
' not built by cl: the shape it pins still goes through the un-fixed station where an interface
' view is handed to vb6_VariantFromValue as a struct (cl C2172) -- that cell is booked on its own.

Public Function Bar(ByVal n As Long) As Long
    Bar = n + 1
End Function
