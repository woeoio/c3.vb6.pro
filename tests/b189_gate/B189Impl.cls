VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B189Impl"
Option Explicit

' The implementor. Its OWN name is what VB6 would answer for an object of this class -- including
' through the B189Iface view -- so B189Impl itself is NOT in the implemented set and the fold fires
' for a variable declared As B189Impl (the positive half of the differential in B189Drv).
Implements B189Iface

Private Function B189Iface_Bar(ByVal n As Long) As Long
    B189Iface_Bar = n * 3
End Function
