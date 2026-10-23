Attribute VB_Name = "K173Types"
Option Explicit

' <vbeclipse> fixture for ledger #305 knife 173 (new account #468).
' Every Type here is declared in THIS module and consumed from K173Main.bas and
' K173Holder.cls, i.e. cross-module.  Before the fix the consumer saw a
' member-table-less placeholder for these Types, so a subscript on a member
' array was never folded into C brackets:  me->Loc.Coord(1) = v  (cl C2064).

Public Type K173Point
    Coord(1 To 4) As Long          ' fixed 1-D member array
    Grid(1 To 2, 1 To 3) As Long   ' fixed 2-D member array (knife 135 / #262 shape)
    Tag As Long
End Type

Public Type K173Bag
    Items() As Long                ' dynamic member array (VB6_SA_AT family)
End Type

Public Type K173Inner
    Row(1 To 3) As Long
End Type

Public Type K173Outer
    Inner As K173Inner             ' nested UDT member, subscript one level deeper
    Name As String * 4
End Type
