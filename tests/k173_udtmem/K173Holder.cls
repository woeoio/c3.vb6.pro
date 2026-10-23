VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "K173Holder"
Option Explicit

' The receiver shape Ray.vbp actually hits: a class-module field typed with a
' Type declared in ANOTHER module, subscripted one level down.

Private Loc As K173Point

Public Sub PutPair(ByVal a As Long, ByVal b As Long)
    Loc.Coord(1) = a
    Loc.Coord(2) = b
    Loc.Grid(1, 2) = a * 10
    Loc.Grid(2, 2) = b * 10
End Sub

Public Function PairSum() As Long
    PairSum = Loc.Coord(1) + Loc.Coord(2)
End Function

' Adjacent cells must stay adjacent-cell-distinct: writing (1,2) then (2,2)
' may not collapse onto one slot (that is the knife-135 / account #262 hazard,
' which cross-module Types used to re-open because the rank was unknown).
Public Function GridSep() As String
    GridSep = CStr(Loc.Grid(1, 2)) & "/" & CStr(Loc.Grid(2, 2))
End Function
