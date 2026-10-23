Attribute VB_Name = "K173Main"
Option Explicit

' <vbeclipse> fixture driver (账 #305 第一百七十三刀 / 新账 #468).
' K173Types.bas declares every Type; this module and K173Holder.cls only consume
' them, so all six receiver shapes below go through the cross-module path.
' Before the fix the consumer's placeholder Type had no member table, so no
' subscript on a member array was folded into C brackets and cl rejected every
' line with C2064 -- on this fixture, as in Ray.vbp, there was no exe at all.

Private mMod As K173Point

Private Function ReadLocal(q As K173Point) As Long
    ReadLocal = q.Coord(1) + q.Coord(4)
End Function

Private Sub FillMod()
    mMod.Coord(1) = 3
    mMod.Coord(3) = 4
End Sub

Public Sub Main()
    Dim h As New K173Holder
    h.PutPair 2, 4
    Debug.Print "K173-field="; h.PairSum()
    Debug.Print "K173-gridsep="; h.GridSep()

    Dim q As K173Point
    q.Coord(1) = 5
    q.Coord(4) = 6
    q.Grid(2, 1) = 7
    Debug.Print "K173-local="; ReadLocal(q)
    Debug.Print "K173-localgrid="; q.Grid(2, 1)

    FillMod
    Debug.Print "K173-modlevel="; mMod.Coord(1) + mMod.Coord(3)

    Dim o As K173Outer
    o.Inner.Row(2) = 8
    o.Inner.Row(3) = o.Inner.Row(2) + 1
    Debug.Print "K173-nested="; o.Inner.Row(3)

    Dim w As K173Point
    With w
        .Coord(2) = 9
        .Coord(2) = .Coord(2) + 1
        Debug.Print "K173-with="; .Coord(2)
    End With

    Dim b As K173Bag
    ReDim b.Items(1 To 3)
    b.Items(2) = 21
    Debug.Print "K173-dyn="; b.Items(2)

    Debug.Print "K173-DONE=1"
End Sub
