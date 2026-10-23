Attribute VB_Name = "K483RedimMemberElem"
' Ledger #475 (B220) regression fixture: ReDim of a UDT *member* dynamic array.
'
' Before this knife the element flavour was only asked of a BARE variable name
' (arrayElemTypes_, keyed by `lowerVar`). A dotted target (`t.Numbers`) never hit that
' table, so the descriptor was allocated as one Variant per slot while every read/write
' went through VB6_SA_AT(int32_t, t.Numbers, i) -- 24 bytes/step allocated, 4 bytes/step
' used (numbers below are x64; x86 differs only in the Variant stride, which is why the
' bug stayed invisible until someone asked the descriptor itself).
'
' The witness is VarType(array): it reads the descriptor's own element flavour through
' vb6_SaVtOfArrayPtr -> vb6_SaVtFromElemType, so it reports what the carrier was BUILT
' as, not what the access macro assumes. Pre-fix it answered 8204 (vtArray|vtVariant)
' for every shape below; the values must now be the declared vtArray|vt<type>.
'
' Two heads feed this decision:
'   * resolveReDimElemType -- the dotted lower name, owner looked up in knownUdtVars_
'     (same-module UDT variable / module member);
'   * resolveReDimComplexElemType -- the subscripted member chain (arr(1).Numbers),
'     whose node.varName is only the BASE array name, so the bare-name fallback must
'     not run for it.
' Both ask the single owner resolveUdtMemberElemInfo (one udtMembers read).
Option Explicit

Private Type TScalars
    Longs() As Long
    Ints() As Integer
    Doubles() As Double
End Type

Private Type TStr
    Names() As String
End Type

Private Type TItem
    Id As Long
End Type

Private Type TOwner
    Items() As TItem
End Type

Public Sub Main()
    Dim s As TScalars
    Dim t As TStr
    Dim o As TOwner
    Dim arr(1 To 2) As TScalars

    ReDim s.Longs(1 To 3)
    s.Longs(1) = 11
    s.Longs(2) = 22
    s.Longs(3) = 33
    Debug.Print "K483-long-vt=" & VarType(s.Longs)
    Debug.Print "K483-long=" & s.Longs(1) & "," & s.Longs(2) & "," & s.Longs(3)
    ReDim Preserve s.Longs(1 To 4)
    s.Longs(4) = 44
    Debug.Print "K483-long-preserve=" & s.Longs(1) & "," & s.Longs(2) & "," & s.Longs(3) & "," & s.Longs(4)

    ReDim s.Ints(1 To 2)
    s.Ints(1) = 5
    s.Ints(2) = 6
    Debug.Print "K483-int-vt=" & VarType(s.Ints) & "/" & s.Ints(1) & "," & s.Ints(2)

    ReDim s.Doubles(1 To 2)
    s.Doubles(1) = 2#
    s.Doubles(2) = 3#
    Debug.Print "K483-dbl-vt=" & VarType(s.Doubles) & "/" & s.Doubles(1) & "," & s.Doubles(2)

    ReDim t.Names(1 To 2)
    t.Names(1) = "A"
    t.Names(2) = "B"
    Debug.Print "K483-str-vt=" & VarType(t.Names) & "/" & t.Names(1) & "," & t.Names(2)

    ReDim o.Items(1 To 2)
    o.Items(1).Id = 7
    o.Items(2).Id = 8
    Debug.Print "K483-udt-vt=" & VarType(o.Items) & "/" & o.Items(1).Id & "," & o.Items(2).Id

    ReDim arr(1).Longs(1 To 2)
    arr(1).Longs(1) = 9
    arr(1).Longs(2) = 10
    Debug.Print "K483-complex-vt=" & VarType(arr(1).Longs) & "/" & arr(1).Longs(1) & "," & arr(1).Longs(2)

    Erase s.Longs
    Erase s.Ints
    Erase s.Doubles
    Erase t.Names
    Erase o.Items
    Erase arr(1).Longs
    Debug.Print "K483-DONE"
End Sub
