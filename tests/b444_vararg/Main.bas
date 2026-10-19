Attribute VB_Name = "Main"
Option Explicit

' Acct #305 knife 121 (B217): the ByRef-Variant actual-argument station mints a named temp
' Variant that holds TWO references per call - one taken by vb6_ComObject_From*Instance and a
' second one taken by the boxing helper itself (vb6_VariantObject AddRefs). The station now
' hands the second one back at the end of the SAME statement. Two things must stay true:
'   the callee still SEES the object (VarType = 9) => the reclaim runs after the call, not before;
'   the caller's own instance is still alive after the call (n / Twice) => the reclaim does not
'   touch the reference the caller owns, and it does not touch it even when the callee clears
'   the slot itself.
' P2 is the shape that would break if the reclaim were attached to the emitter's generic
' post-line queue: two boxed arguments in ONE statement, and both must still be seen.

Private mK As B444Impl

Sub TakeSee(ByRef Item As Variant)
    Debug.Print "P1-see vt=" & VarType(Item) & " tn=" & TypeName(Item)
End Sub

Sub TakeTwo(ByRef A As Variant, ByRef B As Variant)
    Debug.Print "P2-two a-vt=" & VarType(A) & " b-vt=" & VarType(B)
End Sub

Sub TakeClear(ByRef Item As Variant)
    Set Item = Nothing
End Sub

Sub TakeWrite(ByRef Item As Variant)
    Item = 5
End Sub

Function Sees(ByRef Item As Variant) As Long
    Sees = VarType(Item)
End Function

Function Fresh() As Long
    Dim r As B444Impl
    Set r = New B444Impl
    r.N = 7
    TakeSee r
    Fresh = r.N
End Function

Sub Main()
    Dim a As B444Impl
    Dim b As B444Impl

    Set a = New B444Impl
    a.N = 7
    Set b = New B444Impl
    b.N = 9
    Set mK = New B444Impl
    mK.N = 11

    TakeSee a
    TakeTwo a, b
    TakeSee mK
    Debug.Print "P4-after-see n=" & a.N & " twice=" & a.Twice(21)

    TakeClear a
    Debug.Print "P5-after-clear n=" & a.N & " tn=" & TypeName(a) & " twice=" & a.Twice(21)

    TakeWrite b
    Debug.Print "P6-after-write n=" & b.N & " tn=" & TypeName(b)

    If Sees(a) = 9 Then
        Debug.Print "P7-in-condition seen=9"
    Else
        Debug.Print "P7-in-condition seen=" & Sees(a)
    End If

    Debug.Print "P8-in-print vt=" & Sees(mK)
    Debug.Print "P9-fresh=" & Fresh()
    Debug.Print "B444-DONE"
End Sub
