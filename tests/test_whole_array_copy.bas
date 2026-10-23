Attribute VB_Name = "WholeArrayCopy"
Option Explicit

' Acct #305 knife 158 / new bug #460: VB6's WHOLE-ARRAY assignment copies. Today's bare form
' (`d = b`, no parentheses anywhere - VB6's own spelling, 125 sites in the real corpus) emitted a
' carrier-pointer handover instead: `d = b;` (two names, one carrier) and, for a fixed target,
' `pv = vb6_ChkLong(b)` - the pointer squeezed through int32, which is an instant AV on x64 and a
' silent alias on x86. The whole-array question now has one owner (isWholeArrayRef) and it answers
' for the bare name too, so both shapes go through vb6_ArrayAssign1D.
'
' Windows, both architectures, byte-for-byte same expectation:
'   B259 dyn <- dyn, then the source is rewritten  -> the copy must keep 123456 / UBound 5
'   B260 module FIXED array <- dynamic             -> the old x64 error-6 window
'   B261 local fixed array <- dynamic              -> same shape, local storage
'   B262 dynamic <- Variant that holds an array    -> the third arm (VariantToSafeArray1D)
'   B263 a Variant that BORROWED the carrier, then the array is reassigned
'        -> VB6 leaves the Variant holding its own copy, so B263 must keep reading 5. Knife 159 let
'           the bare form hand the REAL destination to the copy (so the old carrier is freed) only
'           where the escape roster says this procedure owns it; here `mv2 = loc` boxed that carrier,
'           so dst stays NULL - this arm is the witness that the gate refuses the borrowed one.
'   B264 DONE
'
' The other half of this knife is in the escape roster (cgen_iface_vtbl.cpp EscapeScan): the LHS of a
' whole-array assignment no longer counts as an escape. Knife 156 put it there because `l = mg` used
' to alias - freeing l at exit would have torn down the module's array. Now it copies, so l owns the
' clone and the exit reclaims it; without that exemption every whole-array assignment leaks one
' carrier per call, which is exactly what test_local_array_carrier's B225 measured (0 -> 1 -> 0).
'
' Mutation control (this machine, both arches, pre-fix binary at .build/wt_base156):
'   x86  -> B259/B260/B261 all read the REWRITTEN source value (999) = the alias, rest unchanged
'   x64  -> B259 reads 999, then the run dies at B260 with "Unhandled VB6 Error #6: Overflow"
'           (the carrier pointer goes through vb6_ChkLong, which is a range check, so the width bug
'           surfaces as error 6 rather than an AV - and the window never prints DONE)
'   B262/B263 read the same pre and post: the boxing path already handed back a fresh-enough thing
'   and the borrow witness is same-direction only. Those two arms are guarded by the emit needles
'   (b158_whole_array_assign_goes_through_the_copy / b158_no_bare_carrier_handover_left), not here.
'   Pre-fix transcripts (BASE = the knife-156 compiler built in .build/wt_base156) archived at
'   .build/k2050_pre_x64.txt and .build/k2052_pre_x86.txt; post-fix at k2050_new_x64 / k2052_new_x86.

Private mFix(0 To 8) As Long
Private mv2 As Variant

Sub CopyNotAlias()
    Dim b() As Long
    Dim d() As Long
    ReDim b(0 To 5)
    b(3) = 123456
    ReDim d(0 To 1)
    d = b
    b(3) = 999
    d(0) = 7
    Debug.Print "B259-copy=" & CStr(d(3)) & " ubound=" & CStr(UBound(d)) & " src=" & CStr(b(3))
End Sub

Sub FixTarget()
    Dim b() As Long
    ReDim b(0 To 6)
    b(3) = 234567
    mFix = b
    b(3) = 999
    Debug.Print "B260-fix=" & CStr(mFix(3)) & " ubound=" & CStr(UBound(mFix))
End Sub

Sub LocFixTarget()
    Dim b() As Long
    Dim p(0 To 4) As Long
    ReDim b(0 To 3)
    b(2) = 345678
    p = b
    b(2) = 999
    Debug.Print "B261-locfix=" & CStr(p(2))
End Sub

Sub FromVariant()
    Dim v As Variant
    Dim e() As Long
    v = Array(11, 22, 33)
    ReDim e(0 To 0)
    e = v
    e(1) = 44
    Debug.Print "B262-var=" & CStr(v(1)) & " copy=" & CStr(e(1))
End Sub

Sub BoxThenReassign()
    Dim loc() As Long
    Dim src() As Long
    ReDim loc(0 To 3)
    loc(1) = 5
    mv2 = loc
    ReDim src(0 To 2)
    src(1) = 6
    loc = src
    Debug.Print "B263-hold=" & CStr(mv2(1)) & " own=" & CStr(loc(1))
End Sub

Sub Main()
    CopyNotAlias
    FixTarget
    LocFixTarget
    FromVariant
    BoxThenReassign
    Debug.Print "B264-DONE=1"
End Sub
