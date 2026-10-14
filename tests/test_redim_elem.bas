' Acct #305 knife 42 = section B154 (2026-10-12). The element type of a ReDim'd
' dynamic array is asked of the DECLARATION, not of the ReDim statement's own As
' clause. ASCII comments only (C3 reads sources as ANSI/GBK).
'
' What was wrong before this knife (readings .build/b1371_pre/r6.txt = BASE x64,
' rn6.txt / rn3.txt = NEW x64 / x86, the two arches line-for-line identical):
'   creation  vb6_SafeArrayReDim1D(vb6_sa_variant, ...)      <- As clause only
'   creation  vb6_SafeArrayReDimND(vb6_sa_variant, 2, ...)   <- As clause only
'   access    VB6_SA_AT(BSTR, s1, 1) / VB6_SA_ND_AT2(BSTR, ..., 2, 3)
'                                                          <- asks arrayElemTypes_
' One descriptor, two answers: allocation walked 24 bytes per element (x64) while
' the write walked 8, so the direct read agreed with the direct write (looked fine)
' while every consumer that trusts the descriptor itself walked the wrong stride.
' The witness that made it visible at run time is R3-box: v = a2 then v(2, 3) read
' back an EMPTY string on both arches (the box path landed in a different slot).
'
' Only UDT arrays had a declaration fallback before (arrayUdtElemTypes_, Fix 055);
' scalar / String / Object arrays did not. Fixed arrays were always right, and that
' is the answer this knife hands to the dynamic path too:
'   Dim f(5) As String   ->  vb6_SafeArrayCreate1D(vb6_sa_bstr, 0, 5)
'
' R1-label / R3-label / R5-label pinned VarType + TypeName as knife 42 MEASURED them
' (8204 / "Variant" for every flavor) and said out loud that VB6 answers
' 8200 / "String()" there. Knife 46 (section B155 cells 1 and 2) fixed the tag, so
' those three now read the declared flavor - which is what R17 below pins for EVERY
' arm of the new table. See the R17 block for the two arms still not VB6's answer.
Option Explicit

Type TRElem
    i As Long
End Type

' Module-level so Main and GrowMp share one descriptor (R14 spans two procedures).
Private mp() As String

Public Sub Main()
    Dim s1() As String
    Dim a2() As String
    Dim b3() As Byte
    Dim l4() As Long
    Dim u5() As TRElem
    Dim w6() As Variant
    Dim x7 As Variant
    Dim s8() As String
    Dim v As Variant
    Dim p0() As String
    Dim g1() As String
    Dim w1() As Variant
    Dim b1() As Byte
    Dim n1() As String
    Dim q0b() As Boolean
    Dim q0y() As Byte
    Dim q0i() As Integer
    Dim q0l() As Long
    Dim q0s() As Single
    Dim q0d() As Double
    Dim q0c() As Currency
    Dim q0t() As Date
    Dim q0o() As Object
    Dim q0u() As TRElem
    Dim q0n() As String

    ' R1: 1-D String, bare ReDim. cell = direct round trip (worked before too -
    ' direct write and direct read use the same stride); label = the open B155 cell.
    ReDim s1(3)
    s1(1) = "A"
    Debug.Print "R1-cell="; s1(1)
    Debug.Print "R1-label-vt="; VarType(s1); " tn="; TypeName(s1)

    ' R2: ReDim Preserve keeps what was there and grows the bound (VB6 semantics).
    ReDim Preserve s1(6)
    Debug.Print "R2-preserve-cell="; s1(1); " ub="; UBound(s1)

    ' R3: ND String. box = the needle this knife fixes (empty before, "X" after).
    ReDim a2(1 To 3, 1 To 4)
    a2(2, 3) = "X"
    Debug.Print "R3-cell="; CStr(a2(2, 3))
    Debug.Print "R3-ub2="; UBound(a2, 2); " lb2="; LBound(a2, 2)
    v = a2
    Debug.Print "R3-box-cell="; CStr(v(2, 3))
    Debug.Print "R3-label-vt="; VarType(a2); " tn="; TypeName(a2)

    ' R4: ND Byte - the single most-used declared element type in the real VB6
    ' corpus (596 bare-ReDim lines), so the stride must be 1 byte, not 24.
    ReDim b3(1 To 2, 1 To 3)
    b3(2, 3) = 200
    Debug.Print "R4-cell="; b3(2, 3)
    v = b3
    Debug.Print "R4-box-cell="; CStr(v(2, 3))

    ' R5: ND Long - regression witness (this one already read back its value
    ' before the knife because write and read agreed with each other).
    ReDim l4(2, 2)
    l4(1, 2) = 77
    Debug.Print "R5-cell="; l4(1, 2)
    Debug.Print "R5-ub1="; UBound(l4, 1); " ub2="; UBound(l4, 2)
    Debug.Print "R5-label-vt="; VarType(l4); " tn="; TypeName(l4)

    ' R6: ND UDT - the family that already had a declaration fallback; must not move.
    ReDim u5(1, 2)
    Debug.Print "R6-ub2="; UBound(u5, 2)

    ' R7: declared As Variant array - Variant IS the declared answer, so the
    ' element stays a slot (this is the shape the old fallback got right by luck).
    ReDim w6(3)
    w6(2) = "V"
    Debug.Print "R7-cell="; CStr(w6(2)); " vt="; VarType(w6)

    ' R8: Variant slot (Dim x As Variant, then ReDim) - nothing declares an element
    ' type, so the descriptor keeps the variant slot and the box path reads it back.
    ReDim x7(1, 2)
    x7(1, 1) = "S"
    Debug.Print "R8-cell="; CStr(x7(1, 1))

    ' R9: ReDim WITH an As clause - the clause still wins (VB6 semantics, and the
    ' corpus measures declared == As on 1580 of 1588 such lines).
    ReDim s8(2) As String
    s8(1) = "Z"
    Debug.Print "R9-cell="; s8(1)

    ' ---- R10..R14: knife 43 = section B155 cell 4 --------------------------
    ' The FIRST allocation of a multi-dimensional array done as a ReDim Preserve.
    ' vb6_SafeArrayReDimPreserveND had no elemType parameter at all, so a NULL
    ' descriptor could only be created as vb6_sa_empty - which the size table answers
    ' with its default 4 bytes - while the access side walks the DECLARED stride
    ' (x64: String 8, Variant 24). The write runs off the end of a block a sixth of
    ' the size the accesses assume. This one is LOUD, not silent: the pre-knife
    ' binary dies with 0xC0000374 (heap corruption) having printed NOTHING (reading
    ' .build/b1382_preserve/rzb6.txt = 0 bytes).
    '
    ' R10..R14 keep one ND ReDim statement per array per procedure on purpose: that
    ' used to be a HARD limit, not a style choice - the generated bounds temp was named
    ' after the ARRAY, so a second ND ReDim of the same array in the same scope declared
    ' it twice inside one C function = cl C2374 (redefinition; multiple initialization),
    ' measured on both binaries (.build/b1395_k43/{n6,b6}/c3-error.log). Knife 44
    ' (section B156) fixed that by naming the temp per STATEMENT through one helper;
    ' R15 below is the shape that would not compile before it, while R14 stays as the
    ' cross-procedure witness (module-level array, two procedures).
    ReDim Preserve p0(1 To 2, 1 To 3)
    p0(1, 1) = "one"
    p0(2, 3) = "six"
    Debug.Print "R10-c11="; CStr(p0(1, 1)); " c23="; CStr(p0(2, 3))
    Debug.Print "R10-ub2="; UBound(p0, 2); " lb2="; LBound(p0, 2)

    ' R11: Variant elements - 24 bytes walked inside a 4-byte block (x64).
    ReDim Preserve w1(1 To 2, 1 To 2)
    w1(1, 1) = "v11"
    w1(2, 2) = "v22"
    Debug.Print "R11-c11="; CStr(w1(1, 1)); " c22="; CStr(w1(2, 2))

    ' R12: Byte - the other half of the same mistake: 1 byte walked inside a 4-byte
    ' block, so it never overflowed and only the descriptor's own consumers (Destroy,
    ' the boxed element read) walked wrong.
    ReDim Preserve b1(1 To 2, 1 To 3)
    b1(1, 2) = 7
    b1(2, 3) = 200
    Debug.Print "R12-c12="; b1(1, 2); " c23="; b1(2, 3)

    ' R13: the box path over a Preserve-created descriptor - knife 42's second cell
    ' and knife 43's first one have to agree, or the value lands in a slot nobody
    ' reads back.
    v = p0
    Debug.Print "R13-box-c23="; CStr(v(2, 3))

    ' R14: Preserve across TWO procedures. Main allocates and writes mp(1, 1); GrowMp grows
    ' the LAST dimension (the direction VB6 allows under Preserve) and writes mp(1, 3). The
    ' split used to be forced, not chosen - see the C2374 note above; knife 44 lifted it and
    ' R15 is the same-procedure shape. What R14 does NOT witness: both of its reads land
    ' where the RTL's two index conventions agree (cell (1, 1) is the one slot they share,
    ' and mp(1, 3) was written AFTER the growth), so it printed exactly this on the
    ' pre-knife-45 compiler too (.build/b1405_k45build/rng6.txt). R16 is the needle that
    ' tests "preserve" instead of merely reading the slot both conventions agree on.
    ReDim Preserve mp(1 To 2, 1 To 3)
    mp(1, 1) = "keep"
    GrowMp
    Debug.Print "R14-kept="; CStr(mp(1, 1)); " grown="; CStr(mp(1, 3))
    Debug.Print "R14-ub2="; UBound(mp, 2)

    ' R15 (knife 44 = section B156): TWO multi-dim ReDim statements for the SAME array
    ' in ONE procedure - plain first, then a Preserve that grows the last dimension.
    ' Before that knife this file would not even compile (cl C2374), which is why this
    ' is the discriminating shape: exit code + the two cells are the witness, and
    ' R15-kept proves the growth kept what the FIRST statement wrote.
    ReDim g1(1 To 2, 1 To 3)
    g1(1, 1) = "g11"
    ReDim Preserve g1(1 To 2, 1 To 5)
    Debug.Print "R15-kept="; CStr(g1(1, 1)); " ub2="; UBound(g1, 2)
    g1(2, 4) = "g24"
    Debug.Print "R15-cell="; CStr(g1(2, 4))

    ' R16 (knife 45 = section B158): a Preserve growth must not MOVE anything. The copy
    ' loop used to walk the LAST dimension fastest while the access side (vb6_SaNdElemPtr,
    ' and vb6_SafeArrayND_Offset next to it) walks the FIRST one fastest, so growing
    ' (1 To 2, 1 To 3) into (1 To 2, 1 To 6) teleported the cells that the two arithmetic
    ' disagree on. Measured on a compiler that has knives 42-44 but not 45
    ' (.build/b1405_k45build/rng6.txt): six cells written BEFORE the growth, three of them
    ' still there afterwards, and the three missing ones reappearing in the fresh columns -
    ' this line read "a11/a12/", that one read "b21//", and b22 turned up as n1(1, 4). The
    ' standalone probe pr4.bas tells the same story cell by cell
    ' (.build/b1403_k44probe/rpn4_pre45.txt), and after the knife it prints byte-for-byte
    ' what the BASE compiler prints (.build/b1405_k45build/{rk5b6,rk5b3}.txt, x64 and x86
    ' identical). R14/R15 cannot see any of this: both read the (1, 1) cell, the one slot
    ' the two conventions agree on.
    ReDim n1(1 To 2, 1 To 3)
    n1(1, 1) = "a11"
    n1(1, 2) = "a12"
    n1(1, 3) = "a13"
    n1(2, 1) = "b21"
    n1(2, 2) = "b22"
    n1(2, 3) = "b23"
    ReDim Preserve n1(1 To 2, 1 To 6)
    Debug.Print "R16-keep=" & CStr(n1(1, 1)) & "/" & CStr(n1(1, 2)) & "/" & CStr(n1(1, 3))
    Debug.Print "R16-keep2=" & CStr(n1(2, 1)) & "/" & CStr(n1(2, 2)) & "/" & CStr(n1(2, 3))
    Debug.Print "R16-new=" & CStr(n1(1, 4)) & "|" & CStr(n1(2, 6)) & " ub2=" & CStr(UBound(n1, 2))

    ' ---- R17: knife 46 = section B155 cells 1 and 2 --------------------------
    ' "Which flavor is this array" now has ONE table (vb6_SaVtFromElemType in
    ' vb6rtl_array.c), sitting next to the element-size table it must move with.
    ' Before this knife the box tag was a hardcoded vb6_vtArray|vb6_vtVariant at all
    ' four consumer points, so every typed array answered VarType=8204 and TypeName
    ' fell through vb6_TypeName's scalar-only switch to "Variant". VB6 answers
    ' vbArray + <element vt> and "<element name>()", so each arm of the new table is
    ' pinned here - a table row nobody prints is a row nobody keeps honest.
    ' The two arms that are NOT VB6's answer yet are printed on purpose too and are
    ' booked as section B159: Date has no arm in this enum (a Date array stores a
    ' double, so it reads 8197/"Double()"), and a UDT array falls back to the
    ' pre-knife answer because VB6 has no copyable convention for boxing one.
    ReDim q0b(1)
    Debug.Print "R17-bool-vt="; VarType(q0b); " tn="; TypeName(q0b)
    ReDim q0y(1)
    Debug.Print "R17-byte-vt="; VarType(q0y); " tn="; TypeName(q0y)
    ReDim q0i(1)
    Debug.Print "R17-int-vt="; VarType(q0i); " tn="; TypeName(q0i)
    ReDim q0l(1)
    q0l(0) = 4242
    Debug.Print "R17-long-vt="; VarType(q0l); " tn="; TypeName(q0l)
    ReDim q0s(1)
    Debug.Print "R17-sing-vt="; VarType(q0s); " tn="; TypeName(q0s)
    ReDim q0d(1)
    Debug.Print "R17-dbl-vt="; VarType(q0d); " tn="; TypeName(q0d)
    ReDim q0c(1)
    Debug.Print "R17-cur-vt="; VarType(q0c); " tn="; TypeName(q0c)
    ReDim q0t(1)
    Debug.Print "R17-date-vt="; VarType(q0t); " tn="; TypeName(q0t)
    ReDim q0o(1)
    Debug.Print "R17-obj-vt="; VarType(q0o); " tn="; TypeName(q0o)
    ReDim q0u(1)
    Debug.Print "R17-udt-vt="; VarType(q0u); " tn="; TypeName(q0u)

    ' R17-boxed: the same descriptor through the OTHER consumer (vb6_VariantArray,
    ' the box a Variant slot takes). Label and descriptor must agree here, so this
    ' line asks three things of one array: VarType of the slot, TypeName of the slot,
    ' and the element read back out of the slot (that last one is knife 42's
    ' descriptor-trusting consumer, so the two answers are pinned together).
    v = q0l
    Debug.Print "R17-boxed-vt="; VarType(v); " tn="; TypeName(v); " cell="; CStr(v(0))

    ' R17-boxw: the same slot, now WRITTEN through. vb6_VariantArraySet / SetND start
    ' with `if (elemType != vb6_sa_variant) return;`, so a write into a typed array's
    ' box is a SILENT NO-OP - section B155 cell 3, still open. Measured on both arches
    ' (.build/k46run2_both.txt): the cell still reads 4242, the value the box was made
    ' from, and 99 never lands. Whoever fixes cell 3 turns this needle red, and that is
    ' the success signal, not a regression.
    v(0) = 99
    Debug.Print "R17-boxw="; CStr(v(0))

    ' R17-undim: never dimensioned (descriptor is NULL) - vb6_SaVtOfArrayPtr refuses
    ' to guess and hands back the pre-knife answer. VB6 raises "Array not even
    ' dimensioned" here; that difference is booked in B159, not fixed by this knife.
    Debug.Print "R17-undim-vt="; VarType(q0n); " tn="; TypeName(q0n)

    Debug.Print "RE-DONE"
End Sub

' R14's second half: kept in its OWN procedure on purpose - R14 is the cross-procedure
' witness for Preserve (module-level array, shared descriptor), R15 is the same-
' procedure one. Both shapes matter: the first walks the module global, the second
' walks two generated bounds temps in one C function.
Private Sub GrowMp()
    ReDim Preserve mp(1 To 2, 1 To 6)
    mp(1, 3) = "later"
End Sub
