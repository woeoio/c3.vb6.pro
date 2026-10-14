' Ledger 248 = handing a floating value to an integer target rounds now, and it rounds
' through ONE exit that CLng/CInt already used. Before, codegen fed the double straight
' into vb6_ChkLong(int64_t), so C truncated at the call boundary: `l = 6.73` answered 6
' and `l = 7 / 2` answered 3, while the *same* expression written `l = CLng(7 / 2)`
' answered 4 (vb6_CLng calls round()). Two answers for one decision = the family this
' repo keeps finding (#234 acquire-DC, #235 pen colour, #239 pen position, #247 geometry
' in the host model). Corpus exposure measured over the 90 emit captures: 36 sites
' (21 of them in Charts 2020's main project) store a division/float result into a narrow
' integer, so this was silently off by one in real layout math.
'
' F2L01..F2L07   values VB6 also rounds this way (no exact .5 tie involved)
' F2L08          the loop form of the same claim, deliberately built so truncation and
'                rounding CANNOT agree (x.75 values) -- the first version of this judge
'                used k/(2k+1), which is always below .5, and won on the broken compiler
' F2L09/F2L10    the invariant this cut creates: implicit narrowing == explicit
'                CLng/CInt, including the exact .5 ties where this repo deliberately
'                keeps half-away-from-zero (rev37 note in vb6rtl_conv.c; VB6 uses
'                banker's rounding there -- that difference is a separate account, which
'                is why ties are pinned as "same answer", not as "VB6's answer")
' F2L11          the overflow net still fires (this cut must not widen or lose it)
' F2L12/F2L13    guards: integer-only math and a Double target must not move
'
' Ledger 261 = the OTHER half of that same decision. The rounding exit existed, but the
' question "is this lvalue a narrow integer at all?" only ever asked bare identifiers
' (narrowCheckAssign started with `target->kind != IdentifierExpr` -> hand back the value).
' So every member write and every array element write kept truncating: `r.Left = 6.73`
' stored 6, `arr(0) = (4*1+3)/4` stored 1 -- while the identical value into a scalar
' variable answered 7 and 2. One decision, two answers again, and it is silent: both
' forms compile, both link, only the exact .75/.5 readings tell them apart.
' F2L14..F2L22   the new target shapes (UDT member Long/Integer/Byte, a member fed by a
'                Double expression, a module-level UDT member, static Long / static Byte /
'                dynamic Long array elements, an element of an array-of-UDT)
' F2L23          the invariant, now for a member: implicit narrowing == explicit CLng.
'                Deliberately pinned as "same answer" and NOT as "VB6's answer" -- the
'                exact .5 tie direction is still an open account (ledger 248 tail).
' F2L24..F2L26   guards on the new surface: integer source into a member/array must not
'                move, and a Double member must stay Double (the resolver answering
'                Unknown is what keeps those shapes byte-identical to before)
Option Explicit

Private Type RectF2L
    Px As Long
    Pi As Integer
    Pb As Byte
    Pd As Double
End Type

Private Type PicF2L
    Pixels(4) As Byte
    Nums() As Long
    Grid(1, 1) As Byte
End Type

Private m_r As RectF2L
Private m_pic As PicF2L

Private Function TF(ByVal b As Boolean) As String
    If b Then TF = "True" Else TF = "False"
End Function

Private Sub Chk(ByVal tag As String, ByVal got As Long, ByVal want As Long)
    Debug.Print "F2L-" & tag & "=" & TF(got = want) & " raw=" & CStr(got) & "," & CStr(want)
End Sub

Public Sub Main()
    Dim d As Double
    Dim s As Single
    Dim l As Long
    Dim i As Integer
    Dim same As Boolean
    Dim k As Long

    d = 6.73
    l = d
    Chk "01-dbl-round-up", l, 7

    l = -6.73
    Chk "02-dbl-round-neg", l, -7

    l = 7 / 2
    Chk "03-div-half-up", l, 4

    l = 5 / 3
    Chk "04-div-third", l, 2

    s = 2.4
    l = s
    Chk "05-single-down", l, 2

    s = -2.6
    l = s
    Chk "06-single-neg-up", l, -3

    i = 6.73
    Chk "07-integer-target", CLng(i), 7

    same = True
    For k = 1 To 8
        l = (4 * k + 3) / 4          ' 1.75, 2.75, ... 8.75: rounding 2..9, truncation 1..8
        If l <> k + 1 Then same = False
    Next k
    Chk "08-rounds-not-truncates", CLng(same), -1

    same = True
    For k = 1 To 8
        l = (2 * k + 1) / 2          ' exact .5 ties: 1.5, 2.5, ... 8.5
        If l <> CLng((2 * k + 1) / 2) Then same = False
    Next k
    Chk "09-parity-clng-ties", CLng(same), -1

    same = True
    For k = 1 To 8
        i = (4 * k + 3) / 4
        If CLng(i) <> CInt((4 * k + 3) / 4) Then same = False
    Next k
    Chk "10-parity-cint", CLng(same), -1

    On Error Resume Next
    d = 1.0E+12
    l = d / 1.0
    Dim e6 As Long
    e6 = Err.Number
    Err.Clear
    On Error GoTo 0
    Chk "11-overflow-still-6", e6, 6

    l = 7 \ 2
    Chk "12-intdiv-untouched", l, 3

    d = 7 / 2
    Chk "13-double-untouched", CLng(d * 10), 35

    ' ---- ledger 261: the target side of the same decision ----
    Dim r As RectF2L
    Dim arr(3) As Long
    Dim bArr(3) As Byte
    Dim dArr() As Long
    Dim ur(2) As RectF2L
    ReDim dArr(2)

    r.Px = 6.73
    Chk "14-member-long", r.Px, 7

    r.Pi = (4 * 1 + 3) / 4
    Chk "15-member-integer", CLng(r.Pi), 2

    r.Pb = 6.73
    Chk "16-member-byte", CLng(r.Pb), 7

    r.Px = r.Px + 0.73
    Chk "17-member-from-double-expr", r.Px, 8

    m_r.Px = (4 * 2 + 3) / 4
    Chk "18-module-udt-member", m_r.Px, 3

    arr(0) = 6.73
    Chk "19-static-array-long", arr(0), 7

    bArr(1) = 6.73
    Chk "20-static-array-byte", CLng(bArr(1)), 7

    dArr(1) = (4 * 3 + 3) / 4
    Chk "21-dyn-array-long", dArr(1), 4

    ur(0).Px = 6.73
    Chk "22-array-of-udt-member", ur(0).Px, 7

    same = True
    For k = 1 To 8
        r.Px = (2 * k + 1) / 2               ' exact .5 ties into a member
        m_r.Px = CLng((2 * k + 1) / 2)       ' same value, written explicitly
        If r.Px <> m_r.Px Then same = False
    Next k
    Chk "23-parity-member-ties", CLng(same), -1

    r.Px = 7 \ 2
    Chk "24-member-intdiv-untouched", r.Px, 3

    arr(2) = 5 * 5
    Chk "25-array-int-source-untouched", arr(2), 25

    r.Pd = 6.73
    Chk "26-member-double-untouched", CLng(r.Pd * 10), 67

    ' F2L27: the empty-parens lvalue is NOT an element slot -- `dst() = src()` assigns the
    ' whole array (Fix 170) and the emitted C hands back the array descriptor pointer. Putting a
    ' scalar gate around that pointer is exactly rev36's run-time error 6 arriving by a new door
    ' (measured: with the guard missing, test_array.bas exits 0x00000006 and its four wa-* lines
    ' never print). So the array tier asks "is there an index" before "what is the element type".
    Dim src2(2) As Long
    Dim dst2() As Long
    ReDim dst2(2)
    src2(0) = 10
    src2(1) = 20
    src2(2) = 22
    dst2() = src2()
    Chk "27-whole-array-assign-untouched", dst2(2), 22

    ' F2L28: an element of a *member* array (`p.Pixels(0)`). Ledger 261 deliberately left this
    ' shape alone because the callee is not a bare identifier AND, more importantly, because the
    ' lvalue itself was wrong (multi-D members were folded to one-D). Ledger 262 fixed the lvalue,
    ' so the store now goes through the same round-and-check source as every other narrow slot:
    ' 6.73 -> 7. (It read 6 before 262 -- PINNED AS 7 HERE so a re-fold of that boundary goes red.)
    Dim p As PicF2L
    p.Pixels(0) = 6.73
    Chk "28-member-array-elem-rounded", CLng(p.Pixels(0)), 7

    ' F2L29/30: WHOLE-member-array assignment -- the shape that bit this knife in CI-prep.
    ' `Data` / `Nums` are array members, so the lvalue is the array itself (a descriptor
    ' pointer), not a scalar slot; inferUdtFieldVb6Type answers the ELEMENT tier (the array-ness
    ' lives in mi.isArrayDynamic / mi.arraySize, not in a Vb6Type flag), so a tier test alone
    ' cannot see it. Wrapping that pointer in vb6_ChkByte is rev36's error 6 through a new door:
    ' measured on VbQRCodegen's Project1 -- BUILD-RC=0 but the exe died at startup with
    ' "Unhandled VB6 Error #6: Overflow" while the same source on the pre-guard compiler sat in
    ' its message loop like always. Both spellings (bare and empty-parens) must stay unwrapped.
    Dim src3(1) As Long
    src3(0) = 7
    src3(1) = 8
    p.Nums = src3
    Chk "29-member-array-whole-assign", p.Nums(1), 8
    p.Nums() = src3()
    Chk "30-member-array-whole-assign-parens", p.Nums(1), 8

    ' F2L31/32 (ledger 262 made these reachable): elements of a MULTI-D member array. Before 262
    ' `Grid(0, 0)` and `Grid(0, 1)` were the SAME C slot, so a rounding pin here could not tell
    ' "rounds" from "collapses". Both cells now exist, both take the wrap, and they keep their own
    ' values -- 6.73 -> 7 in one, 2.4 -> 2 next to it.
    p.Grid(0, 1) = 6.73
    p.Grid(0, 0) = 2.4
    Chk "31-member-2d-elem-rounded", CLng(p.Grid(0, 1)), 7
    Chk "32-member-2d-elem-neighbour", CLng(p.Grid(0, 0)), 2

    Debug.Print "F2L-DONE"
End Sub
