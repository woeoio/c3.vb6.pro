Option Explicit

' Ledgers 262 + 265. A UDT member array with two or more dimensions used to lose every
' dimension but the first: the parser parsed the extra subscripts and threw them away
' ("keep the first bound, discard the rest"), so `M(3, 3) As Long` emitted `int32_t M[4]`
' and `m.M(0, 1)` / `m.M(0, 2)` / `m.M(0, 3)` all collapsed to `m.M[0]`. Three different
' cells, one cell of memory, no error. Real traffic: the GDI+ colour matrix in
' tests/Charts 2020/LabelPlus/LabelPlus.ctl (`M(0 To 4, 0 To 4) As Single`) -- GDI+ reads
' 100 bytes out of a 20-byte struct, and the shadow-opacity write and the picture-opacity
' write landed on the same float.
'
' Ledger 265 is the second thing this file measures: a UDT that HAS a fixed String (or
' Variant) array member and gets copied anywhere (assignment, ByRef, return) had its
' auto-generated deep-copy helper emit `sizeof(d->S[_i])` in the same statement that
' declares `_i` -- C2065, the project does not build. Nothing in the corpus declared that
' shape, which is why it never rang.
'
' Every judgment below writes cells that share the FIRST subscript and differ in the
' second: that is exactly the shape the old fold collapsed. The BASE compiler answers
' 33/33/33, 222/222, 3/3 and Len=16; the fixed one answers the distinct values and Len=64.
'
'   MD-LONG / MD-SINGLE / MD-R3   per-rank element identity (4x4 Long, 5x3 Single, 2x3x4 Long)
'   MD-ONE                        a 1-D member array still folds as before (regression guard)
'   MD-WITH                       the `With mat : .M(i, j)` road, same authority as `mat.M(i, j)`
'   MD-MODULE                     module-level UDT variable (third target road)
'   MD-STR / MD-STR-COPY          the 2-D String member, before and after a whole-UDT copy
'   MD-LEN                        sizeof the struct and of the field -- layout, not reading
'   MD-ALIAS=N                    invariant: the three Long cells are pairwise different
'   MD-DONE                       all judgments were written
' ASCII only.

Public Type MatL
    M(3, 3) As Long
End Type

Public Type MatS
    K(0 To 4, 0 To 2) As Single
End Type

Public Type MatR
    T(1, 2, 3) As Long
End Type

Public Type MatOne
    B(3) As Long
End Type

Public Type MatStr
    S(1, 2) As String
End Type

Public g_m As MatL

Public Sub Main()
    Dim m As MatL
    m.M(0, 1) = 11
    m.M(0, 2) = 22
    m.M(0, 3) = 33
    Debug.Print "MD-LONG=" & m.M(0, 1) & "/" & m.M(0, 2) & "/" & m.M(0, 3)

    Dim s As MatS
    s.K(1, 1) = 111
    s.K(1, 2) = 222
    Debug.Print "MD-SINGLE=" & s.K(1, 1) & "/" & s.K(1, 2)

    Dim t As MatR
    t.T(0, 0, 0) = 1
    t.T(0, 0, 3) = 3
    t.T(1, 0, 0) = 40
    Debug.Print "MD-R3=" & t.T(0, 0, 0) & "/" & t.T(0, 0, 3) & "/" & t.T(1, 0, 0)

    Dim one As MatOne
    one.B(2) = 9
    Debug.Print "MD-ONE=" & one.B(2)

    With m
        .M(0, 2) = 77
        Debug.Print "MD-WITH=" & .M(0, 1) & "/" & .M(0, 2)
    End With

    g_m.M(0, 1) = 44
    g_m.M(0, 2) = 55
    Debug.Print "MD-MODULE=" & g_m.M(0, 1) & "/" & g_m.M(0, 2)

    Dim w As MatStr
    w.S(0, 1) = "ab"
    w.S(0, 2) = "cd"
    Debug.Print "MD-STR=" & w.S(0, 1) & "/" & w.S(0, 2)
    Dim cp As MatStr
    cp = w
    Debug.Print "MD-STR-COPY=" & cp.S(0, 1) & "/" & cp.S(0, 2)

    Debug.Print "MD-LEN=" & Len(m) & "/" & Len(m.M) & "/" & Len(one.B)

    Dim sameQ As String
    If m.M(0, 1) = m.M(0, 2) Or m.M(0, 2) = m.M(0, 3) Or m.M(0, 1) = m.M(0, 3) Then
        sameQ = "Y"
    Else
        sameQ = "N"
    End If
    Debug.Print "MD-ALIAS=" & sameQ

    Debug.Print "MD-DONE"
End Sub
