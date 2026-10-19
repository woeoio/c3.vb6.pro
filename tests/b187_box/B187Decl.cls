VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Decl"
Option Explicit

' Ledger 305 B189 tier two (the eighth arm). TypeName of a variable whose DECLARATION names a
' project class: VB6 answers that class's module name. Before this knife the emitted call handed
' the raw instance to vb6_TypeName, which only knows the host layer and the vt name table, so the
' Set-then-read shape answered "Object". The answer now folds at compile time out of the
' declaration -- which is exactly why it must NOT fire for a class that anything in the project
' Implements: there the declaration is a VIEW and VB6 answers the concrete class. That half is
' witnessed in tests/b189_gate (emit-only -- see the note in run_tests.ps1).
'
' One arm, four shapes, each coded as Len(name)*10 + (name is the one VB6 answers):
'   tg = As B187Good, Set        BASE 60 ("Object")   / this knife 81 ("B187Good")
'   tg = the same variable, Set back to Nothing   BASE 71 ("Nothing") / this knife 71 (must not move)
'   tn = As B187Good, never Set  BASE 71              / this knife 71 -- the NULL head of the fold
'   tx = As B187Name, Set        BASE 60              / this knife 81 ("B187Name") -- the folded name
'                                                    tracks the DECLARATION, it is not one constant
' Twice = n*2*10000000 + a*1000000 + b*10000 + c*100 + d   (the probe calls Twice(21))

Public Function Twice(ByVal n As Long) As Long
    Dim tg As B187Good
    Dim tn As B187Good
    Dim tx As B187Name
    Dim s As String
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long

    Set tg = New B187Good
    s = TypeName(tg)
    a = Len(s) * 10
    If s = "B187Good" Then a = a + 1

    Set tg = Nothing
    s = TypeName(tg)
    b = Len(s) * 10
    If s = "Nothing" Then b = b + 1

    s = TypeName(tn)
    c = Len(s) * 10
    If s = "Nothing" Then c = c + 1

    Set tx = New B187Name
    s = TypeName(tx)
    d = Len(s) * 10
    If s = "B187Name" Then d = d + 1

    Twice = n * 2 * 10000000 + a * 1000000 + b * 10000 + c * 100 + d
End Function
