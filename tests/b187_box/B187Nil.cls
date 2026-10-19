VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Nil"
Option Explicit

' Ledger 305 B189 second tier (the seventh arm). An object variable that references nothing
' is not "Object" in VB6 -- TypeName answers "Nothing", and real code uses that as a guard
' (eight sites in the VB6 corpus: tree-walk loops and empty-object checks). Before this
' knife the NULL-dispatch shape fell through to the scalar vt table and answered "Object".
'
' One arm, three shapes, each coded as Len(name)*10 + (name is the one VB6 answers):
'   o  = a late-bound slot never assigned      BASE 60 ("Object") / this knife 71 ("Nothing")
'   g  = a typed instance, Set then Nothing    BASE 60             / this knife 71
'   v  = a Variant never assigned (must NOT move; VB6 "Empty")    51 on both
' Twice = n*2*1000000 + a*10000 + b*100 + c
'   BASE (knife 86) = 42606051, this knife = 42717151.

Public Function Twice(ByVal n As Long) As Long
    Dim o As Object
    Dim g As B187Good
    Dim v As Variant
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    Dim a As Long
    Dim b As Long
    Dim c As Long

    Set g = New B187Good
    Set g = Nothing

    s1 = TypeName(o)
    s2 = TypeName(g)
    s3 = TypeName(v)

    a = Len(s1) * 10
    If s1 = "Nothing" Then a = a + 1
    b = Len(s2) * 10
    If s2 = "Nothing" Then b = b + 1
    c = Len(s3) * 10
    If s3 = "Empty" Then c = c + 1

    Twice = n * 2 * 1000000 + a * 10000 + b * 100 + c
End Function
