VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Ident"
Option Explicit

' Third arm: the two readers the boxing leg must not break.
'  (a) "Is" compiles to C "==". The Object slot now holds a vb6_ComObject wrapper while the
'      typed variable still holds the raw instance, so identity has to be asked in one
'      representation or `If v Is o` answers false for one and the same object.
'  (b) The plain assignment reaches the server through vb6_ComCallInt while the comparison
'      goes through vb6_ComCall + VariantFromComResult; read the two apart.
'   idBit*100000 + cmpBit*1000 + latev  =>  101042 = both hold
'                                     100000 = identity ok, late binding dead
'                                     001042 = identity broke (the regression to watch)

Public Function Twice(ByVal n As Long) As Long
    Dim o As B187Good
    Dim v As Object
    Dim latev As Long
    Dim idBit As Long
    Dim cmpBit As Long
    Set o = New B187Good
    Set v = o
    If v Is o Then idBit = 1
    On Error Resume Next
    latev = v.Twice(n)
    If v.Twice(n) = 42 Then cmpBit = 1
    Twice = idBit * 100000 + cmpBit * 1000 + latev
End Function
