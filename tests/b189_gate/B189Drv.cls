VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B189Drv"
Option Explicit

' Both functions hit the SAME station -- TypeName of a variable whose declaration names a project
' class -- and differ only in that declaration. o's class is Implements-ed above (a view, fold must
' not fire); m's class is the implementor (fold must fire). Same file, same call, opposite answers:
' a fold that ignored the gate would print the ternary with L"B189Iface" in it here too.

Public Function F1(ByVal o As B189Iface) As String
    F1 = TypeName(o)
End Function

Public Function F2(ByVal m As B189Impl) As String
    F2 = TypeName(m)
End Function

' The fourth gate: a slot's TypeName is a RUNTIME question, so a Variant argument must never fold.
' F4/F3 share one identifier name on purpose -- that is how the real corpus case broke (the
' by-name class table hands F3 the declaration F4 made), so these two lines are the regression
' witness: F4 folds, F3 keeps asking vb6_TypeName.
Public Function F4(ByVal Value As B189Impl) As String
    F4 = TypeName(Value)
End Function

Public Function F3(ByVal Value As Variant) As String
    F3 = TypeName(Value)
End Function
