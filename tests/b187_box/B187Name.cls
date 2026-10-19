VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Name"
Option Explicit

' Ledger 305 B189 first tier (the sixth arm). The boxing leg hands a late-bound slot a real
' vb6_ComObject wrapper, so by now the slot knows which project class it came from -- but
' TypeName still answered "Object", because vb6_TypeName only asks the host layer (HWND /
' control) and then falls through to the VT_DISPATCH name table. The wrapper is the one of
' the three measured shapes whose class name lives in a runtime value we already trust
' (desc->classVariable), so this arm reads exactly that.
'
' Encoding: n*2*1000 + Len(TypeName(v))*10 + (name == "B187Good" ? 1 : 0)
'   BASE (knife 85, no class-name exit) = 42000 + 60 + 0 = 42060  ("Object")
'   this knife                          = 42000 + 80 + 1 = 42081  ("B187Good")
' The Len term is the witness that nothing else was substituted: a wrong-but-nonempty name
' moves the middle digits, so 81 only prints for a length-8 name that also matches.

Public Function Twice(ByVal n As Long) As Long
    Dim g As B187Good
    Dim v As Object
    Dim nm As String
    Dim ok As Long

    Set g = New B187Good
    Set v = g
    nm = TypeName(v)
    If nm = "B187Good" Then ok = 1

    Twice = n * 2 * 1000 + Len(nm) * 10 + ok
End Function
