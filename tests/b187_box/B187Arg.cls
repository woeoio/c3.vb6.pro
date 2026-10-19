VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "B187Arg"
Option Explicit

' Ledger 305 B187 fourth leg (the argument station). The boxing leg only owns Set
' statements, so a class instance handed to a ByVal "As Object" FORMAL still arrived as a
' raw C pointer: the probe emitted
'     vb6_B89Driver_TakeIt((void*)me, o)
' and inside the callee p.Twice(21) went through vb6_ComCall, which asks the first word for
' an IDispatch vtable, gets none, and answers nothing -> 42000 (the early half, 42, is fine).
' The same encoding as the driver arm: early * 1000 + late => 42042 when the formal slot
' also answers a late-bound call, 42000 when it does not.

Public Function Twice(ByVal n As Long) As Long
    Dim o As New B187Good
    Dim argLate As Long
    argLate = TakeIt(o)
    Twice = n * 2 * 1000 + argLate
End Function

Private Function TakeIt(ByVal p As Object) As Long
    On Error Resume Next
    TakeIt = p.Twice(21)
End Function
