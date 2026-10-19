Attribute VB_Name = "B190Main"
Option Explicit

' Ledger 305 knife 89 (#408) driver. knownClassVars_ maps a LOWERCASE NAME to a class name and
' the back end never clears it, so a formal parameter registered as a class instance in one
' procedure answered for a DIFFERENT procedure that declared the same name as Object. Two faces,
' both measured on this project:
'   P3/P4 - the member read went out as Thing->Alpha while the C parameter type is void*:
'           cl C2223, the whole project stopped building (BASE has no exe at all).
'   P6    - knife 88 folded TypeName(<that parameter>) to the OTHER procedure's class name,
'           so the comparison is a constant and is false for Nothing.
' P5 is the non-regression witness: a module-level class variable keeps its native field read
' even after a same-named Object parameter shadowed it in an earlier procedure.

Private Thing As B190A

Private mOA As Object

Private Function PropViaObj(ByVal Item As Object) As String
    PropViaObj = Item.Label
End Function

Private Function TnViaObj(ByVal Item As Object) As String
    TnViaObj = TypeName(Item)
End Function

Private Function PropOnModuleSlot() As String
    PropOnModuleSlot = mOA.Label
End Function

Private Function UsesA(ByVal Item As B190A) As String
    UsesA = Item.Alpha
End Function

Private Function UsesB(ByVal Item As B190B) As String
    UsesB = Item.Beta
End Function

Private Function ShadowParam(ByVal Item As Object) As String
    ShadowParam = Item.Alpha
End Function

Private Function UsesObj(ByVal Thing As Object) As String
    UsesObj = Thing.Alpha
End Function

Private Function TnShadow(ByVal Thing As Object) As String
    TnShadow = TypeName(Thing)
End Function

Public Function UsesModuleVar() As String
    UsesModuleVar = Thing.Alpha
End Function

Public Sub Main()
    Dim a As New B190A
    Dim b As New B190B
    a.Alpha = "A1"
    b.Beta = "B2"
    Set Thing = a
    Debug.Print "C3ProbeB190-P1-classparam=" & UsesA(a)
    Debug.Print "C3ProbeB190-P2-classparamB=" & UsesB(b)
    Debug.Print "C3ProbeB190-P3-objectparam=" & ShadowParam(a)
    Debug.Print "C3ProbeB190-P4-objectparamThing=" & UsesObj(a)
    Debug.Print "C3ProbeB190-P5-modulevar=" & UsesModuleVar()
    Debug.Print "C3ProbeB190-P6-typename-nothing=" & TnShadow(Nothing)
    Debug.Print "C3ProbeB190-P7-typename-set=" & TnShadow(a)
    Debug.Print "C3ProbeB190-P8-typename-shadowB=" & TnShadow(b)
    Set mOA = a
    Debug.Print "C3ProbeB190-P9-propviaobj=" & PropViaObj(a)
    Debug.Print "C3ProbeB190-P10-modslotprop=" & PropOnModuleSlot()
    Debug.Print "C3ProbeB190-P11-tnmodslot=" & TnViaObj(mOA)
    Debug.Print "C3ProbeB190-P12-tn-direct=" & TypeName(mOA)
    Debug.Print "C3ProbeB190-DONE"
End Sub
