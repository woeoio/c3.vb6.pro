Attribute VB_Name = "B168Main"
Option Explicit

' 账 #305 §B168 (第五十九刀). Each W row pins one half of the judgement:
'   W1      scalar into the slot      - BASE: raw store, cl C2440, no exe exists
'   W2      borrowed string in        - the slot keeps its own copy, source may die
'   W3      re-assign twice           - latest wins, previous value reclaimed
'   W4      slot to slot              - no aliasing with the caller's own slot
'   W5/W6   in-class Me.PV = ...      - same answers from inside the class
'   W7      implicit Variant field    - the no-As-clause road
'   W8      String/Long fields        - negative control, must stay unboxed
'   W9      numeric read              - witness through the plain CStr road
' String values are read back through the class's own ReadBack() because reading
' CStr(<obj>.<Variant field>) from outside goes down a different, still-open road
' (台账 §B168c) - that question is not this knife's patient.

Sub Main()
    Dim h As B168Holder
    Dim s As String
    Dim v As Variant
    Dim w As Variant
    Set h = New B168Holder

    h.PV = 7
    Debug.Print "B168-W1=" & CStr(h.PV)

    s = "sharedbuf"
    h.PV = s
    s = "gone"
    Debug.Print "B168-W2=" & h.ReadBack() & "|" & Len(h.ReadBack())

    h.PV = "first"
    h.PV = "second"
    Debug.Print "B168-W3=" & h.ReadBack() & "|" & Len(h.ReadBack())

    w = "wide"
    h.PV = w
    w = 123
    Debug.Print "B168-W4=" & h.ReadBack() & "|" & Len(h.ReadBack())

    v = "callerslot"
    h.StoreViaMe v
    v = "changed"
    Debug.Print "B168-W5=" & h.ReadBack() & "|" & Len(h.ReadBack())

    h.StoreLitViaMe
    Debug.Print "B168-W6=" & h.ReadBack() & "|" & Len(h.ReadBack())

    h.PW = "noclause"
    Debug.Print "B168-W7=" & CStr(h.PW) & "|" & Len(CStr(h.PW))

    h.PS = "plainstring"
    h.PL = 42
    Debug.Print "B168-W8=" & h.PS & "|" & Len(h.PS) & "|" & CStr(h.PL)

    h.PV = 256
    Debug.Print "B168-W9=" & CStr(h.PV) & "|" & h.ReadBack()
    Debug.Print "B168-DONE"
End Sub
