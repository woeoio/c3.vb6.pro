Attribute VB_Name = "B169Main"
Option Explicit

Sub Main()
    Dim h As B169Holder
    Dim s As String
    Dim p As String
    Dim w As String

    Set h = New B169Holder

    ' (1) cross-object String field write from a live variable: the field has to keep
    '     its own copy, so changing the source afterwards must not change the field.
    s = "aaa"
    h.PS = s
    s = "bbb"
    Debug.Print "B169-1=" & h.PS & "|" & Len(h.PS)

    ' (2) two literal writes in a row: the old value has to be released (else leak).
    h.PS = "ccc"
    h.PS = "ddd"
    Debug.Print "B169-2=" & h.PS & "|" & Len(h.PS)

    ' (3) in-class write through a String parameter: must not alias the caller's string.
    p = "eee"
    h.StoreViaMe p
    p = "fff"
    Debug.Print "B169-3=" & h.ReadBack() & "|" & p

    ' (4) cross-object With block: all four field kinds written through the With road.
    w = "ggg"
    With h
        .PS = w
        .PL = 9
        .PV = 5
        .PK = kNumber
    End With
    w = "hhh"
    Debug.Print "B169-4=" & h.ReadBack()
    Debug.Print "B169-5=" & h.PS & "|" & CStr(h.PL) & "|" & CStr(h.PK)

    Debug.Print "B169-DONE"
End Sub
