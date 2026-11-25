Option Explicit

Public gSlot As Object
Public gKeep As Long

Sub Main()
    Dim o As Object
    Dim a As B412Item

    Set a = New B412Item
    a.Seed = 21
    Set o = a
    Debug.Print "C3ProbeB412-P1=" & TypeName(o)

    o.Seed = 22
    Debug.Print "C3ProbeB412-P2=" & CStr(o.Seed)

    Set gSlot = o
    Debug.Print "C3ProbeB412-P3=" & TypeName(gSlot)
    gSlot.Seed = 23
    Debug.Print "C3ProbeB412-P4=" & CStr(gSlot.Seed)

    Set o = Nothing
    Debug.Print "C3ProbeB412-P5=" & CStr(a.Seed)

    Set gSlot = Nothing
    Debug.Print "C3ProbeB412-P6=" & CStr(a.Seed)

    Set a = Nothing
    Debug.Print "C3ProbeB412-DONE"
End Sub
