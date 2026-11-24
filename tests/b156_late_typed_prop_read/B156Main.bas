Option Explicit

Public Sub Main()
    Dim box As B156Box
    Dim o As Object
    Dim n As Long
    Dim s As String

    Set box = New B156Box
    Set o = box
    box.Held = 11
    box.Note = "hi"

    n = o.Held
    Debug.Print "B156-Q1=" & n
    n = o.Held + 1
    Debug.Print "B156-Q2=" & n
    If o.Held = 11 Then Debug.Print "B156-Q3=eq"
    s = o.Note
    Debug.Print "B156-Q4=" & s
    Debug.Print "B156-Q5=" & o.Held
    Set box = Nothing
    Set o = Nothing
    Debug.Print "B156-DONE"
End Sub
