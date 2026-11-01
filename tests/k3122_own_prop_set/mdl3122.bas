Attribute VB_Name = "mdl3122"
Option Explicit

Public Sub Main()
    Dim p As P3122
    Set p = New P3122
    p.LoadIt
    Debug.Print "K3122-p=" & (p.Pic Is Nothing)
    p.LoadRaw
    Debug.Print "K3122-r=" & (p.Pic Is Nothing)
    Debug.Print "K3122-DONE"
End Sub
