Attribute VB_Name = "Entry"
Option Explicit

Public Sub Main()
    Dim o As Impl
    Set o = New Impl
    Debug.Print o.IThing_DoWork(41)
End Sub
