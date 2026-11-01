Attribute VB_Name = "mdl294"
Option Explicit

Public Sub Main()
    Dim c As Collection
    Set c = New Collection
    c.Add 1
    c.Add 2
    c.Add 3
    Dim o As New P294
    o.Index = c.Count
    Debug.Print "K294-idx=" & o.Index
    Debug.Print "K294-DONE"
End Sub
