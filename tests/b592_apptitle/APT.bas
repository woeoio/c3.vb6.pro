Attribute VB_Name = "APT"

Sub Main()
    Debug.Print "APT-seed=" & App.Title
    App.Title = "renamed"
    Debug.Print "APT-set=" & App.Title
    App.Title = ""
    Debug.Print "APT-empty=[" & App.Title & "]"
    Debug.Print "APT-exe=" & App.EXEName
    Debug.Print "APT-done"
End Sub
