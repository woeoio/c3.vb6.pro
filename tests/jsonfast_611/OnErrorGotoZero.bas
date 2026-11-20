Option Explicit

' Account #611: `On Error GoTo 0` disables ONLY the current procedure's handler.
' A raise after it must still reach the CALLER's On Error GoTo label.
' Before the fix C3 emitted `vb6_err_jmp_active = 0`, which killed the caller's
' setjmp frame too => vb6_ErrRaise took the "Unhandled error" branch and called
' ExitProcess (JSonFast button 5 killed the whole process on click).
Sub Main()
    Debug.Print "A=" & TryIt("ok")
    Debug.Print "B=" & TryIt("bad")
End Sub

Private Function TryIt(ByVal mode As String) As String
    On Error GoTo Failed
    Inner mode
    TryIt = "no-error"
    Exit Function
Failed:
    TryIt = "caught=" & CStr(Err.Number)
    Err.Clear
End Function

Private Sub Inner(ByVal mode As String)
    On Error GoTo LocalEH
    Dim x As Long
    x = 1
    On Error GoTo 0
    If mode = "bad" Then Err.Raise 5011, "probe", "boom"
    Exit Sub
LocalEH:
    Err.Raise 5999
End Sub
