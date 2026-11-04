Attribute VB_Name = "ModuleExt"
Option Explicit

' External per-file manifest reuse: the ProgID below exists ONLY in the
' <component>.manifest block that the runtime composer pulls in. It is not in
' the ComLib= table, so this call must go through SxS (CLSIDFromProgID ->
' CoCreateInstance) or fail.
Sub Main()
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("ExtManifest.Calc")
    If Err.Number = 0 Then
        Debug.Print "ext-manifest-create=OK"
    Else
        Debug.Print "ext-manifest-create=FAIL"
    End If
End Sub
