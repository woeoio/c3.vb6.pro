Attribute VB_Name = "Module1"
Option Explicit

' VB6 CLIENT side of the probe: late-bound call into the registered BoolProbe.ProbeCls.
' Writes the readings to %TEMP%\boolprobe_client.txt (and shows them in a MsgBox).
Sub Main()
    Dim o As Object
    Dim x As Long
    Dim r As Long
    Dim p As String
    Dim f As Integer
    Set o = CreateObject("BoolProbe.ProbeCls")
    p = ""
    p = p & "B1 EchoStr(True)  = " & o.EchoStr(True) & vbCrLf
    p = p & "B2 EchoStr(False) = " & o.EchoStr(False) & vbCrLf
    p = p & "B3 EchoLng(True)  = " & o.EchoLng(True) & vbCrLf
    x = 10
    r = o.Bump(x)
    p = p & "B4 Bump: x=" & x & " r=" & r & vbCrLf
    f = FreeFile
    Open Environ$("TEMP") & "\boolprobe_client.txt" For Output As #f
    Print #f, p
    Close #f
    MsgBox p, vbInformation, "boolprobe client"
End Sub
