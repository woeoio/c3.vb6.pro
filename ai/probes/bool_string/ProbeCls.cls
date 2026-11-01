VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProbeCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Automation probe: what does a VB6 SERVER see when a client hand it each variant type?
' Build this project as BoolProbe.dll, register it, then run probe.vbs and send back the output.

Public Function EchoStr(ByVal s As String) As String
    EchoStr = "[" & s & "]"
End Function

Public Function EchoLng(ByVal n As Long) As Long
    EchoLng = n
End Function

Public Function EchoVar(ByVal v As Variant) As String
    EchoVar = TypeName(v) & "=" & CStr(v)
End Function

' ByRef write-back: the caller variable must become 11 if the call really is by reference.
Public Function Bump(ByRef v As Long) As Long
    v = v + 1
    Bump = v
End Function
