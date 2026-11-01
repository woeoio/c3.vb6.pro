Option Explicit
Dim o, x, r
Set o = CreateObject("BoolProbe.ProbeCls")
WScript.Echo "A1 EchoStr(True)  = " & o.EchoStr(True)
WScript.Echo "A2 EchoStr(False) = " & o.EchoStr(False)
WScript.Echo "A3 EchoStr(12.75) = " & o.EchoStr(12.75)
WScript.Echo "A4 EchoStr(40)    = " & o.EchoStr(40)
WScript.Echo "A5 EchoLng(True)  = " & o.EchoLng(True)
WScript.Echo "A6 EchoVar(True)  = " & o.EchoVar(True)
WScript.Echo "A7 EchoVar(12.75) = " & o.EchoVar(12.75)
x = 10
r = o.Bump(x)
WScript.Echo "A8 Bump: x=" & x & " r=" & r
