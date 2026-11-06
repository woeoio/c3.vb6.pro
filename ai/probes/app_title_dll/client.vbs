' 客户端：32 位 cscript 才建得出来（64 位会报 0x800A01AD，见 README 第 3 条坑）
'   C:\Windows\SysWOW64\cscript.exe //nologo client.vbs > client_out.txt 2>&1
' 然后把 client_out.txt 的五行原样贴回来。
Option Explicit

Dim o
Set o = CreateObject("AppTitleProbe.Probe")

WScript.Echo o.P1()
WScript.Echo o.P2()
WScript.Echo o.P3()
WScript.Echo o.P4()
WScript.Echo o.P5()
WScript.Echo o.P6()
WScript.Echo o.P7()
WScript.Echo o.P8()

Set o = Nothing
