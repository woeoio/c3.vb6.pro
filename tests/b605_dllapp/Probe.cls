VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Probe"
Option Explicit

' Ledger 305 knife 330 (ledger 605): the runtime witness for whom the App object answers
' inside an ActiveX DLL. Real VB6 (ai/probes/app_title_dll/client_out.txt, host = 32-bit
' cscript) answers three ways, all about the DLL itself:
'   P1 App.Title   = this project's Title   (not the host process caption)
'   P3 App.EXEName = the DLL's own module name (not cscript)
'   P4 App.Path    = the DLL's own directory  (not cscript's)
' Before this knife C3 answered the HOST exe for all three: nothing seeded App.Title in
' dll_entry.c, and App.Path / App.EXEName asked GetModuleFileNameW(NULL, ...).
' The reading is taken by tests/tools/disp_probe.c through a registry-free in-proc client
' (CALL=ReadApp), so it needs no regsvr32 and nothing installed on the machine. The DLL is
' built into a marked sub-directory while the probe lives outside it -- whether that marker
' shows up in App.Path is the bit that separates answering the DLL from answering the host.

Public Function ReadApp() As String
    ReadApp = "605-title=[" & App.Title & "] exe=[" & App.EXEName & _
              "] tailOk=" & CStr(InStr(1, App.Path, "b605_self") > 0) & _
              "] path=[" & App.Path & "]"
End Function
