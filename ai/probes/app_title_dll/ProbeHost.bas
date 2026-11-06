Attribute VB_Name = "ProbeHost"
Option Explicit

' Ledger 605 follow-up anchor. AddressOf ProbeHost.Anchor is an address inside THIS DLL's code
' section, which is what GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS, ...) needs
' in order to answer "which module is this code running in". That is the independent witness
' for P6 below: it must be the DLL's own HMODULE, never the host process's.
Public Sub Anchor()
End Sub
