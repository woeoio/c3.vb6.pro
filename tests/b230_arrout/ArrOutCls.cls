VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrOutCls"
Option Explicit

' Ledger 305 SB222 (knife 230): the out-array slot. An automation client passes a() ByRef,
' i.e. VT_BYREF plus VT_ARRAY plus VT_I4. The server copies it into its own 1D descriptor,
' the VB body ReDims it (destroy + fresh carrier) and fills it, and the write-back must turn
' that carrier into a real SAFEARRAY and store it through the client's pparray -- otherwise
' the client keeps the array it sent in and every call leaks the carrier the body built.

Public Function Fill(ByVal n As Long, a() As Long) As Long
    Dim i As Long
    ReDim a(n - 1)
    For i = 0 To n - 1
        a(i) = (i + 1) * 7
    Next i
    Fill = n
End Function
