VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneDCls"
Option Explicit

' Ledger 305 B223 (knife 223 measure): control shape -- module-level 1-D FIXED String array. This is the
' one knife 222 fixed, so on any compiler at or after 222 it must survive create+Put+Release.

Private f(3) As String

Public Function Put(ByVal s As String) As String
    f(0) = s
    Put = f(0)
End Function
