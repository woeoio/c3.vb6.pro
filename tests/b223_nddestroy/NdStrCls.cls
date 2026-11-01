VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NdStrCls"
Option Explicit

' Ledger 305 B223 (knife 223 measure): module-level 2-D FIXED String array. Put() stores a real BSTR
' into an ND element, so the destroy leg has something to release: pre-222 it was SysFreeString on the
' carrier, post-222 it is vb6_SA_Destroy on the carrier (both wrong; the question is HOW wrong).

Private nds(1 To 2, 3) As String

Public Function Put(ByVal s As String) As String
    nds(1, 0) = s
    Put = nds(1, 0)
End Function

Public Function At1(ByVal n As Long) As String
    Dim k As Long
    k = n Mod 4
    If k < 0 Then k = 0
    At1 = nds(1, k)
End Function
