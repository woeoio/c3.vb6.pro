VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NdLongCls"
Option Explicit

' Ledger 305 B223 (knife 223 measure): module-level 2-D FIXED Long array. With the code as of
' knife 222 the factory destroy hands this carrier to vb6_SA_Destroy, which casts it to
' vb6_SafeArray1D* -- a completely different layout (ND: dimCount@0 / elemType@4 / elemSize@8 /
' totalElements@12 / data@16; 1D: cDims@0 / fFeatures@2 / elemSize@4 / elemType@8 / data@16 on x64,
' data@12 on x86). This class is the witness for what that cast actually does.

Private nds(1 To 2, 3) As Long

Public Function Put(ByVal n As Long) As Long
    Dim k As Long
    k = n Mod 4
    If k < 0 Then k = 0
    nds(1, k) = n + 100
    Put = nds(1, k)
End Function

Public Function Total() As Long
    Dim i As Long, j As Long, t As Long
    For i = 1 To 2
        For j = 0 To 3
            t = t + nds(i, j)
        Next j
    Next i
    Total = t
End Function
