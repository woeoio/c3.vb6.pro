Attribute VB_Name = "ModProbe"
Option Explicit

' Fuguizhu FIX-011 probe: early-bound (typelib signature) + [in,out] VARIANT* params.
' Before fix: knownVariantVars_ bailed out -> packed by value (vb6_ComPackValue)
'   -> server wrote into its own copy, caller MinPt/MaxPt stayed Empty
'   (same root cause as the GetBoundingBox/GetXData family, 92 sites upstream).
' After fix: vb6_ComPackVariantRef(&MinPt) + vb6_ComVariantOuts(call, 2, &MinPt, &MaxPt).
'
' GetInOnly is the [in] control arm: PARAMFLAG_FIN alone maps to ComParamDir::In
'   -> isByVal=true -> must NOT be emitted as by-ref (a pure input param passed
'   as writable lets the server scribble into the caller's input slot).
Sub Main()
    Dim b As FGZOUTTLB.FgzBox
    Set b = New FGZOUTTLB.FgzBox
    Dim MinPt As Variant, MaxPt As Variant
    b.GetBox MinPt, MaxPt
    Debug.Print MinPt(0), MaxPt(0)
    Dim V As Variant
    b.GetScalar V
    ' [in] control arm: must be packed by value, not by reference.
    Dim Vin As Variant
    Vin = 7
    b.GetInOnly Vin
    Debug.Print "VIN=" & CStr(Vin)
    ' Scalar ByRef control arms: the five-tier ByRef packers must be unchanged
    ' by the fix (Long/Integer/Boolean/Single/Double).
    Dim nL As Long, nI As Integer, bB As Boolean, nS As Single, nD As Double
    nL = 1: nI = 2: bB = True: nS = 3#: nD = 4#
    b.PutLong nL
    b.PutInteger nI
    b.PutBool bB
    b.PutSingle nS
    b.PutDouble nD
    Debug.Print "SC=" & CStr(nL) & "," & CStr(nI) & "," & CStr(bB) & "," & CStr(nS) & "," & CStr(nD)
End Sub
