Option Explicit

' Account #611 probe: does a Variant **array** survive a round trip through a real
' external COM object (Scripting.Dictionary)?
'   1D  -> vb6_VariantComSlotToRealArray converts the C3 descriptor into a real
'         OLE SAFEARRAY, so this leg is expected to work.
'   2D  -> the C3 ND descriptor used to be handed over as-is (it is NOT an OLE
'         SAFEARRAY), and reading it back answered VT_EMPTY.
Sub Main()
    Dim d As Object
    Dim a1 As Variant
    Dim a2(0 To 1, 0 To 1) As Variant
    Set d = CreateObject("Scripting.Dictionary")

    a1 = Array("x", "y")
    d.Add "one", a1

    a2(0, 0) = "k1": a2(1, 0) = "v1"
    a2(0, 1) = "k2": a2(1, 1) = "v2"
    d.Add "two", a2

    Debug.Print "count=" & d.Count
    Debug.Print "v1d=" & VarType(d("one")) & "/" & UBound(d("one")) & "/" & d("one")(1)
    Debug.Print "v2d=" & VarType(d("two"))
End Sub
