Option Explicit

' Account #611: `Array("admin","auditor")` must be a 2-element Variant array with
' LBound=0 / UBound=1, and a ByRef Variant walk over it must visit BOTH elements.
' JSonFast button 7 (WriteTypedArrayFallback) printed only "admin" before the fix
' on the argument-passing side; this fixture pins the array shape itself.
Sub Main()
    Dim a As Variant
    Dim i As Long
    Dim s As String
    a = Array("admin", "auditor")
    Debug.Print "lb=" & LBound(a) & " ub=" & UBound(a)
    For i = LBound(a) To UBound(a)
        s = s & "[" & i & "=" & a(i) & "]"
    Next i
    Debug.Print "walk=" & s
End Sub
