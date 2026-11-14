' Account #611 knife 348: Array() element variant-kind fidelity.
' VB6 semantics: Array(True/False) is VT_BOOL(11), Null=1, Empty=0,
' Nothing=9 (VT_DISPATCH with a null pointer), a nested Array() is
' VT_ARRAY|VT_VARIANT(8204), and a date literal is VT_DATE(7).
' Before this knife every one of them went through vb6_ArraySetLong (VT_I4=3),
' and Nothing / a nested array pushed a pointer into an int32 parameter.
' The plain Long/BSTR/Double arms below are the zero-regression guard:
' their emitted setter must stay byte-for-byte what it was.
Option Explicit

Public Sub Main()
    Dim a As Variant
    Dim v As Variant
    a = Array(True, False, Null, Empty, Nothing)
    v = a(0): Debug.Print "AL-bt="; VarType(v)
    v = a(1): Debug.Print "AL-bf="; VarType(v)
    v = a(2): Debug.Print "AL-null="; VarType(v)
    v = a(3): Debug.Print "AL-empty="; VarType(v)
    v = a(4): Debug.Print "AL-nothing="; VarType(v)

    Dim b As Variant
    Dim w As Variant
    b = Array(Array(1, 2), 3)
    v = b(0): Debug.Print "AL-nested="; VarType(v)
    w = v(1): Debug.Print "AL-nested-val="; w
    v = b(1): Debug.Print "AL-long="; VarType(v)

    Dim d As Variant
    d = Array(#1/2/2000#)
    v = d(0): Debug.Print "AL-date="; VarType(v)

    Dim c As Variant
    c = Array(7, "abc", 1.5)
    v = c(0): Debug.Print "AL-i4="; VarType(v)
    v = c(1): Debug.Print "AL-bstr="; VarType(v)
    v = c(2): Debug.Print "AL-r8="; VarType(v)

    Debug.Print "AL-DONE"
End Sub
