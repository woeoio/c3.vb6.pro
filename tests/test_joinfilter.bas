' <vbeclipse>: Join/Filter 的数组槽对表。两条修复前形态:
'   Join(Array("abc",…))  -> 按 BSTR* 读 Variant 槽 → 解引用垃圾 → 0xC0000005
'   Filter(a, "ab")       -> 可选参没补 → cl C2198 参数太少, 根本编不过
' 读数全取自真实输出 (含 1-based 源数组与"零命中 = 空数组 UBound=-1"两形)。
Option Explicit

Public Sub Main()
    Dim v As Variant
    v = Array("abc", "xyz", "abd")
    Debug.Print "JF-var=["; Join(v, "|"); "]"
    Debug.Print "JF-var-def=["; Join(v); "]"

    Dim sa(1 To 3) As String
    sa(1) = "abc"
    sa(2) = "xyz"
    sa(3) = "abd"
    Debug.Print "JF-str=["; Join(sa, "|"); "]"

    Dim f As Variant
    f = Filter(sa, "ab")
    Debug.Print "JF-f-lb="; LBound(f); " ub="; UBound(f)
    Debug.Print "JF-f=["; Join(f, "|"); "]"

    Dim g As Variant
    g = Filter(sa, "qq")
    Debug.Print "JF-none-ub="; UBound(g)

    Dim h As Variant
    h = Filter(v, "ab", False)
    Debug.Print "JF-excl-ub="; UBound(h)
    Debug.Print "JF-excl=["; Join(h, "|"); "]"

    Dim n As Variant
    n = Array(1, 2, 3)
    On Error GoTo TypeMis
    Debug.Print "JF-num=["; Join(n, ","); "]"
    Debug.Print "JF-DONE"
    Exit Sub
TypeMis:
    Debug.Print "JF-err="; Err.Number
    Debug.Print "JF-DONE"
End Sub
