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

    ' 第五十九刀：数值/布尔/货币/字节元素按 CStr 转（真引擎读数 .temp/join/oracle_join.vbs）。
    ' 旧写法把 `Join(Array(1,2,3), ",")` 钉成 "JF-err=13" —— 那是真机读数的**反面**
    ' （VB6 立即窗口和 cscript 都给 1,2,3）。
    Dim n As Variant
    n = Array(1, 2, 3)
    Debug.Print "JF-num=["; Join(n, ","); "]"

    Dim mm As Variant
    mm = Array(1.5, 2.25, True, Empty)
    Debug.Print "JF-mix=["; Join(mm, ","); "]"

    Dim a1 As Variant, a2 As Variant
    a1 = CByte(65): a2 = CByte(66)
    Debug.Print "JF-byte=["; Join(Array(a1, a2), "-"); "]"

    Dim cy1 As Variant, cy2 As Variant
    cy1 = CCur(1.2345): cy2 = 2
    Debug.Print "JF-cy=["; Join(Array(cy1, cy2), ","); "]"

    ' 错误面：Null 元素 → 13；非数组 → 13；二维数组 → 13
    On Error Resume Next
    Dim e1 As Variant
    Err.Clear
    e1 = Join(Array(Null, 1), ",")
    Debug.Print "JF-null-err="; Err.Number
    Err.Clear
    Dim sv As Variant
    sv = 5
    e1 = Join(sv, ",")
    Debug.Print "JF-nonarr-err="; Err.Number
    Err.Clear
    Dim two(1, 1) As Variant
    two(0, 0) = 1
    e1 = Join(two, ",")
    Debug.Print "JF-2d-err="; Err.Number
    On Error GoTo 0
    Debug.Print "JF-DONE"
End Sub
