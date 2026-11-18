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

    ' 第六十一刀：Filter 的元素读数（.temp/filt/oracle_filt.vbs 逐条）。
    ' 真引擎：命中按文本比、结果数组**逐元素保留子类型**；Null 元素不报错也不参与命中。
    ' 改前这四条分别是 13 / 13 / 13 / 撞崩（结果数组固定 BSTR 档）。
    Dim fn1 As Variant, fr As Variant
    fn1 = Array(1, 2, 21)
    fr = Filter(fn1, "1", True)
    Debug.Print "JF-fnum="; UBound(fr) - LBound(fr) + 1; "/"; VarType(fr(0)); "/"; Join(fr, "|")

    fn1 = Array(1, "x1", True)
    fr = Filter(fn1, "1", True)
    Debug.Print "JF-fmix="; UBound(fr) - LBound(fr) + 1; "/"; VarType(fr(0)); ","; VarType(fr(1)); "/"; Join(fr, "|")

    fn1 = Array(Empty, 1)
    fr = Filter(fn1, "", True)
    Debug.Print "JF-fempty="; UBound(fr) - LBound(fr) + 1; "/"; VarType(fr(0))

    On Error Resume Next
    Err.Clear
    fn1 = Array(Null, 1)
    fr = Filter(fn1, "1", True)
    Debug.Print "JF-fnull="; UBound(fr) - LBound(fr) + 1; "/"; Err.Number

    ' 第六十二刀：Null 转换站。VB6 口径：Variant(Null) 隐式转 String 一律 **94**
    ' （读数 .temp/null/oracle_null.vbs）。改前这一批全是 0（静默空串）。
    ' 注：`CStr(Null)` **不在本刀**（CStr 走自己的重载宏，改名的代价已实测太大，
    ' 门 44311607151 红过；已记入 ai/039 §11.36 遗留）。
    Err.Clear
    Dim nv2 As Variant, nc As Variant
    nv2 = Null
    nc = Asc(nv2)
    Debug.Print "JF-nullasc="; Err.Number

    Err.Clear
    nc = Join(Array(1, 2), nv2)
    Debug.Print "JF-nulldelim="; Err.Number

    Err.Clear
    nc = Split(nv2, ",")
    Debug.Print "JF-nullsplit="; Err.Number

    Err.Clear
    nc = Replace(nv2, "a", "b")
    Debug.Print "JF-nullrep="; Err.Number

    ' 第六十三刀：数值/位置站。VB6 口径 CInt/CLng/CDbl/Chr/Mid(s, Null) 一律 **94**
    ' （读数 .temp/null/oracle_num.vbs）。改前这一批静默给 0。
    Err.Clear
    nc = CInt(nv2)
    Debug.Print "JF-nullcint="; Err.Number

    Err.Clear
    nc = CLng(nv2)
    Debug.Print "JF-nullclng="; Err.Number

    Err.Clear
    nc = Chr(nv2)
    Debug.Print "JF-nullchr="; Err.Number

    Err.Clear
    nc = Mid("abc", nv2)
    Debug.Print "JF-nullmid="; Err.Number

    Err.Clear
    nc = CCur(nv2)
    Debug.Print "JF-nullccur="; Err.Number

    ' CCur 的非 Null 路不许被 V 后缀带跑（BSTR 档靠 vb6_NumVal 解析）
    Err.Clear
    nc = CCur(2.5)
    Debug.Print "JF-ccur25="; CStr(nc); "/"; Err.Number

    ' 反向钉：Empty 不是 Null（必须还是 0），算术传播仍是 Null
    Err.Clear
    nc = CInt(Empty)
    Debug.Print "JF-cintempty="; nc; "/"; Err.Number

    Err.Clear
    nc = 1 + nv2
    Debug.Print "JF-nullplus="; IsNull(nc); "/"; Err.Number

    ' 与上面相反的一侧（不能误伤）：`&` 走的是宽松口径、`IsNull` 照旧
    Err.Clear
    nc = "x" & "y"
    Debug.Print "JF-concat-ok="; nc; "/"; Err.Number

    Err.Clear
    Debug.Print "JF-isnull="; IsNull(nv2); "/"; Err.Number

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
