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
    Dim ne2 As Variant
    nv2 = Null
    ne2 = Empty
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

    ' 第六十四刀：Abs/Int/Fix/Hex/Oct 是“Null 进、Null 出”（真值 Null，不是 0，也不是 94）
    Err.Clear
    nc = Abs(nv2)
    Debug.Print "JF-nullabs="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = Int(nv2)
    Debug.Print "JF-nullint="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = Fix(nv2)
    Debug.Print "JF-nullfix="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = Hex(nv2)
    Debug.Print "JF-nullhex="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = Oct(nv2)
    Debug.Print "JF-nulloct="; IsNull(nc); "/"; Err.Number

    ' 非 Null 路未动（变体位只在实参是 Variant 槽时接手）
    Err.Clear
    nc = Hex(255)
    Debug.Print "JF-hex255="; CStr(nc); "/"; Err.Number

    ' 第六十五刀：比较符的 Null 传播 —— 任一侧是 Null ⇒ 结果 Null（不是 False/True）
    Err.Clear
    nc = (nv2 = "")
    Debug.Print "JF-cmpeq="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 <> "")
    Debug.Print "JF-cmpne="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 < 1)
    Debug.Print "JF-cmplt="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 = nv2)
    Debug.Print "JF-cmpeqeq="; IsNull(nc); "/"; Err.Number

    ' 反向钉：Empty 不是 Null，普通比较仍出布尔
    Err.Clear
    nc = (ne2 = "")
    Debug.Print "JF-cmpempty="; IsNull(nc); "/"; CStr(nc); "/"; Err.Number

    Err.Clear
    nc = ("a" = "a")
    Debug.Print "JF-cmpstr="; IsNull(nc); "/"; CStr(nc); "/"; Err.Number

    ' 第六十六刀：`&` 的双 Null 折叠（只有两侧同时是 Null 才是 Null）
    Err.Clear
    nc = nv2 & nv2
    Debug.Print "JF-nullcat="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = nv2 & "a"
    Debug.Print "JF-nullcatone="; IsNull(nc); "/"; CStr(nc); "/"; Err.Number

    ' CStr(<比较>) / CStr(<拼接>) —— Null 进 CStr 那一站是 94
    Err.Clear
    nc = CStr(nv2 & nv2)
    Debug.Print "JF-nullcatcstr="; Err.Number

    Err.Clear
    nc = CStr(nv2 = "")
    Debug.Print "JF-nullcmpcstr="; Err.Number

    ' 第六十七刀：Or/And 的 Null 真值表
    Err.Clear
    nc = (nv2 Or True)
    Debug.Print "JF-nullortrue="; IsNull(nc); "/"; CStr(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 Or False)
    Debug.Print "JF-nullorfalse="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 And False)
    Debug.Print "JF-nullandfalse="; IsNull(nc); "/"; CStr(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 And True)
    Debug.Print "JF-nullandtrue="; IsNull(nc); "/"; Err.Number

    ' 反向钉：非 Null 的逻辑位不受影响
    Err.Clear
    nc = (True Or False)
    Debug.Print "JF-cmplogic="; IsNull(nc); "/"; CStr(nc); "/"; Err.Number

    ' 数值侧：VB6 IDE 实测 Null Or 5 = Integer 5、Null And 0 = Integer 0
    Err.Clear
    nc = (nv2 Or 5)
    Debug.Print "JF-nullor5="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 And 0)
    Debug.Print "JF-nulland0="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 Or 0)
    Debug.Print "JF-nullor0="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 And 5)
    Debug.Print "JF-nulland5="; IsNull(nc); "/"; Err.Number

    ' 第六十八刀：Xor/Eqv 任一侧 Null 恒 Null；Imp = (Not x) Or y
    Err.Clear
    nc = (nv2 Xor 5)
    Debug.Print "JF-nullxor="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (True Xor nv2)
    Debug.Print "JF-nullxor2="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 Eqv 5)
    Debug.Print "JF-nulleqv="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (False Eqv nv2)
    Debug.Print "JF-nulleqv2="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 Imp 5)
    Debug.Print "JF-nullimp5="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 Imp True)
    Debug.Print "JF-nullimptrue="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 Imp False)
    Debug.Print "JF-nullimpfalse="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (5 Imp nv2)
    Debug.Print "JF-nullimpnot="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number

    Err.Clear
    nc = (False Imp nv2)
    Debug.Print "JF-notimptrue="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number

    Err.Clear
    nc = (True Imp nv2)
    Debug.Print "JF-notimpnull="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 Or -1)
    Debug.Print "JF-nullorneg="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number

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
    ' 第六十九刀：拼接链（VB6 IDE 实测 G1..G6）
    Err.Clear
    nc = (nv2 & nv2 & nv2)
    Debug.Print "JF-nullcat3="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 & nv2 & nv2 & nv2)
    Debug.Print "JF-nullcat4="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = ((nv2 & nv2) & (nv2 & nv2))
    Debug.Print "JF-nullcat4g="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = (nv2 & nv2 & 1)
    Debug.Print "JF-nullcat3b="; IsNull(nc); "/"; CStr(nc); "/"; TypeName(nc); "/"; Err.Number
    ' 第七十一刀: Len/Trim/LTrim/RTrim/UCase/LCase 的 Null → Null
    ' (真 VB6 IDE 读数 A1/B1/B4/B5/B6/B7: 六格都是 Null、Err 0)。
    ' 反向钉: 非 Null 的 Variant 与 Empty 仍走原档 (变体位不是"一律 Null")。
    Dim sv2 As Variant
    sv2 = "  ab  "
    Dim ev2 As Variant
    ev2 = Empty
    Dim iv As Long

    ' 这几格有会真报错的那一条 (数值目标收 Null ⇒ 94), 所以本段自带 Resume Next。
    On Error Resume Next

    Err.Clear
    nc = Len(nv2)
    Debug.Print "JF-lenv-null="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = Trim(nv2)
    Debug.Print "JF-trimv-null="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = LTrim(nv2)
    Debug.Print "JF-ltrimv-null="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = RTrim(nv2)
    Debug.Print "JF-rtrimv-null="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = UCase(nv2)
    Debug.Print "JF-ucasev-null="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = LCase(nv2)
    Debug.Print "JF-lcasev-null="; IsNull(nc); "/"; Err.Number

    Err.Clear
    nc = Len(nv2)
    Debug.Print "JF-lenv-tn="; TypeName(nc); "/"; Err.Number

    Err.Clear
    nc = Trim(sv2)
    Debug.Print "JF-trimv-str=["; nc; "]/"; Err.Number

    Err.Clear
    nc = UCase(sv2)
    Debug.Print "JF-ucasev-str=["; nc; "]/"; Err.Number

    Err.Clear
    nc = Len(sv2)
    Debug.Print "JF-lenv-str="; nc; "/"; Err.Number

    Err.Clear
    nc = Len(ev2)
    Debug.Print "JF-lenv-empty="; nc; "/"; Err.Number

    Err.Clear
    nc = Trim(ev2)
    Debug.Print "JF-trimv-empty=["; nc; "]/"; Err.Number

    ' 数值目标收 Null ⇒ 94 (与第六十三刀数值站同口径; 变体位交回的是 Null, 不是 0)
    Err.Clear
    iv = Len(nv2)
    Debug.Print "JF-lenv-long="; Err.Number
    On Error GoTo 0

    Debug.Print "JF-DONE"
End Sub
