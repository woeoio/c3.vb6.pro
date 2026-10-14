' <vbeclipse>: Variant **装箱表本身**的哨兵夹具。
' 为什么单独有这一条: 装箱档位以前有两份并行表 (wrapVariantValue 的 switch 与
' boxToVariant), 账 #123 只在其中一份修了 Byte、Fix 198 只在另一份修了 Boolean,
' 于是"某一档对不对"取决于表达式走哪条路 —— DT43 (dt1.CheckBox 的 VarType) 就是在
' 一个和装箱毫无关系的控件应用上偶然暴露的。这里把三条真实路线各钉一遍:
'   (1) 内置函数实参装箱   VarType(m_b)        —— DT43 断的那条
'   (2) Variant 赋值装箱   v = m_b             —— wrapVariantValue 那条
'   (3) 过程形参装箱       ReportV ..., m_b    —— packLetValueArg 那条
' 以及模块级/局部/函数返回三种来源 (账 #123 的原始形态就是"模块级 Byte 错、局部对")。
' Date/Single 两档本批已修进权威表也进了 needle (修前两者都读成 5); Currency 仍只打印
' 不钉 —— 它与 VB6 的 VT_CY=6 的差距没有真 VB6 读数可依据, 已在汇报里记成待办。
Option Explicit

Private m_b As Boolean
Private m_by As Byte
Private m_i As Integer
Private m_l As Long
Private m_s As String
Private m_sin As Single
Private m_dbl As Double
Private m_d As Date

' Acct #305 para-B161 (knife 49): MODULE-level arrays handed to a ByRef Variant
' formal. For these inferExprType answers the ELEMENT type, so the switch at the
' call site took a scalar arm and put the descriptor POINTER into that scalar slot
' -- (int32_t)(b161Ml) into .lVal, b161Ms into .bstrVal with no cast at all. Only
' arrays Dim-ed inside the procedure fell through to the arm knife 48 fixed.
Private b161Ml() As Long
Private b161MB() As Byte
Private b161Ms() As String
Private b161Mf(1 To 3) As Long

Private Sub ReportV(ByVal tag As String, ByVal v As Variant)
    Debug.Print tag; VarType(v); "/"; TypeName(v)
End Sub

Private Function RetB() As Boolean
    RetB = True
End Function

Public Sub Main()
    Dim lb As Boolean
    Dim lby As Byte
    lb = True
    lby = 67
    m_b = True
    m_by = 67
    m_i = 7
    m_l = 9
    m_s = "x"
    m_sin = 1.5
    m_dbl = 2.5
    ' Date 用序列号给值: CDate("2026-01-01") 在本仓现在是 0 (日期串解析是另一格问题,
    ' 已单独记), 那会让本夹具的 clng/cdbl 两行读成 0, 分不清是装箱还是解析。
    m_d = 46023

    Debug.Print "VB-mod-bool="; VarType(m_b); "/"; TypeName(m_b)
    Debug.Print "VB-mod-byte="; VarType(m_by); "/"; TypeName(m_by)
    Debug.Print "VB-loc-bool="; VarType(lb); "/"; TypeName(lb)
    Debug.Print "VB-loc-byte="; VarType(lby); "/"; TypeName(lby)
    Debug.Print "VB-func-bool="; VarType(RetB()); "/"; TypeName(RetB())
    Debug.Print "VB-expr-bool="; VarType(lb And True)
    Debug.Print "VB-mod-long="; VarType(m_l)
    Debug.Print "VB-mod-int="; VarType(m_i)
    Debug.Print "VB-mod-str="; VarType(m_s)
    Debug.Print "VB-mod-sin="; VarType(m_sin)
    Debug.Print "VB-mod-dbl="; VarType(m_dbl)
    Debug.Print "VB-mod-date="; VarType(m_d)

    Dim v As Variant
    v = m_b
    Debug.Print "VB-asg-bool="; VarType(v); "/"; TypeName(v)
    v = m_by
    Debug.Print "VB-asg-byte="; VarType(v); "/"; TypeName(v)
    v = lb
    Debug.Print "VB-asg-loc-bool="; VarType(v)
    v = m_l
    Debug.Print "VB-asg-long="; VarType(v)
    v = m_d
    Debug.Print "VB-asg-date="; VarType(v)

    ReportV "VB-call-bool=", m_b
    ReportV "VB-call-byte=", m_by
    ReportV "VB-call-long=", m_l
    ReportV "VB-call-date=", m_d

    ' 新档位的**反向**也要钉: 装成 VT_DATE / VT_R4 之后, 数值提取端与 IsDate 还得答对
    ' (RTL 里 vb6_vtDate 早有 TypeName/Format 档, 缺的是 VariantToLong/ToDouble 那一侧)。
    Dim vd As Variant
    vd = m_d
    Dim vs As Variant
    vs = m_sin
    Debug.Print "VB-date-clng="; CLng(vd)
    Debug.Print "VB-date-cdbl="; CDbl(vd)
    Debug.Print "VB-date-isdate="; IsDate(vd)
    Debug.Print "VB-date-cdate="; (CDate(vd) = m_d)
    Debug.Print "VB-sin-cdbl="; CDbl(vs)
    Debug.Print "VB-sin-clng="; CLng(vs)
    ' 账 #305 (§B138 尾巴): §B138 把 VarPtr/StrPtr/ObjPtr 的类型答案换成 LongPtr 之后, 同一枚地址
    ' 在 x64 装成 VT_I8(20)、在 x86 装成 VT_I4(3)。VarType 因此**不能钉**(它本来就该随位数变);
    ' 钉的是那四问在两台机器上必须给同一个答案 —— 改之前 x64 读出来是 Variant / False。
    ' 刻意不钉 CLng(vp): x64 那个地址放不下 Long, 报 6 (Overflow) 才是 VB6 的正确答案。
    Dim vp As Variant
    Dim vs2 As Variant
    Dim hx As Long
    hx = 7
    vp = VarPtr(hx)
    vs2 = StrPtr("box")
    Debug.Print "VB-ptr-tn="; TypeName(vp)
    Debug.Print "VB-ptr-isnum="; IsNumeric(vp)
    Debug.Print "VB-ptr-cbool="; CBool(vp)
    Debug.Print "VB-ptr-selfeq="; (vp = vp)
    Debug.Print "VB-str-tn="; TypeName(vs2)
    ' 账 #305 (§B141 + §B142): CDec 造出来的那一档 (VT_DECIMAL=14) 在「往外读」这一侧的判据。
    ' §B141 补的是两张表的格子(TypeName / IsNumeric), 第三十四刀补的是**换算本体与写侧的一枚标签**,
    ' 于是 CStr / CLng / CDbl 一起对上。两处根因都是「格式」不是「数值」:
    ' ① DECIMAL.wReserved 其实是 oleaut32 自己的 vt 标签(它写 14), 手填字段那两路以前留 0 ——
    '    数据全对而 OLE 不认, VariantChangeType 返回 S_OK 却交 0 / 交空串;
    ' ② x86 的 VARIANT 装 VT_DECIMAL 不走 pdecVal 指针, 是整枚 16 字节 DECIMAL 内联盖住 VARIANT
    '    —— 只填指针时 oleaut32 把那枚地址当数值答出来(此前那个「堆地址」读数就是这一条)。
    ' Format 那条读数照打但不进针面: 带小数点的格式串随区域设置变(与 AR-fmt2 同口径)。
    ' 刻意不钉 CDec 的算术/比较面: 那要另一族出口 (VarDec* 一族), 与本格的「读出来」无关。
    Dim vd2 As Variant
    Dim vd3 As Variant
    vd2 = CDec(5)
    vd3 = CDec("1.25")
    Debug.Print "VB-dec-tn="; TypeName(vd2)
    Debug.Print "VB-dec-cstr=[" & CStr(vd2) & "]"
    Debug.Print "VB-dec2-cstr=[" & CStr(vd3) & "]"
    Debug.Print "VB-dec-isnum="; IsNumeric(vd2)
    Debug.Print "VB-dec-fmt=[" & Format(vd2, "0.00") & "]"
    Debug.Print "VB-dec-clng="; CLng(vd2)
    Debug.Print "VB-dec-cdbl="; CDbl(vd3)
    ' Acct #305 para-B142 cell 2 (the float arm's tag). Comments here stay ASCII: C3 reads
    ' sources as ANSI/GBK, so UTF-8 Chinese lines round-trip as mojibake in the product.
    ' CDec(2.5) takes the VarDecFromR8 arm, and that DECIMAL comes back with wReserved = 0
    ' (only VarDecFromStr writes 14). Knife 34 tagged just the two hand-filled arms, so on
    ' x64 CStr(CDec(2.5)) printed empty and CDbl gave 0, while x86 was saved by the inline
    ' overlay on the read side -- an arch asymmetry. Deliberately NOT pinning
    ' CLng(CDec(2.5)): oleaut32 rounds 2.5 to 2 (banker's, = VB6), we currently answer 3
    ' (it goes through the double converter), and that half-way caliber is still unowned.
    Dim vd4 As Variant
    vd4 = CDec(2.5)
    Debug.Print "VB-decf-tn="; TypeName(vd4)
    Debug.Print "VB-decf-cstr=[" & CStr(vd4) & "]"
    Debug.Print "VB-decf-cdbl="; CDbl(vd4)
    ' Acct #305 para-B143 cell B1: the numeric ports a Decimal was read through. The
    ' BASE build answers 0/False for every one of them on BOTH arches (probe
    ' .build/b1069_reach/probe4.bas): assign to Long / LongPtr / Boolean, a Boolean
    ' If, arithmetic, comparison and a numeric formal arg. Integer Decimals only on
    ' purpose -- 2.5/3.5 ties belong to the unowned para-B77-1 rounding caliber, so
    ' nothing here pins a tie answer. Arith folds through vb6_VariantToLong and
    ' compare through vb6_VarToDouble_internal, which is why the four ports ship as
    ' one batch: a table that was patched while its neighbour was not is the defect.
    Dim vi As Long
    Dim vl As LongPtr
    Dim vb As Boolean
    Dim vsum As Variant
    vi = vd2
    Debug.Print "VB-dec-i4="; vi
    vl = vd2
    Debug.Print "VB-dec-lp="; vl
    vb = vd2
    Debug.Print "VB-dec-bb="; vb
    If vd2 Then
        Debug.Print "VB-dec-if=True"
    Else
        Debug.Print "VB-dec-if=False"
    End If
    vsum = vd2 + 1
    Debug.Print "VB-dec-add="; vsum
    Debug.Print "VB-dec-eq="; (vd2 = 5)
    ' Acct #305 para-B145: exact arithmetic. Codegen used to unwrap a Variant operand
    ' through vb6_VariantToLong for + - * -- it picked the numeric arm at compile time --
    ' so the BASE readings were the same on BOTH arches (probe .build/b1095_arith):
    '   Double box:  1.5+1 -> 3,  1.5*2 -> 4,  1.5-1 -> 1
    '   Decimal box: 1.25+1 -> 2, 1.25*2 -> 2, 0.1+0.2 -> 0, TypeName -> Long
    ' The four operators now go through one runtime exit in the RTL, which boxes both
    ' operands and hands them to oleaut32's VarAdd/VarSub/VarMul/VarDiv -- the engine
    ' VB6's own Variant arithmetic is built on -- so the arm comes from the value's vt
    ' at run time. Nothing here pins a tie answer (2.5/3.5 belong to the unowned
    ' para-B77-1 rounding caliber), and 1.5+1.5=3 is used where a Long target is asked.
    Dim ar1 As Variant
    Dim ar2 As Variant
    Dim ar3 As Variant
    Dim arx As Variant
    Dim arl As Long
    ar1 = 1.5
    Debug.Print "VB-ar-dbl-add="; ar1 + 1
    Debug.Print "VB-ar-dbl-mul="; ar1 * 2
    Debug.Print "VB-ar-dbl-sub="; ar1 - 1
    ar2 = CDec("1.25")
    Debug.Print "VB-ar-dec-add="; CStr(ar2 + 1)
    Debug.Print "VB-ar-dec-mul="; CStr(ar2 * 2)
    Debug.Print "VB-ar-dec-tn="; TypeName(ar2 + 1)
    ar3 = CDec("0.1")
    ar2 = CDec("0.2")
    Debug.Print "VB-ar-exact="; CStr(ar3 + ar2)
    Debug.Print "VB-ar-exact-mul="; CStr(ar3 * ar2)
    arx = CDec("5.75")
    arl = arx + 1
    Debug.Print "VB-ar-target-long="; arl
    Debug.Print "VB-ar-if="; CBool(ar3 + ar2)
    ' Acct #305 para-B147: the narrowing check. "/" let the type oracle answer
    ' "Variant" from one predicate while codegen decided the routing from a
    ' second one (the C-text prefix list), so for operands the list cannot read
    ' (a Variant array's element, a host pseudo-member like TextHeight) the
    ' emitted C stayed a plain double while consumers believed it was a slot --
    ' they skipped BOTH the rounding step and the range check. BASE answer for
    ' the two lines below was 4 (3.5 -> round -> 4); the regressed build gave 3
    ' (C truncation at the vb6_ChkLong call). One shared predicate now says
    ' "is this a slot": routing = AST-union-text, oracle = the AST arm only, so
    ' every oracle-Variant result really is routed. 3.5 is not a tie, so this
    ' does not touch the unowned para-B77-1 rounding caliber.
    Dim arv As Variant
    Dim arb As Byte
    arv = 7
    arl = arv / 2
    Debug.Print "VB-ar-div-target-long="; arl
    arb = arv / 2
    Debug.Print "VB-ar-div-target-byte="; arb
    ' para-B147 second arm: an operand the type oracle cannot read at all -- a
    ' fixed-size Byte array element. inferExprType defaults such unknowns to
    ' Variant (that is what made 67f03217 answer Variant for the Charts lines
    ' while codegen still emitted a plain double), so the strict predicate must
    ' answer "not a slot" here and the Double arm must keep its rounding exit.
    ' The regressed build emitted vb6_ChkByte(<double>) = C truncation = 3.
    Dim cad(2) As Byte
    Dim cab As Byte
    cad(1) = 7
    cab = cad(1) / 2
    Debug.Print "VB-ar-div-target-byte-fx="; cab
    ' Acct #305 para-B146 (knife 39): the builtin exits whose C return IS a whole slot
    ' were never reconciled with the RTL declarations, so four names (CDec / CVErr /
    ' GetAllSettings / IIf) and the five boxing constructors (Byte/Integer/Single/Date/
    ' Identity) were missing from the codegen text predicate. Read through a Variant
    ' *variable* the AST arm covered them; called *directly* the consumer emitted a bare
    ' scalar: `CDec(1234567) + 1` -> (int32_t)(vb6_CDec(...) + 1) = C2088, `d = CDec("2.5")`
    ' = C2440. The negative control for the eight lines below is therefore "BASE cannot
    ' compile this file at all" (11 errors on both arches, probe .build/b1216_probe), not
    ' "BASE answers a wrong number". Readings are from the NEW build; both arches agree.
    Dim b6v As Variant
    Dim b6d As Double
    Debug.Print "VB-b146-dec-add="; CDec(1234567) + 1
    Debug.Print "VB-b146-dec-vt="; VarType(CDec(5))
    Debug.Print "VB-b146-dec-tn="; TypeName(CDec("1.25") * 2)
    b6d = CDec("2.5")
    Debug.Print "VB-b146-dec-dbl="; b6d
    b6v = IIf(True, CDec(1), CDec(2)) + 1
    Debug.Print "VB-b146-iif="; CStr(b6v)
    b6v = CVErr(11)
    Debug.Print "VB-b146-err-vt="; VarType(b6v)
    Debug.Print "VB-b146-err-iserr="; IsError(b6v)
    ' Fifth exit (GetAllSettings): a shape witness that turned out to be load-bearing.
    ' Since para-B149 cells 1+2 shipped (knife 40) our RTL answers like VB6: an app/section
    ' with no entries is Null -- VarType 1, IsNull True. Before that knife these two lines
    ' read vt=12 / isnull=False, and gate #538 turned them red the moment the product started
    ' answering correctly, which is the proof the needles are pinned to a real reading.
    ' The two-entry round trip (the 8204 two-dimensional shape) lives in its own fixture,
    ' tests/test_registry_settings.bas (para-B149/B150).
    Debug.Print "VB-b146-reg-vt="; VarType(GetAllSettings("C3NoSuchApp146", "Sec"))
    Debug.Print "VB-b146-reg-isnull="; IsNull(GetAllSettings("C3NoSuchApp146", "Sec"))
    ' Acct #305 para-B160 (knife 48): an array variable handed to a ByRef Variant
    ' formal. Before this knife codegen emitted &arr -- the ADDRESS OF THE DESCRIPTOR
    ' POINTER -- while the callee parameter is a vb6_VARIANT*, so the callee read a
    ' slot out of 4/8 bytes of pointer. BASE readings (probe .build/k47probe/p47proj,
    ' x64 real run) were VarType 40320 / TypeName "Variant"; cl only warned about the
    ' mismatched pointer type, BUILD-RC=0 and nothing was printed. VB6 answers the
    ' element flavor (8195 / "Long()"). The boxing now goes through the same carrier
    ' exit as the assignment v = a (8209 / "Byte()"), so the label comes from the
    ' descriptor itself. Deliberately read-only: ByRef write-back through a Variant
    ' formal is a separate caliber (para-B160).
    Dim b160ba() As Byte
    Dim b160lg() As Long
    Dim b160fx(1 To 3) As Long
    Dim b160nd() As Long
    ReDim b160ba(0 To 2)
    ReDim b160lg(1 To 4)
    ReDim b160nd(1 To 2, 1 To 3)
    b160ba(0) = 65
    b160lg(2) = 300
    b160fx(2) = 77
    b160nd(2, 3) = 512
    Call B160Take("byte", b160ba, 0)
    Call B160Take("long", b160lg, 2)
    Call B160Take("fixed", b160fx, 2)
    Call B160Take("nd", b160nd, -1)
    ReDim b161Ml(0 To 3)
    ReDim b161MB(0 To 2)
    ReDim b161Ms(0 To 1)
    b161Ml(1) = 210
    b161MB(0) = 66
    b161Ms(0) = "zz"
    b161Mf(3) = 88
    Call B160Take("m-long", b161Ml, -1)
    Call B160Take("m-byte", b161MB, -1)
    Call B160Take("m-str", b161Ms, -1)
    Call B160Take("m-fix", b161Mf, -1)
    Call B160Take("m-elem", b161Ml, 1)
    Debug.Print "VB-DONE"
End Sub

' Acct #305 para-B160: a ByRef Variant formal. The callee must see the array flavor,
' not the address of the descriptor pointer. ix below zero means label only -- a two-
' dimensional box has no single-index element to read.
Private Sub B160Take(ByVal tag As String, ByRef v As Variant, ByVal ix As Long)
    Debug.Print "VB-b160-" & tag & "="; VarType(v); "/"; TypeName(v)
    If ix >= 0 Then
        Debug.Print "VB-b160-" & tag & "-evt="; VarType(v(ix))
        Debug.Print "VB-b160-" & tag & "-val="; CLng(v(ix))
    End If
End Sub