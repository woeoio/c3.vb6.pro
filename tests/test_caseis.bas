Option Explicit

' <vbeclipse> 账 #217 第一刀: `Case Is > 2` 里的 Is 不是标识符。
' 改前 parser 为了借一次优先级解析, 在 AST 里造了一枚 IdentifierExpr("Is") 当左操作数,
' 于是语义层每条 `Case Is` 都按「未声明标识符」处理: 本模块写 Option Explicit 时报一条
' VB3001 (两份真工程合起来 36 条, 全仓 138 条), 没写时更实在 —— 每枚用到的过程发一枚
' 没人引用的 vb6_VARIANT 局部 (探针实测)。发码那侧从来只读比较符与右操作数, 值一直是对的。
' 现在比较符存进 CaseValue.relOp, value 只装右操作数, 那枚假名字不进 AST。
' 这里钉的是「形状换了、判定没换」: 六个关系符各自的分档、Is 与值混写的 Case、
' 字符串 Select、Single Select、Variant 测试表达式。

Private Function Rel(n As Long) As String
    Select Case n
        Case Is < 5
            Rel = "lt"
        Case Is <= 5
            Rel = "le"
        Case Is = 6
            Rel = "eq"
        Case Is > 8
            Rel = "gt"
        Case Is >= 7
            Rel = "ge"
        Case Else
            Rel = "rest"
    End Select
End Function

Private Function Ne(n As Long) As String
    Select Case n
        Case Is <> 4
            Ne = "ne4"
        Case Else
            Ne = "is4"
    End Select
End Function

' Case Is > 2, 1 —— 关系式与普通值混写在一个 Case 里 (优先级解析不能把 `, 1` 吃进右操作数)
Private Function Mixed(n As Long) As String
    Select Case n
        Case Is > 2, 1
            Mixed = "big"
        Case 2
            Mixed = "two"
        Case Else
            Mixed = "rest"
    End Select
End Function

Private Function Word(s As String) As String
    Select Case s
        Case Is = "b"
            Word = "beq"
        Case Is > "m"
            Word = "gtm"
        Case Else
            Word = "rest"
    End Select
End Function

' Fix 136 那一档的 Case 侧: 测试表达式是 Single, 临时变量与右操作数都按 float 走
Private Function Half(v As Single) As String
    Select Case v
        Case Is > 0.5
            Half = "hi"
        Case Is < 0.5
            Half = "lo"
        Case Else
            Half = "mid"
    End Select
End Function

' 测试表达式是 Variant: 序言按 VariantToLong 拆, Case 侧的 relOp 走的是同一条发射路
Private Function VarCase(v As Variant) As String
    Select Case v
        Case Is >= 7
            VarCase = "hi"
        Case Is < 7
            VarCase = "lo"
        Case Else
            VarCase = "rest"
    End Select
End Function

' 账 #305 第311刀: 主语静态答 Variant 而 Case 标签**全是**串档 —— 承载档必须跟标签走。
' 靶子 = ComCtlsDemo 工具栏 `Select Case Button.Tag` (TbrButton.cls:73 `Property Get Tag() As Variant`):
' 改前发成 `int32_t _vb6_select = vb6_VariantToLong(...)`, 标签却是 `vb6_BSTR_FromStr(L"...")` 的指针,
' C 模式只警告不报红 ⇒ 每一档都假 ⇒ 恒落 Case Else (点 "Show ListBoxW1" 只弹一句 "clicked button 5",
' ListView3 / ListBoxW1 一枚都不换)。
' ⚠ 主语刻意用**Variant 形参/变量**这一档 (发码认得出是槽 ⇒ 补 vb6_VariantToString)。另有一形
'   `Select Case <用户函数>`, 该函数声明 As Variant: 改前改后都编不过 (cl C2440 "vb6_VARIANT →
'   int32_t/BSTR", 实测 1 条), 因为 cExprIsVariant 与 knownVariantVars_ 都不认**用户函数的返回值**是槽
'   —— 独立缺陷、响亮、与本刀不同, 已记为开口，不在这里钉。
Private Function ByTag(v As Variant) As String
    Select Case v
        Case "a"
            ByTag = "aaa"
        Case "b", "bb"
            ByTag = "bbb"
        Case Else
            ByTag = "else"
    End Select
End Function

' 账 #305 第319刀: 混合标签（串 + 数）—— 311 的两条整档出口都不成立，只有逐项 VarCmp 才对。
' 改前发成 `int32_t _vb6_select = vb6_VariantToLong(v)` 再拿它跟 `vb6_BSTR_FromStr(L"s")` 的指针比，
' 每一档都假 ⇒ 恒落 Case Else（本夹具第一格 CI-mixtag 的旧期望 else/num 就是这毛病，VB6 答 str）。
' 现在承载档是整枚槽、每项按该标签自己那一档走 vb6_VarCmp*（与 `If a = b` 同一份比较出口）。
Private Function MixedTag(v As Variant) As String
    Select Case v
        Case "s"
            MixedTag = "str"
        Case 7
            MixedTag = "num"
        Case Else
            MixedTag = "else"
    End Select
End Function

' 同一混合档的另两种条件形状：范围（VarCmpGe && VarCmpLe）与 Case Is（VarCmp<Op>）。
' 串标签取可数值化的 "5"/"9"：串数跨档比较落的是数值档（本机 VB-family 运行期实测
' Select Case "7" 能命中 Case 7）；不可数值化的串对数值会报 13，那是账 #416 那一族
' 在比较出口上的既有缺口，不在这枚夹具里钉。
Private Function MixRange(v As Variant) As String
    Select Case v
        Case "5" To "9"
            MixRange = "srng"
        Case 20 To 30
            MixRange = "nrng"
        Case Else
            MixRange = "else"
    End Select
End Function

Private Function MixIs(v As Variant) As String
    Select Case v
        Case Is > 100
            MixIs = "gt100"
        Case Is >= "5"
            MixIs = "ge5"
        Case Else
            MixIs = "lo"
    End Select
End Function

Public Sub Main()
    Debug.Print "CI-rel=" & Rel(3) & "/" & Rel(5) & "/" & Rel(6) & "/" & Rel(9) & "/" & Rel(7) & "/" & Rel(8)
    Debug.Print "CI-ne=" & Ne(4) & "/" & Ne(5)
    Debug.Print "CI-mixed=" & Mixed(1) & "/" & Mixed(3) & "/" & Mixed(2) & "/" & Mixed(0)
    Debug.Print "CI-word=" & Word("b") & "/" & Word("zz") & "/" & Word("a")
    Debug.Print "CI-half=" & Half(0.7) & "/" & Half(0.1) & "/" & Half(0.5)
    Debug.Print "CI-var=" & VarCase(9) & "/" & VarCase(2) & "/" & VarCase(CDbl(7))
    Debug.Print "CI-vartag=" & ByTag("b") & "/" & ByTag("bb") & "/" & ByTag("a") & "/" & ByTag("z")
    ' 混合标签那一形（第319刀已修）: 串标签命中串、数值标签命中数；改前第一格恒答 else。
    Debug.Print "CI-mixtag=" & MixedTag("s") & "/" & MixedTag(7)
    Debug.Print "CI-mixrng=" & MixRange(7) & "/" & MixRange(25) & "/" & MixRange(1)
    Debug.Print "CI-mixis=" & MixIs(200) & "/" & MixIs(50) & "/" & MixIs(1)
    Debug.Print "CI-DONE"
End Sub
