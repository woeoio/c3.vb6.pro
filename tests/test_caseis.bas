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

Public Sub Main()
    Debug.Print "CI-rel=" & Rel(3) & "/" & Rel(5) & "/" & Rel(6) & "/" & Rel(9) & "/" & Rel(7) & "/" & Rel(8)
    Debug.Print "CI-ne=" & Ne(4) & "/" & Ne(5)
    Debug.Print "CI-mixed=" & Mixed(1) & "/" & Mixed(3) & "/" & Mixed(2) & "/" & Mixed(0)
    Debug.Print "CI-word=" & Word("b") & "/" & Word("zz") & "/" & Word("a")
    Debug.Print "CI-half=" & Half(0.7) & "/" & Half(0.1) & "/" & Half(0.5)
    Debug.Print "CI-var=" & VarCase(9) & "/" & VarCase(2) & "/" & VarCase(CDbl(7))
    Debug.Print "CI-DONE"
End Sub
