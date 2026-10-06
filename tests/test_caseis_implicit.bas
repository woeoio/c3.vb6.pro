' <vbeclipse> 账 #217 的第二面: 这里**故意不写** Option Explicit。
' 改前每条 `Case Is` 会让语义层把那个造出来的 Is 登记成隐式 Variant 局部, 发码就在
' 用到它的每枚过程序言发一枚 `vb6_VARIANT Is = vb6_VariantEmpty();` —— 没人引用,
' 但每形一枚 (探针实测 2 枚)。现在 AST 里压根没有这枚标识符, 序言该是干净的。

Public Function ImpBig(n As Long) As String
    Select Case n
        Case Is > 2
            ImpBig = "big"
        Case Else
            ImpBig = "rest"
    End Select
End Function

Public Function ImpWord(s As String) As String
    Select Case s
        Case Is = "a"
            ImpWord = "aa"
        Case Else
            ImpWord = "rest"
    End Select
End Function
