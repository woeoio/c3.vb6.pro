Option Explicit

' <vbeclipse> 账 #217 的负例: 这一刀只该让**造出来的**那个 Is 消失, 不该把
' 「未声明标识符」这条诊断一起关掉 —— 一个工程里 500 条噪声会把真打错的名字埋掉,
' 所以两头都要钉: `Case Is > 2` 不再报 'Is', 而真没声明过的名字照旧报。

Public Function Neg(n As Long) As String
    Dim real As Long
    real = n
    Select Case real
        Case Is > 2
            Neg = "big"
        Case Else
            ' nopeHereIsNotAName 从来没声明过 —— 它必须继续报 VB3001
            Neg = "rest" & nopeHereIsNotAName
    End Select
End Function
