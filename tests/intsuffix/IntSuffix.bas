Attribute VB_Name = "IntSuffix"
Option Explicit
' 账 #195: VB 的 Integer 类型后缀 `%` 以前在**词法层**就被拒 —— case '%' 那一支只把字符吃进
' rawText、不置任何标志，于是 `3%` 落回「无后缀十进制按数值大小定档」那一段，parseIntLit
' 看见残留的 '%' 就报「十进制数字超出 64 位整数表示范围」，一条合法语句级联出 9-13 条诊断。
' 这一支把标志补上 (十进制) 并在三种进制扫描器里认下这个后缀，另外 radixDigits 的剥离表
' 也补了 '%' —— 同一条件四处各抄一份，漏的就是这次那一份。
Public Function SixtyThree() As Integer
    Dim y As Integer
    y = 3%
    SixtyThree = y
End Function
Public Function FromHex() As Integer
    Dim z As Integer
    z = &HFF%
    FromHex = z
End Function
Public Function FromOct() As Integer
    Dim z As Integer
    z = &O17%
    FromOct = z
End Function
Public Function FromBin() As Integer
    Dim z As Integer
    z = &B101%
    FromBin = z
End Function
Public Function Minus() As Integer
    Dim z As Integer
    z = -4%
    Minus = z
End Function
Public Function InExpr() As Integer
    Dim z As Integer
    z = 3% + 1
    InExpr = z
End Function
Public Sub Take(Optional ByVal a As Integer = 3%)
    Debug.Print a
End Sub
' 证人 (排除"整条后缀都不支持"那种假红): & / ^ / 无后缀三形本来就通, 改完必须照通。
Public Function LongForm() As Long
    Dim y As Long
    y = 12&
    LongForm = y
End Function
Public Function PtrForm() As Long
    Dim y As Long
    y = &H10^
    PtrForm = y
End Function
Public Function PlainForm() As Integer
    Dim y As Integer
    y = 3
    PlainForm = y
End Function
