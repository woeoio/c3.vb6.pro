Attribute VB_Name = "IntSuffixNeg"
Option Explicit
' 负例 (账 #195 的边界): 显式 `%` 就是 VB 的 Integer 档 (-32768..32767)，超出**必须报**，
' 不许按 int32 收下再让 int16 形参去截 —— 那是把值改错 (2147483648 会静默回绕成 -2147483648)。
Public Function TooBig() As Integer
    Dim y As Integer
    y = 2147483648%
    TooBig = y
End Function
