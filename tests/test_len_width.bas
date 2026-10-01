' 账 #115：Len() 的"存储宽度"兜底桶把**模块级 String** 也吞了 —— `Len(gS)` 一度发成
' (int32_t)sizeof(gS) ⇒ x64 读出 8、x86 读出 4，而 VB6 要的是字符数。根因: knownBstrVars_
' 那张表在每过程入口 clear，只有局部声明/形参进表，模块级变量压根不在里面。
' 本用例把三条分界一起钉住，防止"修一边、坏另一边"：
'   · 字面量 / 局部 String / String 形参 —— 本来就对，别改坏；
'   · 模块级 String 与 Len(vbCrLf) —— 本批修的这两格，走 vb6_Len；
'   · 数值型变量 —— 继续走 sizeof（Task #44 的 SSTabEx/ChooseColor 堆越界就是它救回来的、
'     LEN-BYTE 那条 0xC0000005 崩溃的反证也在这里）。
Option Explicit

Dim gS As String
Dim gE As String         ' 刻意留空: sizeof 在这里读成指针宽度, vb6_Len 读成 0
Dim gL As Long
Dim gI As Integer
Dim gB As Byte

Sub Main()
    Dim lS As String
    gS = "abcde"
    lS = "abcde"
    Debug.Print "LW1=" & Len("你好hello")      ' 7  字面量 = 字符数
    Debug.Print "LW2=" & Len(gS)               ' 5  模块级 String（本批修的那条）
    Debug.Print "LW3=" & Len(lS)               ' 5  局部 String
    Debug.Print "LW4=" & Len(vbCrLf)           ' 2  字符串常量
    Debug.Print "LW5=" & LenP(gS)              ' 5  ByVal String 形参
    Debug.Print "LW6=" & LenR(gS)              ' 5  ByRef String 形参
    Debug.Print "LW7=" & Len(gL)               ' 4  模块级 Long = 存储宽度
    Debug.Print "LW8=" & Len(gI)               ' 2  模块级 Integer = 存储宽度
    Debug.Print "LW9=" & Len(gB)               ' 1  模块级 Byte = 存储宽度
    Debug.Print "LW10=" & Len(gS & gS)         ' 10 拼接表达式走通用路
    Debug.Print "LW11=" & Len(gE)              ' 0  空串
    Debug.Print "LW12=" & Len(Five())          ' 5  返回 String 的函数
End Sub

Function LenP(ByVal s As String) As Long
    LenP = Len(s)
End Function

Function LenR(ByRef s As String) As Long
    LenR = Len(s)
End Function

Function Five() As String
    Five = "abcde"
End Function
