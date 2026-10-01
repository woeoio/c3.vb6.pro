' 账 #122 夹具: VB6 在**赋值那一刻**做隐式 CStr —— `Dim s As String: s = 123` 合法。
'
' 修复前这一格把值原样交给 vb6_BSTR_Assign, 后果分两种, 都不响不等于没事:
'   整型那半 (Long/Integer/Boolean) = **静默**把数字当 BSTR 指针存进去, 之后第一次读/释放
'     就崩 (实测 BASE 二进制 `s = n` 编得过、跑起来 SIGSEGV, rc=139);
'   浮点/日期那半 (Double/Currency/Date) = cl `error C2440 无法从 double 转换为 BSTR`, 响的。
' Charts 2020 的 `sDiplay = CLng(me->m_Value)` 就是整型那一半的现存样本 (发的是
' `vb6_BSTR_AssignMove(&sDiplay, vb6_CLng(...))`, 本批修成外面套一层 vb6_CStrLong)。
'
' 读数一律不带方括号: 套件用 -like 比针, [ABC] 在那里是字符类 (门 #178 就是这么红的)。

Option Explicit

Dim gS As String
' Byte 那半边只测**模块级**的: 局部 `Dim bt As Byte` 在 inferExprType 里根本不可见
' (backend 只有 Bstr/Single/Date/Double/Bool/Long/LongPtr/Variant 八张 knownXxxVars_ 表,
'  没有 Byte 那张), 于是 `s = bt` 仍发裸值、跑起来 SIGSEGV —— 那是另一条根因, 另立一账,
' 不该混进这一格的判据里假装它已经好了。
Dim gb As Byte

Sub Main()
    Dim s As String
    Dim t As String
    Dim n As Long, d As Double, bo As Boolean
    Dim dt As Date
    Dim i As Long

    n = 42: d = 2.5: gb = 65: bo = (1 = 1)
    dt = DateSerial(2026, 1, 2)

    ' (1) 字面量 / 变量 / 表达式 / 内置标量函数
    s = 123
    Debug.Print "S122-lit="; s
    s = n
    Debug.Print "S122-long="; s
    s = d
    Debug.Print "S122-dbl="; s
    s = gb
    Debug.Print "S122-byte="; s
    s = CByte(65)
    Debug.Print "S122-cbyte="; s
    s = Asc("A")
    Debug.Print "S122-asc="; s
    s = bo
    Debug.Print "S122-bool="; s
    s = n + 1
    Debug.Print "S122-expr="; s
    s = Len("abcd")
    Debug.Print "S122-len="; s
    Dim a(1 To 3) As Long
    s = UBound(a)
    Debug.Print "S122-ubound="; s

    ' (2) 日期那半: 文本随区域设置变, 所以不比字面量, 只比**同一条折算路**的两个来源
    s = dt
    t = CStr(dt)
    If s = t Then
        Debug.Print "S122-date-ok=Y"
    Else
        Debug.Print "S122-date-ok=N"
    End If

    ' (3) 负向护栏: 值本来就是字符串的, 一条都不许被折算包住 (套上就是数字打出来或崩)
    gS = "ABCDEFG"
    s = Left(gS, 3)
    Debug.Print "S122-str-left="; s
    s = UCase("xy")
    Debug.Print "S122-str-ucase="; s
    s = MakeStr(4)
    Debug.Print "S122-str-func="; s
    s = MakeNum(7)
    Debug.Print "S122-func-num="; s

    ' (4) 拼接那条路不许跟着变 (它早就有自带的逐操作数折算)
    s = "n=" & n & " d=" & d
    Debug.Print "S122-concat="; s

    ' (5) 借用侧仍是深拷贝: 别名之后改源不许动副本
    s = "SHARED"
    t = s
    s = "CHANGED"
    Debug.Print "S122-borrow="; t

    ' (6) 反复赋值 + 大量标量赋值: 折出来的临时块必须被接手 (所有权判错当场堆损坏)
    For i = 1 To 20000
        s = n + i
    Next i
    Debug.Print "S122-loop="; s
    Debug.Print "S122-DONE"
End Sub

Function MakeStr(k As Long) As String
    Dim r As String
    r = String$(k, 77)
    MakeStr = r
End Function

Function MakeNum(k As Long) As Long
    MakeNum = k * 3
End Function
