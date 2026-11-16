Attribute VB_Name = "ReMain"
Option Explicit

' C30-10 正则引擎的 C3 侧用例 —— 与 ReCases.vbs **逐行复刻**，输出写文件后与 cscript 的输出 diff。
' 注意：这里刻意继续用 CreateObject("VBScript.RegExp")，走的是 RTL 里的内置实现
' （vb6com_regexp.c，vb6_CreateObject 抢先拦截），不加载 vbscript.dll。
' 要跑"真组件"对照组：set C3_REGEXP_PREFER_MS=1 再运行。

Dim gRe As Object
Dim gF As Integer
Dim gF2 As Integer
Dim gIC As Boolean

Sub Emit(s As String)
    Print #gF, s
End Sub

Function IdxList(mc As Object) As String
    Dim k As Long, s As String
    s = ""
    For k = 0 To mc.Count - 1
        If k > 0 Then s = s & ","
        s = s & mc(k).FirstIndex & ":" & mc(k).Length
    Next
    IdxList = s
End Function

Sub RunCase(p As String, gl As Boolean, ic As Boolean, ml As Boolean, subj As String, mode As String, rep As String)
    Dim out As String, mc As Object, m As Object
    out = ""
    On Error Resume Next
    gRe.Pattern = p
    gRe.Global = gl
    gRe.IgnoreCase = ic
    gRe.Multiline = ml
    Err.Clear
    If mode = "test" Then
        If gRe.Test(subj) Then out = "t=1" Else out = "t=0"
    ElseIf mode = "exec" Then
        Set mc = gRe.Execute(subj)
        If Err.Number = 0 Then out = "n=" & mc.Count & " idx=" & IdxList(mc)
    ElseIf mode = "repl" Then
        out = "r=" & gRe.Replace(subj, rep)
    ElseIf mode = "sub" Then
        Set mc = gRe.Execute(subj)
        If Err.Number = 0 Then
            If mc.Count > 0 Then
                Set m = mc(0)
                out = "v=" & m.Value & " subs=" & m.SubMatches.Count & " s0=[" & m.SubMatches(0) & "] s1=[" & m.SubMatches(1) & "]"
            Else
                out = "v= n=0"
            End If
        End If
    ElseIf mode = "pat" Then
        out = "pat=" & gRe.Pattern
    End If
    If Err.Number <> 0 Then out = "err=" & Err.Number
    Err.Clear
    On Error GoTo 0
    Emit p & " | " & mode & " | " & out
End Sub

' ----------------------------------------------------------------
' 扩展语法用例（默认开启，不设开关）：判据不再来自 VBScript.RegExp，
' 而是 Python re（见 ext_oracle.py）；输出格式 "id|mode|结果"，与 oracle 逐行比。
' ----------------------------------------------------------------
Sub ExtCase(id As String, p As String, subj As String, mode As String)
    Dim out As String, mc As Object, m As Object, k As Long
    On Error Resume Next
    Err.Clear
    gRe.Pattern = p
    gRe.Global = False
    gRe.IgnoreCase = gIC
    gRe.Multiline = False
    If mode = "t" Then
        If gRe.Test(subj) Then out = id & "|t|1" Else out = id & "|t|0"
    Else
        Set mc = gRe.Execute(subj)
        If Err.Number = 0 Then
            If mc.Count = 0 Then
                out = id & "|m|-"
            Else
                Set m = mc(0)
                out = id & "|m|" & m.FirstIndex & ":" & m.Length
                For k = 0 To m.SubMatches.Count - 1
                    out = out & "|[" & m.SubMatches(k) & "]"
                Next
            End If
        End If
    End If
    If Err.Number <> 0 Then out = id & "|err|" & Err.Number
    Err.Clear
    On Error GoTo 0
    Print #gF2, out
End Sub

Sub ExtCases()
    gF2 = FreeFile
    Open App.Path & "\re_c3_ext.txt" For Output As #gF2
    gIC = False
    ' ---- 命名组 / 命名反向引用 ----
    ExtCase "e01", "(?<c>a)\k<c>", "aa", "t"
    ExtCase "e02", "(?P<c>a)(?P=c)", "aa", "t"
    ExtCase "e03", "(?'x'b)\k'x'", "bb", "t"
    ExtCase "e04", "(?<n>\d+)-(\w+)", "12-ab", "m"
    ExtCase "e05", "(?<x>a)", "a", "t"
    ExtCase "e06", "\k<x>(?<x>a)", "a", "t"      ' 前向引用：两边都报错
    ' ---- 后顾（定长） ----
    ExtCase "e07", "(?<=foo)bar", "foobar", "m"
    ExtCase "e08", "(?<=\d{2})x", "a12x", "m"
    ExtCase "e09", "(?<!foo)bar", "foobar", "t"
    ExtCase "e10", "(?<!foo)bar", "xbar", "t"
    ExtCase "e11", "(?<=a*)b", "ab", "t"        ' 变长后顾：5017
    ' ---- 原子组 / 占有量词 ----
    ExtCase "e12", "(?>a+)ab", "aaab", "t"
    ExtCase "e13", "(?>a+)c", "aaac", "t"
    ExtCase "e14", "a++ab", "aaab", "t"
    ExtCase "e15", "a*+a", "aaa", "t"
    ExtCase "e16", "\d++", "123", "t"
    ' ---- \A \z \Z ----
    ExtCase "e17", "\Aabc", "abc", "t"
    ExtCase "e18", "\Aabc", "xabc", "t"
    ExtCase "e19", "abc\z", "abc", "t"
    ExtCase "e20", "abc\z", "abc" & vbLf, "t"
    ExtCase "e21", "abc\Z", "abc" & vbLf, "t"
    ExtCase "e22", "abc\Z", "abc", "t"
    ' ---- \Q...\E ----
    ExtCase "e23", "\Qa.b\E", "a.b", "t"
    ExtCase "e24", "\Qa.b\E", "axb", "t"
    ExtCase "e25", "\Qa+b\E", "a+b", "t"
    ' ---- 内联选项 (?i:...) (?-i:...) (?s:...) (?m:...) (?i) ----
    ' gIC 是对象上的 IgnoreCase；内联选项应当能在它之上**局部**覆盖，且不影响外层。
    gIC = False
    ExtCase "e26", "(?i:abc)", "ABC", "t"
    ExtCase "e27", "(?i:abc)d", "ABCd", "t"
    ExtCase "e28", "(?i:abc)d", "ABCD", "t"          ' 作用域在 ) 处结束
    ExtCase "e29", "(?i)(?<x>ab)", "AB", "t"          ' 与命名组叠加
    ExtCase "e30", "(?-i:a)", "A", "t"               ' 对象开 IC、局部关
    ExtCase "e31", "(?-i:a)", "a", "t"
    ExtCase "e32", "(?s:.)", vbLf, "t"                 ' 局部 DotAll
    ExtCase "e33", ".", vbLf, "t"                     ' 基线：. 不吃换行
    ExtCase "e34", "(?m:^b$)", "a" & vbLf & "b", "t" ' 局部 Multiline
    ExtCase "e35", "(?i)[a-z]+", "ABC", "m"           ' 类也认内联选项
    ExtCase "e36", "(?i:(a)\1)", "aA", "t"            ' 反向引用在 icase 下也认大小写
    gIC = True
    ExtCase "e37", "abc", "ABC", "t"                  ' 对象级 IC 的基线
    ExtCase "e38", "(?-i:abc)", "ABC", "t"            ' 对象开 IC、整串局部关
    gIC = False
    ' ---- 条件组 (?(n)yes|no)：**没有 Python 对照**，按 PCRE2 语义手写期望（见 ext_oracle.py 注释）
    ExtCase "e39", "^(a)?(?(1)b|c)$", "ab", "t"      ' 1 组已闭合 -> 走 yes 支
    ExtCase "e40", "^(a)?(?(1)b|c)$", "c", "t"       ' 组未参与匹配 -> 走 no 支
    ExtCase "e41", "^(a)?(?(1)b|c)$", "ac", "t"      ' 组已闭合但仍要 b -> 不匹配
    ExtCase "e42", "^(?<n>\d+)(?(<n>)x|y)$", "12x", "t" ' 命名条件
    ExtCase "e43", "^(a)(?(1)b)$", "ab", "t"           ' 无 no 支
    ExtCase "e44", "^(a)(?(1)b|c)$", "ac", "t"
    ExtCase "e45", "(?(2)a|b)", "a", "t"              ' 引用不存在的组 -> 5017
    Close #gF2
End Sub

Sub PerfRun()
    ' 性能实测（不是门禁，是读数）：每项报 ms 并同时报“做了多少事”，
    ' 免得“快”来自根本没匹配。Timer 步进 ~15.6ms，所以每项故意重复到 >=0.5s。
    Dim re As Object, s As String, t0 As Single, ms As Single, i As Long, n As Long, mc As Object
    gF = FreeFile
    Open App.Path & "\re_perf.txt" For Output As #gF
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = False
    re.Multiline = False

    ' 1) 长串扫描：200k 字符，命中在末尾（最坏情况走完全程），20 轮
    s = String$(200000, "x") & "needle"
    re.Pattern = "needle"
    n = 0
    t0 = Timer
    For i = 1 To 20
        Set mc = re.Execute(s)
        n = n + mc.Count
    Next i
    ms = (Timer - t0) * 1000
    Print #gF, "perf|literal200k_x20|" & Fmt1(ms) & " ms|matches=" & n

    ' 2) 高频 Execute：短串 20000 次
    re.Pattern = "\d+"
    n = 0
    t0 = Timer
    For i = 1 To 20000
        Set mc = re.Execute("abc123def456")
        n = n + mc.Count
    Next i
    ms = (Timer - t0) * 1000
    Print #gF, "perf|execute_small_x20000|" & Fmt1(ms) & " ms|" & Fmt1(ms / 20) & " us/call|matches=" & n

    ' 3) 灾难回溯：^(a+)+b$ 在 24 个 a 上（真组件指数燃；本引擎应在 RE_STEP_LIMIT 处认输）
    re.Pattern = "^(a+)+b$"
    n = 0
    t0 = Timer
    For i = 1 To 20
        If re.Test(String$(24, "a")) Then n = n + 1
    Next i
    ms = (Timer - t0) * 1000
    Print #gF, "perf|catastrophic_a24_x20|" & Fmt1(ms) & " ms|matched=" & n

    ' 4) 灾难回溯加长：同样模式 40 个 a（超步数上限后应认输，而不是跑满）
    n = 0
    t0 = Timer
    For i = 1 To 5
        If re.Test(String$(40, "a")) Then n = n + 1
    Next i
    ms = (Timer - t0) * 1000
    Print #gF, "perf|catastrophic_a40_x5|" & Fmt1(ms) & " ms|matched=" & n

    ' 5) 交替回溯：^(a|aa)+$ 在 26 个 a 上（另一种经典燃法）
    re.Pattern = "^(a|aa)+$"
    n = 0
    t0 = Timer
    For i = 1 To 20
        If re.Test(String$(26, "a")) Then n = n + 1
    Next i
    ms = (Timer - t0) * 1000
    Print #gF, "perf|alternation_a26_x20|" & Fmt1(ms) & " ms|matched=" & n

    ' 6) 全局找：\d+\s+\w+ 在 16k 字的串上、200 处命中，20 轮
    s = ""
    For i = 1 To 2000
        s = s & "123 abc "
    Next i
    re.Pattern = "\d+\s+\w+"
    n = 0
    t0 = Timer
    For i = 1 To 20
        Set mc = re.Execute(s)
        n = n + mc.Count
    Next i
    ms = (Timer - t0) * 1000
    Print #gF, "perf|global_16k_x20|" & Fmt1(ms) & " ms|matches=" & n

    Close #gF
End Sub

Function Fmt1(ByVal v As Single) As String
    ' 一位小数，不依赖 Format$（免得为了一个读数把 RTL 面拖进来）
    Dim k As Long
    k = Int(v * 10 + 0.5)
    Fmt1 = CStr(k \ 10) & "." & CStr(k Mod 10)
End Function

Sub Main()
    gF = FreeFile
    If InStr(1, Command$, "/perf", vbTextCompare) > 0 Then
        PerfRun
        End
    End If
    Open App.Path & "\re_c3.txt" For Output As #gF
    Set gRe = CreateObject("VBScript.RegExp")

    ' ---- 基本 ----
    RunCase "abc", False, False, False, "xxabcxx", "exec", ""
    RunCase "abc", True, False, False, "abcabc", "exec", ""
    RunCase "abc", True, False, False, "ABCabc", "test", ""
    RunCase "abc", True, True, False, "ABCabc", "test", ""
    RunCase "", False, False, False, "abc", "test", ""
    RunCase "", True, False, False, "ab", "exec", ""
    ' ---- 点号 / 锚点 ----
    RunCase "a.b", False, False, False, "a" & vbLf & "b", "test", ""
    RunCase "a.b", False, False, False, "a" & vbCr & "b", "test", ""
    RunCase "^cd$", False, False, True, "ab" & vbLf & "cd", "test", ""
    RunCase "^cd$", False, False, False, "ab" & vbLf & "cd", "test", ""
    RunCase "^", True, False, True, "a" & vbLf & "b", "exec", ""
    RunCase "$", True, False, True, "a" & vbLf & "b", "exec", ""
    ' ---- 空匹配 ----
    RunCase "a*", True, False, False, "bbb", "exec", ""
    RunCase "x*", True, False, False, "aaa", "exec", ""
    ' ---- 字符类 ----
    RunCase "[a-c]+", True, False, False, "xxabcabczz", "exec", ""
    RunCase "[^a-c]+", True, False, False, "abcXYabc", "exec", ""
    RunCase "[\d\-\.]+", True, False, False, "a1-2.3b", "exec", ""
    RunCase "\d+", True, False, False, "a12b345", "exec", ""
    RunCase "\w+", True, False, False, "ab_cd-ef", "exec", ""
    RunCase "\s+", True, False, False, "a  b" & vbTab & "c", "exec", ""
    RunCase "[\x41-\x43]+", True, False, False, "zABz", "exec", ""
    RunCase "\u0041", True, False, False, "zAz", "exec", ""
    RunCase "[\b]", False, False, False, "a", "test", ""
    ' ---- 量词 ----
    RunCase "a{2,3}", True, False, False, "aaaaa", "exec", ""
    RunCase "a{2,3}?", True, False, False, "aaaaa", "exec", ""
    RunCase "a{2,}", True, False, False, "aaaa", "exec", ""
    RunCase "ab?c", True, False, False, "ac abc", "exec", ""
    ' ---- 分组 / 交替 ----
    RunCase "(a|b)+", True, False, False, "abxba", "exec", ""
    RunCase "(?:ab)+", True, False, False, "ababxab", "exec", ""
    RunCase "(a)(b)?", False, False, False, "a", "sub", ""
    RunCase "(a)|(b)", False, False, False, "b", "sub", ""
    RunCase "((a)(b))?", False, False, False, "ab", "sub", ""
    ' ---- 断言 ----
    RunCase "(?=ab)a", True, False, False, "ab ac", "exec", ""
    RunCase "(?!ab)a", True, False, False, "ab ac", "exec", ""
    RunCase "\bcd\b", False, False, False, "ab cd ef", "test", ""
    RunCase "\b", True, False, False, "ab cd", "exec", ""
    RunCase "\Bcd", False, False, False, "abcd", "test", ""
    ' ---- 后向引用 ----
    RunCase "(a)\1", False, False, False, "aa", "test", ""
    RunCase "(a)\1", False, False, False, "ab", "test", ""
    RunCase "(\w)\1", True, False, False, "aabbc", "exec", ""
    ' ---- 替换 ----
    RunCase "(\w+)@(\w+)", False, False, False, "x=ab@cd;", "repl", "[$&|$1|$2|$$|$`|$']"
    RunCase "(\w+)@(\w+)", True, False, False, "ab@cd ef@gh", "repl", "<$2:$1>"
    RunCase "o", True, True, False, "Foo BOO", "repl", "0"
    RunCase "x*", True, False, False, "abc", "repl", "-"
    ' ---- 属性回读 ----
    RunCase "a+b", False, True, True, "x", "pat", ""
    ' ---- 错误 ----
    RunCase "[", False, False, False, "a", "test", ""
    RunCase "a{3,2}", False, False, False, "a", "test", ""
    RunCase "(abc", False, False, False, "a", "test", ""
    RunCase "a)b", False, False, False, "a", "test", ""
    RunCase "((a)", False, False, False, "a", "test", ""

    Close #gF
    ExtCases
End Sub
