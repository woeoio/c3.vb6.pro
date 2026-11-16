Attribute VB_Name = "ReMain"
Option Explicit

' C30-10 正则引擎的 C3 侧用例 —— 与 ReCases.vbs **逐行复刻**，输出写文件后与 cscript 的输出 diff。
' 注意：这里刻意继续用 CreateObject("VBScript.RegExp")，走的是 RTL 里的内置实现
' （vb6com_regexp.c，vb6_CreateObject 抢先拦截），不加载 vbscript.dll。
' 要跑"真组件"对照组：set C3_REGEXP_PREFER_MS=1 再运行。

Dim gRe As Object
Dim gF As Integer

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

Sub Main()
    gF = FreeFile
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
    RunCase "(?<x>a)", False, False, False, "a", "test", ""
    RunCase "(abc", False, False, False, "a", "test", ""
    RunCase "a)b", False, False, False, "a", "test", ""
    RunCase "((a)", False, False, False, "a", "test", ""

    Close #gF
End Sub
