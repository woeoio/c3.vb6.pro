' 账 #119 夹具: 局部/形参 String 目标 + 自有临时串 = 每句话漏一只 BSTR。
'
' 读数面 = kernel32!K32GetProcessMemoryInfo 的 PagefileUsage(私有提交字节), 判据问**增量**
' 而不是绝对值(绝对值随机器/运行时而异)。泄漏引擎是 `For i = 1 To N: s = Left(gBig, 64)`
' —— 每轮一只 64 字符的临时串, 修复前一只都不释放。
'
' 剩下的六条是语义护栏: AssignMove 换了所有权, 但不许换掉"赋值取的是当时那份值"的语义,
' 所以借用(变量/形参赋值)、别名(先抄后改)、ByRef 写穿、拼接链、返回串的过程这五条都逐条读数。
'
' 读数一律**不带方括号**: 套件用 PowerShell 的 -like 比针 (见 run_tests.ps1 的
' "output mismatch" 那一支), 而 -like 把 [ABC] 当**字符类** (匹配 A/B/C 中的一个字符),
' 针面写 [ABC] 就永远匹配不上 —— 门 #178 正是这么红的 (产品读数全对, 针自己不可能绿)。


Option Explicit

Private Type PMC
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As LongPtr
    WorkingSetSize As LongPtr
    QuotaPeakPagedPoolUsage As LongPtr
    QuotaPagedPoolUsage As LongPtr
    QuotaPeakNonPagedPoolUsage As LongPtr
    QuotaNonPagedPoolUsage As LongPtr
    PagefileUsage As LongPtr
    PeakPagefileUsage As LongPtr
End Type

Declare Function K32GetProcessMemoryInfo Lib "kernel32" (ByVal hProcess As LongPtr, ByRef ppmc As PMC, ByVal cb As Long) As Long

Dim gBig As String

Function MakeTag(n As Long) As String
    Dim t As String
    t = String$(n, 67)
    MakeTag = t
End Function

' ByRef String 形参做赋值目标: 所有权从被调方交到调用方那只变量上。
Sub SetOut(ByRef o As String)
    o = MakeTag(3)
End Sub

Sub Main()
    Dim pmc As PMC
    Dim hr As Long
    Dim pf0 As Long, pf1 As Long
    Dim i As Long
    Dim s As String, d As String, u As String
    Dim got As String

    gBig = String$(400, 65)

    pmc.cb = Len(pmc)
    hr = K32GetProcessMemoryInfo(-1, pmc, Len(pmc))
    Debug.Print "S119-api1="; hr
    pf0 = CLng(pmc.PagefileUsage)

    For i = 1 To 100000
        s = Left(gBig, 64)
    Next i

    pmc.cb = Len(pmc)
    hr = K32GetProcessMemoryInfo(-1, pmc, Len(pmc))
    Debug.Print "S119-api2="; hr
    pf1 = CLng(pmc.PagefileUsage)
    Debug.Print "S119-delta="; (pf1 - pf0)
    ' 两次读数都必须真的取到, 否则 delta 会是"自己减自己"的假零 ⇒ 判据当场就不成立。
    If (hr = 1) And (pf1 >= pf0) And ((pf1 - pf0) < 2000000) Then
        Debug.Print "S119-leak-ok=Y"
    Else
        Debug.Print "S119-leak-ok=N"
    End If

    ' --- 语义护栏 ---
    ' (1) 借用侧仍走深拷贝: 改源不许动已抄下的副本
    s = "SHARED-ME"
    d = s
    s = "CHANGED"
    Debug.Print "S119-borrow="; d

    ' (2) 自有侧换所有权之后再读, 值本身必须还在(不是悬垂也不是清零)
    s = MakeTag(5)
    Debug.Print "S119-func-len="; Len(s)

    ' (3) 拼接链
    u = "A" & "B" & "C"
    Debug.Print "S119-concat="; u

    ' (4) ByRef 写穿
    got = "OLD"
    SetOut got
    Debug.Print "S119-byref-len="; Len(got); " head="; Left(got, 1)

    ' (5) 同一只变量反复赋值不互相踩
    s = "AAA"
    s = s & "BBB"
    s = UCase(s)
    Debug.Print "S119-chain="; s

    ' (6) 泄漏引擎跑完之后进程还得能正常收尾(所有权改错当场就堆损坏)
    Debug.Print "S119-DONE"
End Sub
