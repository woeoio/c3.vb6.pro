' 账 #604（第321刀）夹具: 混合标签的 Variant Select Case —— 标签箱的 BSTR 谁收回。
'
' 病灶（修复前实测）: 发码形 = `vb6_VARIANT _vb6_select_lbl_N = vb6_VariantString(vb6_BSTR_FromStr(L"…"));`
' —— 每执行一次这条 Select，每枚**字符串字面量标签**现造一只 BSTR，那枚临时 VARIANT 从不 Clear。
' 读数面 = kernel32!K32GetProcessMemoryInfo 的 PagefileUsage 增量（20 万轮 × 3 枚字面标签）：
'   修复前 x64 = 31,002,624 字节（判据 N）；修复后 = 0（判据 Y）。
'
' 另外四条是**所有权语义护栏**，专门钉住"不许顺手清借来的那几枚"：
'   S604-borrow / -borrow-after  裸串变量的标签箱是 vb6_VariantString(s) —— 别名，清下去就是放别人的块；
'   S604-slot / -slot-after      Variant 槽标签箱是 vb6_VariantFromValue(v) —— 浅别名，同上。
' 修复前/后这两条读数相同；把清收那一行改成"无条件清"之后 reassign 当场踩已释放块。
'
' 读数一律不带方括号: 套件用 PowerShell 的 -like 比针，[ABC] 会被当字符类。

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

Sub Main()
    Dim pmc As PMC
    Dim hr As Long, hr2 As Long
    Dim pf0 As Long, pf1 As Long
    Dim i As Long, acc As Long
    Dim v As Variant
    Dim s As String
    Dim v2 As Variant

    ' ---- 语义护栏: 先各走一遍, 读数取期望值 ----
    v = "alpha"
    Select Case v
        Case "alpha"
            Debug.Print "S604-lit="; 1
        Case 7
            Debug.Print "S604-lit="; 2
        Case "beta" To "betz"
            Debug.Print "S604-lit="; 3
        Case Else
            Debug.Print "S604-lit="; 0
    End Select

    s = "alpha"
    Select Case v
        Case s
            Debug.Print "S604-borrow="; 1
        Case 7
            Debug.Print "S604-borrow="; 2
        Case Else
            Debug.Print "S604-borrow="; 0
    End Select
    ' 借用标签那只缓冲不许被回收: 改一次再读回, 释放错当场双放/悬垂
    s = "omega"
    Debug.Print "S604-borrow-after="; s

    ' 槽标签 + 一枚串字面量把这一条 Select 拉进 Mixed 档: 槽标签走浅别名, 不许被 Clear
    v2 = "alpha"
    Select Case v
        Case v2
            Debug.Print "S604-slot="; 1
        Case "nomatch"
            Debug.Print "S604-slot="; 2
        Case 7
            Debug.Print "S604-slot="; 3
        Case Else
            Debug.Print "S604-slot="; 0
    End Select
    v2 = "zzz"
    Debug.Print "S604-slot-after="; v2

    v = 7
    Select Case v
        Case "alpha"
            Debug.Print "S604-num="; 1
        Case 7
            Debug.Print "S604-num="; 2
        Case Else
            Debug.Print "S604-num="; 0
    End Select

    v = "betq"
    Select Case v
        Case "alpha"
            Debug.Print "S604-range="; 1
        Case 7
            Debug.Print "S604-range="; 2
        Case "beta" To "betz"
            Debug.Print "S604-range="; 3
        Case Else
            Debug.Print "S604-range="; 0
    End Select

    ' ---- 泄漏引擎: 字面标签每次现造一只 BSTR ----
    v = "alpha"
    acc = 0
    For i = 1 To 50
        Select Case v
            Case "alpha"
                acc = acc + 1
            Case 7
                acc = acc + 100
            Case "beta" To "betz"
                acc = acc + 10000
        End Select
    Next i
    Debug.Print "S604-warm="; acc

    pmc.cb = Len(pmc)
    hr = K32GetProcessMemoryInfo(-1, pmc, Len(pmc))
    Debug.Print "S604-api1="; hr
    pf0 = CLng(pmc.PagefileUsage)

    For i = 1 To 200000
        Select Case v
            Case "alpha"
                acc = acc + 1
            Case 7
                acc = acc + 100
            Case "beta" To "betz"
                acc = acc + 10000
        End Select
    Next i

    pmc.cb = Len(pmc)
    hr2 = K32GetProcessMemoryInfo(-1, pmc, Len(pmc))
    Debug.Print "S604-api2="; hr2
    pf1 = CLng(pmc.PagefileUsage)
    Debug.Print "S604-delta="; (pf1 - pf0)
    If (hr = 1) And (hr2 = 1) And (pf1 >= pf0) And ((pf1 - pf0) < 2000000) Then
        Debug.Print "S604-leak-ok=Y"
    Else
        Debug.Print "S604-leak-ok=N"
    End If
    Debug.Print "S604-acc="; acc
    Debug.Print "S604-DONE"
End Sub
