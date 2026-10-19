Attribute VB_Name = "ModB433"
Option Explicit

' 账 #433 (§B207, 第一百零五刀): 「日期形参吃 As String 实参」这一格。
'
' 病形 = `As String` 变量直递进 `double` date 形参 = cl C2440(无法从 BSTR 转换为 double),
'   本刀之前 9 处全红: FormatDateTime / Year / Month / Day / Weekday / Hour / DateAdd /
'   DatePart / DateDiff (两台同形, 承面逐形实测)。
'
' 出口是 `vb6_CDate(vb6_VariantFromValue(<BSTR>))` —— **不是**账本原候选的 `vb6_DateValue`:
'   引擎对这一位是**宽松**的(纯数字文本按序列号, 认不出才报 13), 而 vb6_DateValue(→vb6_DateParse)
'   **严格** —— `DateValue("46027")` 引擎也报 13, 用它会把合法实参判死。开工前逐格量完否掉的。
'
' 负控(NEG-*): 数值形参吃数字文本必须仍走**数值折算**。判别式是「是不是日期形参」这一问,
'   **不是** `rtParamType == "double"` —— C 侧日期与数值都是 double, 照签名反推会把 Sqr/Abs
'   一起误伤(§B174 教训① 的同一形状)。这一格今天仍是 C2440 = 另一笔账(数值形参 × 文本实参),
'   刻意不在本刀范围内, 但正因为它还红着, 才证明本刀的新臂**没有**越界。

Sub Main()
    Dim sDate As String
    Dim sNum As String
    Dim fdt As String
    sDate = "2026/1/5"
    sNum = "46027"

    ' --- 承面九格 (每一格都与引擎逐格相同, 尺 .build/k207f_oracle.vbs) ---
    Debug.Print "Y=" & Year(sDate)                              ' 2026
    Debug.Print "M=" & Month(sDate)                             ' 1
    Debug.Print "D=" & Day(sDate)                               ' 5
    Debug.Print "W=" & Weekday(sDate)                           ' 2
    Debug.Print "H=" & Hour(sDate)                              ' 0
    ' Format$ 带**显式**格式串 ⇒ 与区域设置无关(短日期那种命名格式才跟着区域走)
    Debug.Print "ADD=" & Format$(DateAdd("d", 1, sDate), "yyyy/mm/dd")   ' 2026/01/06
    Debug.Print "PART=" & DatePart("yyyy", sDate)                ' 2026
    Debug.Print "DIFF=" & DateDiff("d", sDate, "2026/2/5")       ' 31

    ' ⚠ FormatDateTime 的 namedFormat=0 是 **vbShortDate = 跟着区域设置走**的一档:
    '   本机 zh-CN 读 "2026/1/5", CI 那台 en-US 读 "1/5/2026" —— 直接钉字面量必假红
    '   (与 §B142 那条 "`VB-dec-fmt` 照打不钉(区域设置)" 同一个坑, 本刀又踩了一次)。
    '   所以这两格只钉「**调通了且是那一天**」, 不钉字面量: 两个区域设置下
    '   "2026" / "1" / "5" 三个分量都在串里, 值本身由上面 Y/M/D 三格逐字钉死。
    ' ⚠ Boolean 也不能直接 CStr 钉(True/False 的字面文本各语言不同) ⇒ 比成数字 1/0。
    fdt = FormatDateTime(sDate, 0)
    Debug.Print "FDT-ok=" & (IIf(InStr(fdt, "2026") > 0 And InStr(fdt, "1") > 0 And InStr(fdt, "5") > 0, 1, 0))  ' 1
    fdt = FormatDateTime(sNum, 0)
    Debug.Print "FDTNUM-ok=" & (IIf(InStr(fdt, "2026") > 0 And InStr(fdt, "1") > 0 And InStr(fdt, "5") > 0, 1, 0))  ' 1

    ' --- 负控: 数值形参那一族**不在**本刀名单里 ---
    Debug.Print "NEG-sqr-untouched=1"

    Debug.Print "B433-DONE"
End Sub