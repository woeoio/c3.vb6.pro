Attribute VB_Name = "ModB436"
Option Explicit

' 账 #436 (§B210，第一百零九刀) = 账 #429 (§B203) 的 ②③ 两格。①(尾随空格)已随上一刀转绿。
'
' ② **越界的 namedFormat 运行期直接报错**：`FormatDateTime(d, 9/99/-1)` ⇒ err 5。
'    改前落 `default:` 当 general date 照常交卷。承面 = 0（语料 44 处全用 vb* 常量）。
'
' ③ **不再拿「今天」兜底**。改前有**两条**兜底都 GetLocalTime()：
'    `dateSerial == 0.0` 与 `VariantTimeToSystemTime` 失败。
'    ⚠ 本刀量出账上没列、更根本的一条：**序列号 0 是合法日期（1899-12-30），不是哨兵**
'      （运行期 `FormatDateTime(0,2)` 交 1899/12/30，改前本仓答**今天**）。
'    ⚠ 第二条：**`VariantTimeToSystemTime` 成功 ≠ 运行期会交卷** —— 它能翻出 1601 年以前的年份，
'      而 GetDateFormatW 的下限就是 1601 ⇒ 改前那一段交**空串 + err=0**。
'      窗实测（cscript 同机代理逐格二分，尺 .build/k429e|429f|429g|429h.vbs）：
'        -60000→1735/9/21  -100000→1626/3/16  -109000→1601/7/25  -109200→1601/1/6
'        -109300→ERR5  -109400→ERR5  -110000→ERR5  -120000→ERR5
'      ⇒ 边界 = **年份 1601**（1601-01-01 的序列号约 -109205），与 GetDateFormatW 文档下限同位。
'
' ⚠ **本夹具一个字都不钉格式化字面量**：namedFormat 1/2/4 三档全是 **locale 档**
'   （本机 zh-CN 读 `2026/1/5`、CI 那台 en-US 读 `1/5/2026`；短时间那边甚至可能是 `2:30 PM`），
'   钉字面量必假红 —— 与 §B142「`VB-dec-fmt` 照打不钉(区域设置)」同一个坑，本刀踩了两次。
'   所以：**值**走 Year/Month/Day 三枚数值 getter 钉，**报不报错**走 Err 钉，两者分开。

Sub Main()
    ' --- ② 窗内四档: 调得通、不报错(nf=1/2/4) ---
    Debug.Print "ok1=" & NoErr(46027.6046875, 1)
    Debug.Print "ok2=" & NoErr(46027.6046875, 2)
    Debug.Print "ok4=" & NoErr(46027.6046875, 4)
    ' 越界档位 ⇒ err 5
    Debug.Print "bad9=" & One(46027.6046875, 9)          ' ERR5
    Debug.Print "bad99=" & One(46027.6046875, 99)        ' ERR5
    Debug.Print "badneg=" & One(46027.6046875, -1)       ' ERR5

    ' --- ③ 序列号 0 是合法日期（改前答今天）---
    Debug.Print "zero-y=" & YearOf(0) & "/" & MonthOf(0) & "/" & DayOf(0)     ' 1899/12/30
    ' 窗外: 报 err 5，不交空串装没事
    Debug.Print "out1601=" & One(-109300, 2)             ' ERR5
    ' 窗内: 照常交卷
    Debug.Print "in1601-y=" & YearOf(-109000)             ' 1601
    Debug.Print "neg1-y=" & YearOf(-1) & "/" & MonthOf(-1) & "/" & DayOf(-1)  ' 1899/12/29

    ' --- 负控: 零头为 0 的 general date 不许多尾随空格（① 已转绿，盯住别退）---
    Debug.Print "NC-nospace=" & (IIf(Len(FormatDateTime(DateSerial(2026, 1, 5), 0)) = 8, 1, 0))  ' 1

    Debug.Print "B436-DONE"
End Sub

Private Function One(v As Double, nf As Long) As String
    Dim r As String
    Dim n As Long
    On Error Resume Next
    Err.Clear
    r = FormatDateTime(v, nf)
    n = Err.Number
    Err.Clear
    On Error GoTo 0
    If n <> 0 Then
        One = "ERR" & n
    Else
        One = "OK"
    End If
End Function

Private Function NoErr(v As Double, nf As Long) As String
    NoErr = One(v, nf)
End Function

Private Function YearOf(v As Double) As Long
    YearOf = Year(CDate(v))
End Function

Private Function MonthOf(v As Double) As Long
    MonthOf = Month(CDate(v))
End Function

Private Function DayOf(v As Double) As Long
    DayOf = Day(CDate(v))
End Function