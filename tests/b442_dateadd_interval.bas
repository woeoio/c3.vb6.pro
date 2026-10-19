Attribute VB_Name = "ModB442"
Option Explicit

' 账 #442 (§B211，第一百一十一刀) = **只钉行为、不改行为**：DateAdd 的「w / ww」那一格。
'
' 为什么开这一道: 账 #423 给 `DateDiff` 的 w / ww 补了「按 builtin 分岔」之后，隔壁 `DateAdd`
' 的同款分岔看起来**可疑** —— 共享枚举把单枚 `w` 答成 `vb6_di_weekday`，而 DateAdd 的 weekday
' 臂是 `datePart += n`（加**天**），于是「DateAdd("w",1,·) 只加一天」看着像个 bug。
'
' ⚠ **量完证明它是对的**（本机 cscript 同机代理逐格量，12 格零失配）：
'     DateAdd("w",  1, #2026/1/5#) = 2026-01-06   ← w  = 加**天**
'     DateAdd("ww", 1, #2026/1/5#) = 2026-01-12   ← ww = 加**周**
' 所以那一格**一个字都不改**。
'
' 这一刀同时**反证了 #423 的架构是对的**：那枚共享枚举对 **DateAdd 与 DatePart 两家**都恰好
' 是对的，只有 DateDiff 一家分岔 ⇒ #423 选「不动枚举、让 DateDiff 自己认」正是唯一正解。
' 若当时「顺手把枚举统一成 DateDiff 口径」，DateAdd 与 DatePart 会**同时**被弄坏。

Sub Main()
    ' --- 分岔本体: w 加天 / ww 加周 ---
    Debug.Print "w_p1=" & D("w", 1, #1/5/2026#)      ' 2026-1-6
    Debug.Print "ww_p1=" & D("ww", 1, #1/5/2026#)    ' 2026-1-12
    ' 负数 / 0 / 三周: 两档各自成比例（w 是天数、ww 是周数）
    Debug.Print "w_m1=" & D("w", -1, #1/5/2026#)     ' 2026-1-4
    Debug.Print "w_p3=" & D("w", 3, #1/5/2026#)      ' 2026-1-8
    Debug.Print "w_0=" & D("w", 0, #1/5/2026#)       ' 2026-1-5
    Debug.Print "ww_m2=" & D("ww", -2, #1/5/2026#)   ' 2025-12-22
    ' 跨月与闰月
    Debug.Print "w_cross=" & D("w", 1, #12/28/2026#)  ' 2026-12-29
    Debug.Print "w_leap=" & D("w", 1, #2/26/2024#)   ' 2024-2-27

    ' --- 四档对照: 共用那张表，改 w/ww 时最容易顺手碰到它们 ---
    Debug.Print "d_p1=" & D("d", 1, #1/5/2026#)      ' 2026-1-6
    Debug.Print "m_p1=" & D("m", 1, #1/5/2026#)      ' 2026-2-5
    Debug.Print "y_p1=" & D("y", 1, #1/5/2026#)      ' 2026-1-6
    Debug.Print "h_p1=" & D("h", 24, #1/5/2026#)     ' 2026-1-6

    Debug.Print "B442-DONE"
End Sub

Private Function D(iv As String, n As Long, dt As Date) As String
    Dim r As Date
    r = DateAdd(iv, n, dt)
    D = Year(r) & "-" & Month(r) & "-" & Day(r)
End Function