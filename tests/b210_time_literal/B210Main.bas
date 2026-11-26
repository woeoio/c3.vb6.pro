Option Explicit

' §11.186 第二百零三刀: 纯时间形日期字面量 `#h:mm:ss AM/PM#`。
' 真机读数 `.temp/k202p/p247.txt`（VB6Mini /make + 运行，11 行）与
' `.temp/k202p/stdout247.txt`（同一份源经 Debug.Print 的 stdout 口径）逐字相同。
' 改前 C3 在**词法**就把 `#12:00:00 PM#` 当成文件号 `#12` 切下去（判据只认 `/` 与 `-`），
' 后面那串成为垃圾 token，连报 VB2002/VB2003/VB2004。解析侧的
' `foldDateLiteralToOADate` 本来就支持纯时间形，缺的只是词法那一格。
' 本夹具用 Debug.Print 出 stdout（run_tests.ps1 比的就是 stdout）。
Public Sub Main()
    On Error Resume Next

    Err.Clear: Debug.Print "T1 isdate=" & IsDate(#12:00:00 PM#) & " err=" & Err.Number
    Err.Clear: Debug.Print "T2 fmt=" & Format(#12:00:00 PM#, "hh:nn:ss") & " err=" & Err.Number
    Err.Clear: Debug.Print "T3 cdbl=" & CDbl(#12:00:00 PM#) & " err=" & Err.Number
    Err.Clear: Debug.Print "T4 fmt=" & Format(#1:05:00 PM#, "hh:nn:ss") & " err=" & Err.Number
    Err.Clear: Debug.Print "T5 fmt=" & Format(#12:30:00 AM#, "hh:nn:ss") & " err=" & Err.Number
    Err.Clear: Debug.Print "T6 fmt=" & Format(#11:59:59 PM#, "hh:nn:ss") & " err=" & Err.Number
    Err.Clear: Debug.Print "T7 datepart=" & Format(#12:00:00 PM#, "yyyy-mm-dd") & " err=" & Err.Number
    Err.Clear: Debug.Print "T8 cmp=" & (#12:00:00 PM# < #1:00:00 PM#) & " err=" & Err.Number
    Err.Clear: Debug.Print "T9 cdate=" & Format(CDate("12:00:00 PM"), "hh:nn:ss") & " err=" & Err.Number
    Err.Clear: Debug.Print "T10 cdbl=" & CDbl(#12:00:00 AM#) & " err=" & Err.Number
    Err.Clear: Debug.Print "T11 fmt=" & Format(#1:2:3 PM#, "hh:nn:ss") & " err=" & Err.Number
End Sub
