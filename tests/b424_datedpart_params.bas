Attribute VB_Name = "ModB424"
Option Explicit

' 账 #424 (§B198，第一百零九刀) = DatePart/DateDiff 那两条形参的两档口径 —— **由用户拍板**，不是实现默默选的。
'
' ① `firstWeekOfYear >= 4`：运行期**直接报错**（cscript 实测 fwy=4 / 5 两格当场抛，
'    尺 .build/k424_oracle.vbs）；改前本仓静默按 1 档答 2。
'    口径 = **对齐运行期，报 err 5**（与 #436「越界 namedFormat 报 err 5」同一口径）。
'
' ② `firstDayOfWeek = 0` (vbGeneral)：运行期**交系统**（本机 = 周一）⇒ 任何「跟随系统」的
'    发法都**跨机器不可复现**。口径 = **继续 clamp 到 vbSunday(1)**，作为一条**显式的已知偏离**：
'    在「系统周首 ≠ 周日」的机器上与运行期不同，换来读数跨机器稳定 —— 这正是本仓整套
'    emit-manifest / 夹具逐字节对账赖以成立的前提。
'    ⚠ 本机量不出这一格的差别（DatePart("ww") 的 f=0/1/2/7 **同答 2**、DateDiff("ww") 的
'      f=0/1/7 **同答 8**）⇒ 它是一处**有意保留**的偏离，不是漏改。

Sub Main()
    ' --- ① 合法档 0..3 逐格与运行期相同（0 与 1 同档；2 / 3 另一档）---
    Debug.Print "fwy0=" & Fwy(0)          ' 2
    Debug.Print "fwy1=" & Fwy(1)          ' 2
    Debug.Print "fwy2=" & Fwy(2)          ' 1
    Debug.Print "fwy3=" & Fwy(3)          ' 1
    ' 越界 ⇒ err 5（改前静默答 2）
    Debug.Print "fwy4=" & Fwy(4)          ' ERR5
    Debug.Print "fwy5=" & Fwy(5)          ' ERR5
    Debug.Print "fwy9=" & Fwy(9)          ' ERR5

    ' --- ② 已知偏离: fdow 统一 clamp 到周日，本机与运行期同读数 ---
    Debug.Print "NC-fdow0=" & Fdow(0)      ' 2
    Debug.Print "NC-fdow7=" & Fdow(7)      ' 2

    Debug.Print "B424-DONE"
End Sub

Private Function Fwy(w As Long) As String
    Dim v As Long
    Dim n As Long
    On Error Resume Next
    Err.Clear
    v = DatePart("ww", #1/5/2026#, 1, w)
    n = Err.Number
    Err.Clear
    On Error GoTo 0
    If n <> 0 Then
        Fwy = "ERR" & n
    Else
        Fwy = CStr(v)
    End If
End Function

Private Function Fdow(f As Long) As String
    Dim v As Long
    Dim n As Long
    On Error Resume Next
    Err.Clear
    v = DatePart("ww", #1/5/2026#, f, 1)
    n = Err.Number
    Err.Clear
    On Error GoTo 0
    If n <> 0 Then
        Fdow = "ERR" & n
    Else
        Fdow = CStr(v)
    End If
End Function