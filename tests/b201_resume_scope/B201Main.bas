Option Explicit

' §11.176 第一百九十三刀: Resume Next 的作用面 = **过程级**。
' raise 一旦发生，被调方的剩余语句与调用方这条语句都不再执行；控制流落到最近的
' Resume Next 帧，从「出错语句的下一条」继续，且**隐式 resume 不清 Err**。
' 真机读数见 .temp/k193p/P220.bas（VB6Mini /make 实跑，读数 p220.txt）。

Private dv As Long
Private hits As Long

Private Sub PEscape()
    Dim z As Long
    z = 100 \ dv
    hits = hits + 100
End Sub

Private Sub PSwallow()
    On Error Resume Next
    Dim z As Long
    z = 100 \ dv
    Err.Clear
    hits = hits + 1
End Sub

Private Function FEscape() As Long
    Dim z As Long
    z = 100 \ dv
    hits = hits + 1000
    FEscape = 9
End Function

Private Function FMid() As Long
    hits = hits + 10
    FMid = FEscape()
    hits = hits + 20
End Function

Private Sub PLastErr()
    On Error Resume Next
    Dim z As Long
    z = 100 \ dv
End Sub

Public Sub Main()
    dv = 0
    Dim e As Long, n As Long

    On Error Resume Next

    hits = 0: Err.Clear
    Call PEscape
    e = Err.Number: Err.Clear
    Debug.Print "G1 call_escape: hits=" & hits & " err=" & e

    hits = 0: Err.Clear
    Call PSwallow
    e = Err.Number: Err.Clear
    Debug.Print "G2 call_swallow: hits=" & hits & " err=" & e

    hits = 0: Err.Clear
    n = FEscape()
    e = Err.Number: Err.Clear
    Debug.Print "G3 assign_escape: n=" & n & " hits=" & hits & " err=" & e

    hits = 0: Err.Clear
    n = FMid()
    e = Err.Number: Err.Clear
    Debug.Print "G4 two_hop: n=" & n & " hits=" & hits & " err=" & e

    hits = 0: Err.Clear
    n = FEscape()
    hits = hits + 1
    e = Err.Number: Err.Clear
    Debug.Print "G5 next_stmt: n=" & n & " hits=" & hits & " err=" & e

    hits = 0: Err.Clear
    PLastErr
    hits = hits + 1
    e = Err.Number: Err.Clear
    Debug.Print "G6 last_stmt: hits=" & hits & " err=" & e
End Sub
