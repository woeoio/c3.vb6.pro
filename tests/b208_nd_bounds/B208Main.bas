Option Explicit

' §11.183 第二百刀: 多维数组的 UBound / LBound —— 载体与维号同一套错误口径。
' 真机读数 `.temp/k200p/p239.txt`（VB6Mini /make）：
'   · 未分配（N1）与 Erase 之后（N4）的多维数组，UBound/LBound 一律 **9**；
'   · 维号越界（N3 的 3 / 0 / -1）也是 **9**；
'   · 1D 数组上 dim≠1（N5 的 UBound(b,2) / LBound(b,2) / UBound(b,0)）同样 **9**。
' 改前 C3 这 19 行全部答 err=0（ND 那两枚对 NULL/越界维号直接返 0；1D 那两枚无视 dimension）。
' 用 Debug.Print 出到 stdout（夹具 harness 比的就是 stdout）。
' 注: N5 的 UBound(b) 在 C3 侧会走 ND 入口 —— 同一过程里出现过 UBound(b,2) 时，
'     发码侧 scanNDArraysInExpr 按名字把同名的 1D 数组也发成 ND 版。§11.183 让 ND 入口
'     先认 1D 描述符（按自己的秩答），所以这里仍答 4，与真机相同。

Public Sub Main()
    Dim a2() As Long
    Dim b() As Long
    Dim e As Long

    On Error Resume Next

    ' ---- N1: 未分配的多维 ----
    Err.Clear: e = UBound(a2)
    Debug.Print "N1 ub err=" & Err.Number
    Err.Clear: e = LBound(a2)
    Debug.Print "N1 lb err=" & Err.Number
    Err.Clear: e = UBound(a2, 2)
    Debug.Print "N1 ub2 err=" & Err.Number

    ' ---- N2: 已分配 2x3（界要答对）----
    ReDim a2(1 To 2, 1 To 3)
    Err.Clear: Debug.Print "N2 ub1=" & UBound(a2, 1) & " err=" & Err.Number
    Err.Clear: Debug.Print "N2 lb1=" & LBound(a2, 1) & " err=" & Err.Number
    Err.Clear: Debug.Print "N2 ub2=" & UBound(a2, 2) & " err=" & Err.Number
    Err.Clear: Debug.Print "N2 lb2=" & LBound(a2, 2) & " err=" & Err.Number
    Err.Clear: Debug.Print "N2 ub0d=" & UBound(a2) & " err=" & Err.Number

    ' ---- N3: 维号越界 / 维号 0 / 负维号 ----
    Err.Clear: e = UBound(a2, 3)
    Debug.Print "N3 ub3 err=" & Err.Number
    Err.Clear: e = LBound(a2, 3)
    Debug.Print "N3 lb3 err=" & Err.Number
    Err.Clear: e = UBound(a2, 0)
    Debug.Print "N3 ub0 err=" & Err.Number
    Err.Clear: e = UBound(a2, -1)
    Debug.Print "N3 ubm1 err=" & Err.Number

    ' ---- N4: Erase 之后 ----
    Erase a2
    Err.Clear: e = UBound(a2)
    Debug.Print "N4 ub err=" & Err.Number
    Err.Clear: e = LBound(a2)
    Debug.Print "N4 lb err=" & Err.Number
    Err.Clear: e = UBound(a2, 2)
    Debug.Print "N4 ub2 err=" & Err.Number

    ' ---- N5: 1D 对照 ----
    ReDim b(0 To 4)
    Err.Clear: Debug.Print "N5 ub=" & UBound(b) & " err=" & Err.Number
    Err.Clear: e = UBound(b, 2)
    Debug.Print "N5 ub2 err=" & Err.Number
    Err.Clear: e = LBound(b, 2)
    Debug.Print "N5 lb2 err=" & Err.Number
    Err.Clear: e = UBound(b, 0)
    Debug.Print "N5 ub0 err=" & Err.Number
End Sub
