Option Explicit

' §11.175 第一百九十二刀: 赋值站的语句级中止 —— 右值里有「必然爬出被调方」的 raise 时，
' 目标那一步不写。真机读数见 .temp/k192p/P219.bas（VB6Mini /make 实跑）。
' 反例同在这枚夹具里: F2 (被调方自带 Resume Next) 与 F7 (自带 GoTo 处理器) 必须照写 ——
' 它们的 Err 序号同样会变，只靠「序号变了就中止」会把这两格写错。

Private dv As Long
Private mPub As Long

Private Function FEscape() As Long
    Dim z As Long
    z = 100 \ dv
    FEscape = 9
End Function

Private Function FNest() As Long
    On Error Resume Next
    Dim z As Long
    z = FEscape()
    FNest = 5
End Function

Private Function FHandled() As Long
    On Error GoTo H
    Dim z As Long
    z = 100 \ dv
    FHandled = 9
    Exit Function
H:
    Err.Clear
    FHandled = 3
End Function

Private Function FOuter() As Long
    FOuter = FEscape() + 1
End Function

Private Function FDbl() As Double
    Dim z As Long
    z = 100 \ dv
    FDbl = 2.5
End Function

Public Sub Main()
    dv = 0
    Dim n As Long, d As Double, e As Long

    On Error Resume Next

    n = -1: Err.Clear
    n = FEscape()
    e = Err.Number: Err.Clear
    Debug.Print "F1 n=" & CStr(n) & " err=" & e

    n = -1: Err.Clear
    n = FNest()
    e = Err.Number: Err.Clear
    Debug.Print "F2 n=" & CStr(n) & " err=" & e

    n = -1: Err.Clear
    n = FEscape() + 1
    e = Err.Number: Err.Clear
    Debug.Print "F3 n=" & CStr(n) & " err=" & e

    mPub = -1: Err.Clear
    mPub = FEscape()
    e = Err.Number: Err.Clear
    Debug.Print "F4 mPub=" & CStr(mPub) & " err=" & e

    n = -1: Err.Clear
    n = Abs(FEscape())
    e = Err.Number: Err.Clear
    Debug.Print "F6 n=" & CStr(n) & " err=" & e

    n = -1: Err.Clear
    n = FHandled()
    e = Err.Number: Err.Clear
    Debug.Print "F7 n=" & CStr(n) & " err=" & e

    n = -1: Err.Clear
    n = FOuter()
    e = Err.Number: Err.Clear
    Debug.Print "F8 n=" & CStr(n) & " err=" & e

    d = -1: Err.Clear
    d = FDbl()
    e = Err.Number: Err.Clear
    Debug.Print "F9 d=" & CStr(d) & " err=" & e

    n = -1: Err.Clear
    n = FEscape(): n = n + 1
    e = Err.Number: Err.Clear
    Debug.Print "F10 n=" & CStr(n) & " err=" & e
End Sub
