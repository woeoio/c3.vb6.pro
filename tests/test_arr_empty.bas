' <vbeclipse> 数组"空"与"缺少数组"两形对表 (用户在真 VB6 里实测确认):
'   Array()            -> 空数组: UBound=-1 / LBound=0 / UBound-LBound+1=0 / For Each 零次
'   UBound(<未分配>)   -> 运行时错误 9 (工程靠 On Error 接住, 见 List.cls)
' 常量实参那一形 (`UBound(vbNull)`) 是**编译期**错误, 见 tests/arr_neg/n01_ubound_vbnull.bas。
Option Explicit

Public Sub Main()
    Dim a()
    a = Array()
    Debug.Print "EA-ub="; UBound(a)
    Debug.Print "EA-lb="; LBound(a)
    Debug.Print "EA-n="; UBound(a) - LBound(a) + 1
    Debug.Print "EA-isarr="; IsArray(a)
    Debug.Print "EA-join=["; Join(a, ","); "]"

    Dim n As Long
    n = 0
    Dim x As Variant
    For Each x In a
        n = n + 1
    Next
    Debug.Print "EA-fe="; n

    Dim v As Variant
    v = Array()
    Debug.Print "EA-vub="; UBound(v)

    Dim w()
    w = Array(7, 8, 9)
    Debug.Print "EA-3ub="; UBound(w)
    Debug.Print "EA-e2="; w(2)

    Dim s() As String
    ReDim s(1 To 3)
    Debug.Print "EA-rb-lb="; LBound(s)
    Debug.Print "EA-rb-ub="; UBound(s)

    Dim u() As Long
    On Error GoTo Miss
    Debug.Print "EA-never="; UBound(u)
    Debug.Print "EA-DONE"
    Exit Sub
Miss:
    Debug.Print "EA-err="; Err.Number
    Debug.Print "EA-DONE"
End Sub
