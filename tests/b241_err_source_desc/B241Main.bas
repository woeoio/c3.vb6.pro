Option Explicit

' §11.217 (第二百三十四刀): Err.Source / Err.Description 的「空串面」与「运行期 raise 面」。
' 真机读数 .temp/k225p/p329.txt（工程名 P329，本夹具的工程名是 B241Res）:
'   E1/E2 空: 没有错误时两者都是 ""（Len=0，Number=0）；
'   E3 用户 Err.Raise 5, "mySrc", "myDesc" ⇒ 源与描述**原样保留**；E4 Err.Clear 后全空；
'   E7 `n = CLng(v)`（v="x"）⇒ 13 / 工程名 / "Type mismatch"；
'   E9 `d = 1 / 0` ⇒ 11 / 工程名 / "Division by zero"；E10 清掉后 0/0/0。
' 工程名这一格是**整刀的靶心**: 改前 C3 交 ""（raise 站点根本没有源），真机交工程名。
Public Sub Main()
    Dim s As String
    On Error Resume Next

    s = "[" & Err.Source & "]"
    Debug.Print "E1=" & s & "/" & Len(Err.Source)
    s = "[" & Err.Description & "]"
    Debug.Print "E2=" & s & "/" & Len(Err.Description)

    Err.Raise 5, "mySrc", "myDesc"
    Debug.Print "E3=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number
    Err.Clear
    Debug.Print "E4=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number

    Dim v As Variant
    v = "x"
    Dim n As Long
    n = CLng(v)
    Debug.Print "E7=" & Err.Number & "/[" & Err.Source & "]/[" & Err.Description & "]"
    Err.Clear
    Debug.Print "E8=[" & Err.Source & "]/[" & Err.Description & "]/" & Err.Number

    Dim d As Double
    d = 1 / 0
    Debug.Print "E9=" & Err.Number & "/[" & Err.Source & "]/[" & Err.Description & "]"
    Err.Clear

    Debug.Print "E10=" & Len(Err.Source) & "/" & Len(Err.Description) & "/" & Err.Number
    On Error GoTo 0
End Sub