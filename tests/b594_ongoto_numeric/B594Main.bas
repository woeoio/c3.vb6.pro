Attribute VB_Name = "B594Main"
Option Explicit

' B7 回归: `On <expr> GoTo/GoSub 100, 200` —— 计算跳转的入口分发。
' VB6 语义: 索引 <1 或 >标签数时落到 On 语句的下一条。
Sub Main()
    Dim k As Integer
    Dim hits As Long

    k = 1
    On k GoTo 100, 200
    Debug.Print "F1"
100:
    Debug.Print "L100"

    k = 2
    On k GoTo 400, 300
    Debug.Print "F2"
300:
    Debug.Print "L300"

    k = 0
    On k GoTo 500, 600
    Debug.Print "F0"

    k = 3
    On k GoTo 700, 800
    Debug.Print "F3"

    k = 1
    On k GoSub 900, 1000
    Debug.Print "AFTER-SUB1"

    k = 2
    On k GoSub 900, 1000
    Debug.Print "AFTER-SUB2"
    GoTo 1100

900:
    hits = hits + 1
    Debug.Print "SUB900"
    Return
1000:
    hits = hits + 1
    Debug.Print "SUB1000"
    Return
1100:
    Debug.Print "HITS=" & hits
    Debug.Print "DONE"
    Exit Sub

200:
    Debug.Print "L200"
    Exit Sub
400:
    Debug.Print "L400"
    Exit Sub
500:
    Debug.Print "L500"
    Exit Sub
600:
    Debug.Print "L600"
    Exit Sub
700:
    Debug.Print "L700"
    Exit Sub
800:
    Debug.Print "L800"
End Sub
