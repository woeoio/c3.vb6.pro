Attribute VB_Name = "B599Main"
Option Explicit

' B599: Write # 是**另一套**语义, 不能与 Print # 共用格式化:
'   - 数值去 Str 前导空格, |x|<1 去前导 0 (.25 / -.25)
'   - Date 恒 "#YYYY-MM-DD HH:MM:SS#" (只有时间时是 1899-12-30)
'   - Boolean #TRUE#/#FALSE#, Null #NULL#, Empty 写空
'   - 字符串加引号且内部 " 双写; 项间恒逗号; 语句末一个 CRLF
' 修复前实测: `Write #f, 3.25` 写盘为 "3", —— 带引号、逗号, 且没有 CRLF。
Sub Main()
    Dim f As Integer
    Dim ln As String
    Dim n As Long
    Dim s As Single
    Dim d As Double
    Dim v As Variant
    Dim e As Variant
    Dim dt As Date

    f = FreeFile
    Open "b599_out.txt" For Output As #f
    Write #f, 123
    Write #f, -123
    s = 3.25: Write #f, s
    Write #f, 0.25
    Write #f, -0.25
    d = 1 / 3: Write #f, d
    Write #f, 1E+20
    dt = DateSerial(2020, 1, 2) + TimeSerial(15, 4, 5): Write #f, dt
    dt = TimeSerial(15, 4, 5): Write #f, dt
    Write #f, True
    Write #f, Null
    Write #f, e
    Write #f, "A""B"
    Write #f, 1; "x"
    Close #f

    f = FreeFile
    Open "b599_out.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        Debug.Print "[" & ln & "]"
        n = n + 1
    Loop
    Close #f

    Debug.Print "COUNT=" & n
    Debug.Print "DONE"
End Sub
