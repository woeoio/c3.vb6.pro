Attribute VB_Name = "B596Main"
Option Explicit

' B11 回归: Print # 的 String 语义判定必须走**表达式类型**, 不能靠发射形状名单.
' 修复前实测: ByRef String 形参 s 被写成 -1092513192 (BSTR 指针截成 int32);
' 字面量/拼接/Format$ 当时恰好在名单内, 一并钉住. 数字是反向用例:
' 12345 必须按数值写盘, 不能当成 BSTR.
Sub Main()
    Dim f As Integer
    Dim a As String
    Dim b As String
    Dim ln As String
    Dim n As Long

    f = FreeFile
    Open "b596_out.txt" For Output As #f
    Print #f, "LIT"                     ' 字面量
    Print #f, "A" & "B"                 ' 字面量拼接
    a = "C": b = "D"
    Print #f, a & b                     ' 变量拼接
    Call W(f, "EF")                     ' ByRef String 形参
    Print #f, Format$(1234.5, "0.00")   ' Format$
    Print #f, 12345                     ' 数字 (反向用例)
    Close #f

    ' 读回断言: 逐行读回并加方括号打印, 空行/缺失行一眼可见.
    f = FreeFile
    Open "b596_out.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        Debug.Print "[" & ln & "]"
        n = n + 1
    Loop
    Close #f

    Debug.Print "COUNT=" & n
    Debug.Print "DONE"
End Sub

Sub W(ByRef f As Integer, ByRef s As String)
    Print #f, s
End Sub
