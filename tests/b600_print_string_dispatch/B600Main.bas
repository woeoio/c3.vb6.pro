Attribute VB_Name = "B600Main"
Option Explicit

' B600: 第二步的回归夹具 —— Print # / Write # 的**字符串分流**。
' 旧实现靠一张 34 名的 C 出口名表 (bstrFuncs/isBstrExpr) 兜底; 第二步删掉它,
' 改为"语义类型 String + 发射形状 BSTR"两问。本夹具逐个钉住那些曾经靠名字兜底的
' builtin: 只要类型推断退化, 它们就会被当成数值走 vb6_Str((int32_t)) —— 要么编不过,
' 要么写盘字节变成数字 (方括号里一看就出来)。
' 末尾三项是**反用例**: 数值实参必须继续走数值路径 (前导/尾随空格 = Str 语义),
' 证明两问没有把 String 判据放宽到数值上。
Sub Main()
    Dim f As Integer
    Dim ln As String
    Dim n As Long

    f = FreeFile
    Open "b600_out.txt" For Output As #f
    Print #f, Left$("ABCDE", 2)
    Print #f, Right$("ABCDE", 2)
    Print #f, Mid$("ABCDE", 2, 2)
    Print #f, UCase$("abc")
    Print #f, LCase$("ABC")
    Print #f, Trim$("  x  ")
    Print #f, LTrim$("  x")
    Print #f, RTrim$("x  ")
    Print #f, Chr$(65)
    Print #f, CStr(3.25)
    Print #f, CStr(123)
    Print #f, Format$(1234.5, "0.00")
    Print #f, Hex$(255)
    Print #f, Oct$(8)
    Print #f, Replace$("abc", "b", "X")
    Print #f, Space$(3)
    Print #f, String$(3, "x")
    Print #f, StrReverse$("abc")
    Print #f, IIf(True, "yes", "no")
    ' 成员式 String (App.Title): 以前靠那两张名字表里的 "vb6_App_Title" 兜底,
    ' 删表后必须由 inferExprType 答 String 才能走字符串路径。
    Print #f, App.Title
    ' 反用例: 数值实参仍走数值路径
    Print #f, 123
    Print #f, 3.25
    Close #f

    f = FreeFile
    Open "b600_out.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        Debug.Print "[" & ln & "]"
        n = n + 1
    Loop
    Close #f

    Debug.Print "COUNT=" & n
    Debug.Print "DONE"
End Sub
