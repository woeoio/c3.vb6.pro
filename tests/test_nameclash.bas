Option Explicit

' 账 #220: B / BF 在 VB6 里是完全合法的模块级变量名 (`For B = 1 To 3` 这种写法到处都是),
' 而 RTL 曾在文件作用域导出两枚**外部链接**的 C 全局 (const int32_t B = 1; const int32_t
' BF = 2;) 去接住 Picture.Line 发码原样吐出的语法旗标。于是用户模块一发 `Public B As Long`
' 就撞成 C2373 重定义 + C2166 给 const 赋值, 连 exe 都出不来 (改前实测: BUILD-RC=1 /
' 5 条诊断 / no exe)。旗标现在由 parser 在 Line 的 style 位置折成字面量, RTL 不再需要名字。
' NC-* 钉的是「这两个名字真归用户了」: 数值、字符串长度、累加、装箱四形都按用户的意思走。

Public B As Long
Public BF As String
Private acc As Long

Public Sub Main()
    Dim i As Long

    B = 7
    BF = "xy"
    For i = 1 To 3
        B = B + i
    Next
    acc = B + Len(BF)

    Debug.Print "NC-B=" & B & " NC-BFLEN=" & Len(BF)
    Debug.Print "NC-ACC=" & acc
    Debug.Print "NC-BOX=" & VarType(B) & "/" & VarType(BF)
    Debug.Print "NC-DONE"
End Sub
