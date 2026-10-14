Option Explicit

' 账 #220: B / BF 在 VB6 里是完全合法的模块级变量名 (`For B = 1 To 3` 这种写法到处都是),
' 而 RTL 曾在文件作用域导出两枚**外部链接**的 C 全局 (const int32_t B = 1; const int32_t
' BF = 2;) 去接住 Picture.Line 发码原样吐出的语法旗标。于是用户模块一发 `Public B As Long`
' 就撞成 C2373 重定义 + C2166 给 const 赋值, 连 exe 都出不来 (改前实测: BUILD-RC=1 /
' 5 条诊断 / no exe)。旗标现在由 parser 在 Line 的 style 位置折成字面量, RTL 不再需要名字。
' NC-* 钉的是「这两个名字真归用户了」: 数值、字符串长度、累加、装箱四形都按用户的意思走。

Public B As Long
Public BF As String
' 账 #219 第二半: 同一族的另外六枚。这些名字改前是 RTL 文件作用域的**裸名**
' (g_hoCount / g_uc_descs / g_hInstance / ocxCreateAny / twipsToHimetric, 外加改前 extern 过的
' Changed) —— 用户模块级变量在 C 里也是裸名, 于是撞成 C2371(头里 extern) 或 LNK2005+LNK1169(.c 里
' 非 static 定义), 两台都出不了 exe。现在这些内部名字一律带 vb6_ 前缀, 名字归还给用户。
Public g_hoCount As Long
Public g_uc_descs As Long
Public g_hInstance As Long
Public ocxCreateAny As Long
Public twipsToHimetric As Long
Public Changed As Long

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

    g_hoCount = 100
    g_uc_descs = 200
    g_hInstance = 300
    ocxCreateAny = 400
    twipsToHimetric = 500
    Changed = 600
    g_hoCount = g_hoCount + 1
    Dim clashSum As Long
    clashSum = g_hoCount + g_uc_descs + g_hInstance + ocxCreateAny + twipsToHimetric + Changed
    ' 101+200+300+400+500+600 = 2101 —— 六枚各存各的, 没有一枚被 RTL 那份同名存储吃掉
    Debug.Print "NC2-SUM=" & clashSum
    Debug.Print "NC2-INDEP=" & g_hoCount & "/" & Changed
    Debug.Print "NC-DONE"
End Sub
