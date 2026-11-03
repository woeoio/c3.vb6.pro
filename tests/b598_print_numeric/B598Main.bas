Attribute VB_Name = "B598Main"
Option Explicit

' B598: Print # 的数值/日期/布尔/Empty/Null 必须按**真 VB6 实测字节**写盘
' (参照工程 .temp/m4/vb6num, 逐文件 dump)。修复前是 (int32_t) 截断:
' 3.25 -> "3", 1E+20 -> "-2147483648", Date -> 43832, True -> -1。
' 断言走"写盘后逐行读回 + 方括号", 前导/尾随空格与 CRLF 位置都可见。
' Date 那一档不钉区域形状的字面量: Print # 的 Date 就是**本机 locale 的 CStr** + 尾空格,
' 所以夹具在同一个进程里再写一行 CStr(dt) 做对照 —— 两边同一 locale, 跨机器不会假红
' (哨兵 scripts/check_no_locale_needles.ps1 的口径)。
Sub Main()
    Dim f As Integer
    Dim ln As String
    Dim n As Long
    Dim s As Single
    Dim d As Double
    Dim c As Currency
    Dim v As Variant
    Dim e As Variant
    Dim dt As Date
    Dim b As Boolean
    Dim lns(1 To 40) As String
    Dim i As Long

    f = FreeFile
    Open "b598_out.txt" For Output As #f
    Print #f, 123
    Print #f, -123
    Print #f, 0
    s = 3.25: Print #f, s
    d = 3.1415926535: Print #f, d
    c = 1.2345: Print #f, c
    v = CDec(3.25): Print #f, v
    Print #f, 1E+20
    Print #f, 1E-20
    dt = DateSerial(2020, 1, 2) + TimeSerial(15, 4, 5): Print #f, dt: Print #f, CStr(dt)
    dt = DateSerial(2020, 1, 2): Print #f, dt: Print #f, CStr(dt)
    dt = TimeSerial(15, 4, 5): Print #f, dt: Print #f, CStr(dt)
    b = True: Print #f, b
    b = False: Print #f, b
    Print #f, e
    Print #f, Null
    Print #f, 1; 2
    Print #f, 1, 2
    Close #f

    f = FreeFile
    Open "b598_out.txt" For Input As #f
    Do While Not EOF(f)
        n = n + 1
        Line Input #f, lns(n)
    Loop
    Close #f

    For i = 1 To 9
        Debug.Print "[" & lns(i) & "]"
    Next i
    If lns(10) = lns(11) & " " Then
        Debug.Print "DATE_TIME=OK"
    Else
        Debug.Print "DATE_TIME=FAIL"
    End If
    If lns(12) = lns(13) & " " Then
        Debug.Print "DATE_ONLY=OK"
    Else
        Debug.Print "DATE_ONLY=FAIL"
    End If
    ' 只有时间的 Date: Print # 只写时间([0,1) 那一档), 而 CStr 会带上 1899-12-30 的日期。
    ' 用"CStr(整个) 去掉 CStr(1899-12-30) 那段前缀"取时间, 跨 locale 都成立。
    If lns(14) = Mid$(lns(15), Len(CStr(DateSerial(1899, 12, 30))) + 2) & " " Then
        Debug.Print "TIME_ONLY=OK"
    Else
        Debug.Print "TIME_ONLY=FAIL"
    End If
    For i = 16 To 21
        Debug.Print "[" & lns(i) & "]"
    Next i

    Debug.Print "COUNT=" & n
    Debug.Print "DONE"
End Sub
