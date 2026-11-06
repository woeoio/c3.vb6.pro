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
    ' 只有时间的 Date: Print # 只写时间([0,1) 那一档), 不写 1899-12-30 的日期前缀, 恒一个尾空格。
    ' 判据不再用"CStr(整个) 去掉 CStr(1899-12-30) 前缀"那种按长度做算术的写法 —— 前缀长度
    ' 随 locale 的日期段长度/分隔符而变, 且一旦 Print # 那侧读到的是活动时钟, 跨秒就会漂。
    ' 改成三条结构断言 (仍守原意「这一行 = 该 Date 的时间文本 + 一个尾空格」):
    '   ① 尾恒一个空格 (RTrim 后加回一个); ② 行内不得出现 CStr(1899-12-30) 那段日期前缀;
    '   ③ 时间文本本身是 24 小时制 HH:MM:SS, 且三个字段就是 TimeSerial(15,4,5)。
    ' ③ 不能拿 CStr(dt) 逐字比: en-US 的 CStr 会把只含时间的 Date 写成
    ' "12/30/1899 3:04:05 PM"（12 小时制 + 1899 日期前缀），与 Print # 的 24 小时制口径
    ' 本就不同（CI 实测 lns(14)=[15:04:05 ]、lns(15)=[12/30/1899 3:04:05 PM]）；
    ' zh-CN 下两者恰好同形只是巧合，不是判据可以依赖的口径。
    Dim t14 As String
    Dim tShape As Boolean
    t14 = RTrim$(lns(14))
    tShape = (Len(t14) = 8)
    If tShape Then
        tShape = (Mid$(t14, 3, 1) = ":") And (Mid$(t14, 6, 1) = ":") _
             And (Left$(t14, 2) = "15") And (Mid$(t14, 4, 2) = "04") And (Right$(t14, 2) = "05")
    End If
    If (lns(14) = t14 & " ") _
       And (InStr(lns(14), CStr(DateSerial(1899, 12, 30))) = 0) _
       And tShape Then
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
