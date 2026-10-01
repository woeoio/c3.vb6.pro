Option Explicit

' ============================================================
'  test_overflow.bas - ai/009 5.10 (P3): 窄整型收窄赋值的溢出检查
'
'  背景: `Dim a As Byte: a = -1` 原先生成裸 C 赋值 `uint8_t a; a = (-1);`,
'  被 C **静默截断**成 255, 程序照跑不误 —— 属于静默错编, 不是风格问题。
'  VB6 这里抛 run-time error 6 "Overflow", 且能被 On Error 正常捕获。
'
'  范围口径 (三类宽度不可混, 见 cgen_base_type.cpp:30-31):
'    Byte    1 字节  uint8_t   0 .. 255
'    Integer 2 字节  int16_t   -32768 .. 32767
'    Long    4 字节  int32_t   -2147483648 .. 2147483647
'
'  两个方向都要锁住 (**反向护栏**):
'    越界 ⇒ 必须报 6 (否则就是原来那个静默截断的回归)
'    边界值 ⇒ 必须**不**报错 (255 / 32767 / 2147483647 都在范围内;
'      把 255 判成越界, 比不检查更糟 —— 那是把合法程序改成编不过)
'
'  已知取舍 (故意不测, 见 vb6rtl_conv.c 注释): On Error Resume Next 之下
'  越界赋值会把被截断的值真的存进目标, 而 VB6 是**放弃赋值、目标保持原值**。
'  所以下面只断言 Err.Number, 不断言 Resume Next 之后变量的值。
' ============================================================

' 共享给 ChkRange 的收窄目标 (模块级: ChkRange 与 Main 都要看到同一批变量,
' 这样 Main 末尾"边界值确实原样存住"的断言才是在检查**真正被检查过的那几个变量**)
Private b As Byte
Private i As Integer
Private l As Long
Private x As Long
Private y As Long

Private Sub ChkRange(tag As String, expectErr As Long)
    ' expectErr = 0 表示不该报错; =6 表示该报 Error 6
    Dim gotErr As Long
    On Error Resume Next
    Err.Clear
    Select Case tag
    Case "b-1":   b = -1
    Case "b-256": b = 256
    Case "b-255": b = 255
    Case "b-0":   b = 0
    Case "i-40000":  i = 40000
    Case "i-min":    i = -32768
    Case "i-max":    i = 32767
    Case "i-i32max": i = 2147483647
    Case "l-5e9":    l = 5000000000#
    Case "l-min":    l = -2147483647
    Case "l-max":    l = 2147483647
    Case "cb-300":  x = CByte(300)
    Case "cb-200":  x = CByte(200)
    Case "ci-40000": y = CInt(40000)
    Case "ci-30000": y = CInt(30000)
    End Select
    gotErr = Err.Number
    Err.Clear
    On Error GoTo 0
    If (gotErr = expectErr) Then
        Debug.Print "OVF-OK " & tag & " err=" & gotErr
    Else
        Debug.Print "OVF-BAD " & tag & " got=" & gotErr & " want=" & expectErr
    End If
End Sub

Sub Main()
    ChkRange "b-1", 6          ' Byte 负方向: 原来静默变 255
    ChkRange "b-256", 6        ' Byte 正方向溢出一格
    ChkRange "b-255", 0        ' 边界上界必须合法
    ChkRange "b-0", 0          ' 边界下界必须合法
    ChkRange "i-40000", 6
    ChkRange "i-i32max", 6     ' Long 的上界塞进 Integer
    ChkRange "i-min", 0
    ChkRange "i-max", 0
    ChkRange "l-5e9", 6
    ChkRange "l-min", 0
    ChkRange "l-max", 0
    ChkRange "cb-300", 6       ' CByte/CInt 在 VB6 里同样查表
    ChkRange "cb-200", 0
    ChkRange "ci-40000", 6
    ChkRange "ci-30000", 0

    ' --- 边界值确实原样存住, 不只是"没报错" ---
    ' Long 下界这里用 -2147483647 而不是 -2147483648: 后者撞上一个**与本次改动无关**
    ' 的既有 cgen 缺陷 (VB6 字面量 2147483648 超出 int32, 一元负号被直译成
    ' `--2147483648` → C2105; BASE 版同样复现)。等那条单独修。
    b = 255: i = -32768: l = 2147483647
    Debug.Print "OVF-VAL b=" & b & " i=" & i & " l=" & l
    b = 0: i = 32767
    Debug.Print "OVF-VAL2 b=" & b & " i=" & i

    ' --- On Error GoTo 能抓到 (不是只能 Resume Next) ---
    On Error GoTo Trap
    b = -1
    Debug.Print "OVF-BAD goto-did-not-fire"
    GoTo Done
Trap:
    Debug.Print "OVF-GOTO err=" & Err.Number
    Err.Clear
Done:
    On Error GoTo 0
    Debug.Print "OVERFLOW-DONE"
End Sub
