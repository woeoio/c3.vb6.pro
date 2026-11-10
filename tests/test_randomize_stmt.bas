Option Explicit

' Randomize / Rnd 与真 VB6 的**逐值对齐**。
' oracle = 真 VB6 6.00.9782 编同一段代码打出的 24-bit 内部状态序列
' (Hex$(CLng(Rnd() * 2^24)) 精确还原 LCG 的 state)。断言钉的是数值, 不是"能编译"。
' 被钉住的语义:
'   · Randomize n  = (double)n 高 32 位折半 XOR 播种, 且**沿用旧状态低字节**
'   · Randomize 0  = 确定性 fold=0 (不是 Timer)
'   · 裸 Randomize = Timer 的 Single 位型播种 (只能验值域)
'   · Rnd(0) 不推进; Rnd(<0) 用 Single 位型重播种再走一步; Rnd() / Rnd(>0) 走一步
'   · 连投同一整数种子**不复现**序列 (RZ-5-A != RZ-5-B)
Private Sub DumpSix(ByVal label As String)
    Dim i As Long
    Dim s As String
    Dim v As Single
    s = label & "="
    For i = 1 To 6
        v = Rnd()
        s = s & " " & Hex$(CLng(v * 16777216#))
    Next i
    Debug.Print s
End Sub

Public Sub Main()
    Dim sv As Single
    Dim dv As Double
    Dim iv As Integer
    Dim lv As Long
    Dim vv As Variant

    ' 不带 Randomize: 初值 0x50000
    DumpSix "BASELINE"

    Randomize 0
    DumpSix "RZ-0"
    Randomize 1
    DumpSix "RZ-1"
    Randomize 2
    DumpSix "RZ-2"
    Randomize 5
    DumpSix "RZ-5"
    Randomize 42
    DumpSix "RZ-42"
    Randomize 100
    DumpSix "RZ-100"
    Randomize 42.5
    DumpSix "RZ-42.5"
    Randomize 42.25
    DumpSix "RZ-42.25"
    Randomize 0.5
    DumpSix "RZ-0.5"
    Randomize -1
    DumpSix "RZ-NEG1"
    Randomize -2
    DumpSix "RZ-NEG2"
    Randomize -1.5
    DumpSix "RZ-NEG1.5"
    vv = 42
    Randomize vv
    DumpSix "RZ-VAR42"
    Randomize 16777216
    DumpSix "RZ-2P24"
    Randomize 16777217
    DumpSix "RZ-2P24P1"

    ' 变量种子 (Single / Double / Integer / Long / Variant)
    sv = 42.5
    Randomize sv
    DumpSix "RZ-VAR-SGL"
    dv = 42.25
    Randomize dv
    DumpSix "RZ-VAR-DBL"
    iv = 7
    Randomize iv
    DumpSix "RZ-VAR-INT"
    lv = 42
    Randomize lv
    DumpSix "RZ-VAR-LNG"
    vv = 9.5
    Randomize vv
    DumpSix "RZ-VAR-VAR"

    ' Rnd 负种: 重播种 + 走一步; 同一负种连投得到同一串
    sv = Rnd(-1)
    DumpSix "RND-NEG1"
    sv = Rnd(-1)
    DumpSix "RND-NEG1-AGAIN"
    Debug.Print "RND-0=" & Hex$(CLng(Rnd(0) * 16777216#))
    Debug.Print "RND-POS=" & Hex$(CLng(Rnd(1) * 16777216#))
    Debug.Print "RND-BARE=" & Hex$(CLng(Rnd * 16777216#))

    ' 同一整数种子连投两次 → 不复现 (低字节沿用旧状态)
    Randomize 5
    DumpSix "RZ-5-A"
    Randomize 5
    DumpSix "RZ-5-B"

    ' 裸 Randomize (Timer 播种) 只能验可跑/值域, 不能钉序列
    Randomize
    sv = Rnd()
    Debug.Print "RZ-BARE-RANGE=" & ((sv >= 0!) And (sv < 1!))

    Debug.Print "RZ PASS"
End Sub
