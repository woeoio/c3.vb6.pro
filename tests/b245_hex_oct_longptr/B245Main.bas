Option Explicit

' §11.221 (第二百三十八刀): Hex/Oct 的**第五阶** —— 指针宽载体（LongPtr）。
' VB6 没有 LongPtr，判据仍由 twinBASIC 台出（.temp/k229p/tbptr, 0.15.1047），**两个架构各量一遍**：
'   win64 (LenB=8): Hex(LongPtr(-1)) = 16 位 / Oct 22 位
'   win32 (LenB=4): Hex(LongPtr(-1)) = 8 位  / Oct 11 位
' 规则与第四阶同一条：位数 = 载体的存储宽度，而 LongPtr 的宽度就是 intptr_t 的宽度。
' 所以本夹具的 P1/P4/P7 三格随架构不同（run_tests 里两套针），其余各格两架构同值。
'
' 改前 x64 上 `Hex(p)` 落进 vb6_Hex(int32_t)：指针 8 字节却只打 8 位（P1 = FFFFFFFF）。
Public Sub Main()
    Dim p As LongPtr
    p = -1
    Debug.Print "P1=" & Hex(p) & "/" & Oct(p) & "/lenb=" & LenB(p)

    p = 16
    Debug.Print "P2=" & Hex(p) & "/" & Oct(p)

    p = CLngPtr(305419896)
    Debug.Print "P3=" & Hex(p) & "/" & Oct(p)

    Debug.Print "P4=" & Hex(CLngPtr(-1)) & "/" & Oct(CLngPtr(-1))

    ' 对照：第四阶（LongLong）不随架构变
    Debug.Print "P5=" & Hex(CLngLng(-1)) & "/len=" & Len(Hex(CLngLng(-1)))

    ' 槽那一面：槽的 vt 决定宽度（x64 装 VT_I8、x86 装 VT_I4），与实参文本无关
    Dim v As Variant
    v = p
    Debug.Print "P6=" & Hex(v) & "/" & Oct(v)
    v = CLngPtr(-1)
    Debug.Print "P7=" & Hex(v) & "/" & Oct(v)

    ' 前三阶不动（真机 VB6 口径）
    Debug.Print "P8=" & Hex(CLng(-1)) & "/" & Oct(CLng(-1))
    Debug.Print "P9=" & Hex(CInt(-1)) & "/" & Oct(CInt(-1))
End Sub