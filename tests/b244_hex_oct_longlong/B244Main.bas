Option Explicit

' §11.220 (第二百三十七刀): Hex/Oct 的**第四阶** —— 64 位载体（LongLong / ULongLong）。
' VB6 没有 LongLong 类型，所以这一阶的真机读数不存在；判据是 twinBASIC（用户 2026-11-28 指定），
' 读数 .temp/k228p/tbhex（0.15.1047，Win64）的 L1..L14。
' 规则与前三阶同一条递推：位数 = 载体的存储宽度（Byte 2 / Integer 4 / Long 8 / LongLong 16 位十六进制，
' 八进制 3 / 6 / 11 / 22）；正值不补零，负值掩到该宽度的二补码。
'
' 改前 C3 把 int64_t 递进 vb6_Hex(int32_t) 的形参，低 32 位静默截断：
'   L1 = FFFFFFFF（少了 8 位）、L5 2^32 = 0、L6 = 0、L11/L12（变体槽装 I8）同样是截断值。
Public Sub Main()
    Dim q As LongLong
    q = -1
    Debug.Print "L1=" & Hex(q) & "/" & Oct(q) & "/len=" & Len(Hex(q))
    q = 0
    Debug.Print "L2=" & Hex(q) & "/" & Oct(q) & "/len=" & Len(Hex(q))
    q = 16
    Debug.Print "L3=" & Hex(q) & "/" & Oct(q) & "/len=" & Len(Hex(q))
    q = 2147483648
    Debug.Print "L4=" & Hex(q) & "/" & Oct(q) & "/len=" & Len(Hex(q))
    q = 4294967296
    Debug.Print "L5=" & Hex(q) & "/" & Oct(q) & "/len=" & Len(Hex(q))
    q = -4294967296
    Debug.Print "L6=" & Hex(q) & "/" & Oct(q) & "/len=" & Len(Hex(q))

    ' 前三阶一字不动（真机 VB6 .temp/k228p/p332.txt 的 C1/C9/C10 同族）
    Debug.Print "L9=" & Hex(CLng(-1)) & "/" & Oct(CLng(-1))
    Debug.Print "L10=" & Hex(CInt(-1)) & "/" & Oct(CInt(-1))

    ' 变体槽装着 LongLong 的那一面：宽度在槽的 vt 上，不在实参文本上
    Dim v As Variant
    v = CLngLng(-1)
    Debug.Print "L11=" & Hex(v) & "/" & Oct(v)
    v = CLngLng(4294967296)
    Debug.Print "L12=" & Hex(v) & "/" & Oct(v)

    ' 直接内建调用（CLngLng 的返回值就是 LongLong）
    Debug.Print "L13=" & Hex(CLngLng(-1)) & "/" & Oct(CLngLng(-1))
    Debug.Print "L14=" & Hex(CLngLng(305419896)) & "/" & Oct(CLngLng(305419896))

    ' 无符号 64 位那一档：按位同值（高位看得见），与 LongLong 共用一枚出口
    Dim u As ULongLong
    u = 4294967296
    Debug.Print "L15=" & Hex(u) & "/" & Oct(u)
    u = 18446744073709551615
    Debug.Print "L16=" & Hex(u) & "/" & Oct(u)
End Sub