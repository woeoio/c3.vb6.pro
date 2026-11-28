Option Explicit

' §11.232 (第二百五十刀): 定长串 (`As String * N`) 的各条赋值路径 —— 局部标量 / 模块级标量 /
' UDT 字段 / 定长串数组元素 (1 维、2 维) / LSet / Mid 语句 / 数·布尔·Empty·Null 右值。
' 夹具 = 真机探针 .temp/k233p/P361+P362+P363 的同一批语句（Print #f 换成 Debug.Print）。
' 真机 P363 C2 那一格 (`未赋值元素 = ""` → False) 刻意**不进夹具**: 它量的不是定长串的收口,
' 而是 `vb6_StrCmp` 的 wcscmp 口径 (内嵌 NUL 不参与比较)，属字符串比较那一族，留在背账。
Private Type TRec
    f As String * 5
End Type

Private g As String * 5
Private ga(1 To 2) As String * 4

Public Sub Main()
    Dim fx As String * 5
    Dim fy As String * 3
    Dim a(1 To 2) As String * 5
    Dim a0(1 To 2) As String * 5
    Dim b(1 To 2, 1 To 2) As String * 4
    Dim s As String
    Dim v As Variant
    Dim r As TRec

    On Error Resume Next

    ' 局部定长串：超长 / 短 / 空串 / 数 / 布尔
    fx = "abcdefgh"
    Debug.Print "A1|[" & fx & "]|len=" & CStr(Len(fx)) & "|num=" & CStr(Err.Number)
    Err.Clear
    fx = "ab"
    Debug.Print "A2|[" & fx & "]|len=" & CStr(Len(fx)) & "|cat=[" & (fx & "x") & "]"
    Err.Clear
    fx = ""
    Debug.Print "A3|[" & fx & "]|len=" & CStr(Len(fx))
    Err.Clear
    fx = 123
    Debug.Print "A4|[" & fx & "]|len=" & CStr(Len(fx)) & "|num=" & CStr(Err.Number)
    Err.Clear
    fx = 12.7
    Debug.Print "A5|[" & fx & "]|len=" & CStr(Len(fx))
    Err.Clear
    fx = True
    Debug.Print "A6|[" & fx & "]|len=" & CStr(Len(fx))
    Err.Clear

    ' Empty / Null 右值
    v = Empty
    fx = v
    Debug.Print "A7|[" & fx & "]|len=" & CStr(Len(fx)) & "|num=" & CStr(Err.Number)
    Err.Clear
    fx = "abc"
    v = Null
    fx = v
    Debug.Print "A8|[" & fx & "]|len=" & CStr(Len(fx)) & "|num=" & CStr(Err.Number)
    Err.Clear

    ' UDT 定长字段：超长 / 短 / Empty / Null
    r.f = "abcdefgh"
    Debug.Print "A9|[" & r.f & "]|len=" & CStr(Len(r.f)) & "|num=" & CStr(Err.Number)
    Err.Clear
    r.f = "ab"
    Debug.Print "A10|[" & r.f & "]|len=" & CStr(Len(r.f))
    Err.Clear
    v = Empty
    r.f = v
    Debug.Print "A11|[" & r.f & "]|len=" & CStr(Len(r.f)) & "|num=" & CStr(Err.Number)
    Err.Clear
    r.f = "abc"
    v = Null
    r.f = v
    Debug.Print "A12|[" & r.f & "]|len=" & CStr(Len(r.f)) & "|num=" & CStr(Err.Number)
    Err.Clear

    ' 定长串之间 / 定长串进动态串 / LSet / String$ / 拼接
    fx = "abcdefgh"
    fy = fx
    Debug.Print "A13|fy=[" & fy & "]|len=" & CStr(Len(fy))
    Err.Clear
    s = fx
    Debug.Print "A14|len=" & CStr(Len(s)) & "|[" & s & "]"
    Err.Clear
    fx = "x"
    LSet fx = "abcdefgh"
    Debug.Print "A15|[" & fx & "]|len=" & CStr(Len(fx)) & "|num=" & CStr(Err.Number)
    Err.Clear
    fx = String(10, "Z")
    Debug.Print "A16|[" & fx & "]|len=" & CStr(Len(fx))
    Err.Clear
    fx = "ab" & "cdefgh"
    Debug.Print "A17|[" & fx & "]|len=" & CStr(Len(fx))
    Err.Clear

    ' 模块级定长串
    g = "abcdefgh"
    Debug.Print "A18|[" & g & "]|len=" & CStr(Len(g)) & "|num=" & CStr(Err.Number)
    Err.Clear
    g = "ab"
    Debug.Print "A19|[" & g & "]|len=" & CStr(Len(g))
    Err.Clear
    g = ""
    Debug.Print "A20|[" & g & "]|len=" & CStr(Len(g))
    Err.Clear
    v = Empty
    g = v
    Debug.Print "A21|[" & g & "]|len=" & CStr(Len(g)) & "|num=" & CStr(Err.Number)
    Err.Clear
    g = "abc"
    v = Null
    g = v
    Debug.Print "A22|[" & g & "]|len=" & CStr(Len(g)) & "|num=" & CStr(Err.Number)
    Err.Clear
    g = 123
    Debug.Print "A23|[" & g & "]|len=" & CStr(Len(g))
    Err.Clear

    ' 模块级 Mid / LSet
    g = "abc"
    Mid(g, 2) = "z"
    Debug.Print "A24|[" & g & "]|len=" & CStr(Len(g)) & "|num=" & CStr(Err.Number)
    Err.Clear
    g = "x"
    LSet g = "abcdefgh"
    Debug.Print "A25|[" & g & "]|len=" & CStr(Len(g))
    Err.Clear

    ' 一维定长串数组元素
    a(1) = "abcdefgh"
    Debug.Print "A26|[" & a(1) & "]|len=" & CStr(Len(a(1))) & "|num=" & CStr(Err.Number)
    Err.Clear
    a(2) = "ab"
    Debug.Print "A27|[" & a(2) & "]|len=" & CStr(Len(a(2)))
    Err.Clear

    ' 未赋值元素：N 个 NUL（不是空串，也不是空格）
    Debug.Print "A28|len=" & CStr(Len(a0(1))) & "|asc=" & CStr(Asc(a0(1))) & "|vt=" & CStr(VarType(a0(1)))
    Debug.Print "A29|isempty=" & CStr(IsEmpty(a0(1))) & "|catlen=" & CStr(Len(a0(1) & ""))
    Err.Clear

    ' 数组元素收 Variant：Null / Empty
    a(1) = "abc"
    v = Null
    a(1) = v
    Debug.Print "A30|[" & a(1) & "]|len=" & CStr(Len(a(1))) & "|num=" & CStr(Err.Number)
    Err.Clear
    v = Empty
    a(1) = v
    Debug.Print "A31|[" & a(1) & "]|len=" & CStr(Len(a(1))) & "|num=" & CStr(Err.Number)
    Err.Clear

    ' 模块级定长串数组
    ga(1) = "abcdefgh"
    Debug.Print "A32|[" & ga(1) & "]|len=" & CStr(Len(ga(1))) & "|num=" & CStr(Err.Number)
    Err.Clear
    ga(2) = "ab"
    Debug.Print "A33|[" & ga(2) & "]|len=" & CStr(Len(ga(2)))
    Err.Clear

    ' 二维定长串数组
    b(1, 2) = "abcdefgh"
    Debug.Print "A34|[" & b(1, 2) & "]|len=" & CStr(Len(b(1, 2))) & "|num=" & CStr(Err.Number)
    Err.Clear
    b(2, 1) = "ab"
    Debug.Print "A35|[" & b(2, 1) & "]|len=" & CStr(Len(b(2, 1)))
    Err.Clear

    ' 边界不变
    Debug.Print "A36|lb=" & CStr(LBound(a)) & "|ub=" & CStr(UBound(a)) & "|lb2=" & CStr(LBound(b, 2)) & "|ub2=" & CStr(UBound(b, 2))
End Sub