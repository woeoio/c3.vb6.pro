Option Explicit

' §11.226 (第二百四十四刀): `+` 与 Null/Empty 的口径 —— 与 `&` 相反，`+` 是 Null 传播。
' 判据是真机 .temp/k233p/P355（VB6Mini，Print #f 落文件），夹具与孪生逐格对照（同一份源码，
' 只把 Print #f 换成 Debug.Print）。
'   ① 任一侧 Null ⇒ 结果 Null：String/Long 目标 94 且目标不变，Variant 目标 IsNull=True；
'   ② Empty 当 0 / 空串（"x" + Empty ⇒ "x"、1 + Empty ⇒ 1、Empty + Empty ⇒ 0）；
'   ③ 两侧都是 Null 也是 Null（与 `&` 的「两侧都 Null 才 Null」不同）。
' E13–E17 是 String+数值那一族（113g），本刀不动，读数留在 .temp/k233p/p355full.txt。
Public Sub Main()
    Dim s As String, v As Variant, w As Variant, n As Long
    On Error Resume Next

    ' ① 一侧 Null：String 目标
    Err.Clear
    s = "x" + Null
    Debug.Print "E1|num=" & CStr(Err.Number) & "|[" & Err.Description & "]|s=[" & s & "]"
    Err.Clear
    s = Null + ""
    Debug.Print "E2|num=" & CStr(Err.Number) & "|s=[" & s & "]"
    ' ② 一侧 Null：Variant 目标
    Err.Clear
    v = "x" + Null
    Debug.Print "E3|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v))
    Err.Clear
    v = 1 + Null
    Debug.Print "E4|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v))
    Err.Clear
    v = Null + Null
    Debug.Print "E5|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v))
    Err.Clear
    v = Null + ""
    Debug.Print "E6|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v))
    Err.Clear
    v = Null + 1
    Debug.Print "E7|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v))
    ' ③ Empty
    Err.Clear
    v = "x" + Empty
    Debug.Print "E8|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v)) & "|[" & CStr(v) & "]"
    Err.Clear
    v = 1 + Empty
    Debug.Print "E9|num=" & CStr(Err.Number) & "|[" & CStr(v) & "]"
    Err.Clear
    v = Empty + Empty
    Debug.Print "E10|num=" & CStr(Err.Number) & "|[" & CStr(v) & "]"
    Err.Clear
    n = 1 + Empty
    Debug.Print "E11|num=" & CStr(Err.Number) & "|n=" & CStr(n)
    Err.Clear
    s = "x" + Empty
    Debug.Print "E12|num=" & CStr(Err.Number) & "|s=[" & s & "]"
    ' ④ 变体槽两侧
    Err.Clear
    v = Null
    w = 1
    n = v + w
    Debug.Print "E18|num=" & CStr(Err.Number) & "|n=" & CStr(n)
    Err.Clear
    v = Null
    w = Null
    s = v + w
    Debug.Print "E19|num=" & CStr(Err.Number) & "|[" & Err.Description & "]"
    Err.Clear
    v = Empty
    w = 2
    n = v + w
    Debug.Print "E20|num=" & CStr(Err.Number) & "|n=" & CStr(n)
End Sub