Option Explicit

' §11.225 (第二百四十三刀): `&` 与 Null/Empty 的口径。
' 判据是真机 .temp/k233p/P350/P352/P353（VB6Mini），夹具与孪生逐格对照。
'   ① 一侧 Null ⇒ 当空串（err 0）；② Empty ⇒ 空串；
'   ③ 两侧都 Null ⇒ 结果 Null（String 目标 94、Variant 目标 IsNull=True）；
'   ④ 其余照 CStr（数字/布尔）。
Public Sub Main()
    Dim s As String, v As Variant, w As Variant

    ' ① 一侧 Null 当空串
    s = "x" & Null
    Debug.Print "D1=" & CStr(Err.Number) & "/[" & s & "]"
    s = "x" & Null & "y"
    Debug.Print "D2=" & CStr(Err.Number) & "/[" & s & "]"
    ' ② Empty 当空串
    s = "x" & Empty
    Debug.Print "D3=" & CStr(Err.Number) & "/[" & s & "]"
    s = Empty & Empty
    Debug.Print "D4=" & CStr(Err.Number) & "/[" & s & "]"
    ' ③ 两侧都 Null
    v = Null & Null
    Debug.Print "D5=" & CStr(Err.Number) & "/" & CStr(IsNull(v))
    v = "x" & Null
    Debug.Print "D6=" & CStr(Err.Number) & "/" & CStr(IsNull(v)) & "/" & CStr(VarType(v)) & "/[" & CStr(v) & "]"
    v = Null & ""
    Debug.Print "D7=" & CStr(Err.Number) & "/" & CStr(IsNull(v))
    On Error Resume Next
    Err.Clear
    s = Null & Null
    Debug.Print "D8=" & CStr(Err.Number) & "/[" & Err.Description & "]"
    Err.Clear
    v = Null
    w = Null
    s = v & w
    Debug.Print "D9=" & CStr(Err.Number) & "/[" & Err.Description & "]"
    On Error GoTo 0
    ' 真机在 `On Error GoTo 0` 时会把 Err 清掉；C3 还没这一格（背账），这里显式清，
    ' 免得后面每行都把 94 带下去 —— 本夹具钉的是 `&` 的口径。
    Err.Clear
    ' ④ 数字/布尔照 CStr
    s = 1 & Null
    Debug.Print "D10=" & CStr(Err.Number) & "/[" & s & "]"
    s = "x" & 1.5
    Debug.Print "D11=" & CStr(Err.Number) & "/[" & s & "]"
    s = "x" & True
    Debug.Print "D12=" & CStr(Err.Number) & "/[" & s & "]"
    v = 3.5 & Null
    Debug.Print "D13=" & CStr(Err.Number) & "/" & CStr(IsNull(v)) & "/[" & CStr(v) & "]"
    v = True & Null
    Debug.Print "D14=" & CStr(Err.Number) & "/" & CStr(IsNull(v)) & "/[" & CStr(v) & "]"
    ' 变体槽里装 Null 的那一侧
    v = Null
    s = v & "x"
    Debug.Print "D15=" & CStr(Err.Number) & "/[" & s & "]"
    s = "x" & v
    Debug.Print "D16=" & CStr(Err.Number) & "/[" & s & "]"
    s = Len(v & "")
    Debug.Print "D17=" & CStr(Err.Number) & "/" & CStr(s)
    ' 变体槽装空串不是 Null
    v = ""
    s = "x" & v
    Debug.Print "D18=" & CStr(Err.Number) & "/[" & s & "]"
    ' 链里的中间结果（两侧都 Null 的那壳折回串）
    v = Null
    w = Null
    Debug.Print "D19=" & CStr(Err.Number) & "/[" & ("a" & (v & w) & "b") & "]"
    ' 拼接链一步步传下去：三步全 Null ⇒ Null；夹一个串 ⇒ 串
    v = Null
    w = Null
    v = (v & w & w)
    Debug.Print "D20=" & CStr(Err.Number) & "/" & CStr(IsNull(v))
    v = Null
    v = (v & w & "x")
    Debug.Print "D21=" & CStr(Err.Number) & "/" & CStr(IsNull(v)) & "/[" & CStr(v) & "]"
End Sub