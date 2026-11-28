Option Explicit

' §11.230 (第二百四十八刀): Mid 语句的边角 —— 零长串 / 越界起点 / 长度 0 / 负参 / Variant 目标。
' 夹具 = 真机探针 .temp/k233p/P359 的同一份源码（只把 Print #f 换成 Debug.Print），14 格逐字对照。
Public Sub Main()
    Dim s As String

    On Error Resume Next

    ' 零长串目标
    s = ""
    Mid(s, 1) = ""
    Debug.Print "M1|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    Mid(s, 1) = "x"
    Debug.Print "M2|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    Mid(s, 1, 0) = "x"
    Debug.Print "M3|num=" & CStr(Err.Number) & "|[" & s & "]"

    ' 常规串：空替换串 / 长度 0 / 越界起点
    Err.Clear
    s = "abc"
    Mid(s, 2) = ""
    Debug.Print "M4|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 2, 0) = "z"
    Debug.Print "M5|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 0) = "z"
    Debug.Print "M6|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 5) = "z"
    Debug.Print "M7|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 4) = ""
    Debug.Print "M8|num=" & CStr(Err.Number) & "|[" & s & "]"

    ' 常规替换（对照）
    Err.Clear
    s = "abc"
    Mid(s, 2) = "z"
    Debug.Print "M9|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 2, 5) = "z"
    Debug.Print "M10|num=" & CStr(Err.Number) & "|[" & s & "]"

    ' 负参 / 零长且带长度
    Err.Clear
    s = "abc"
    Mid(s, -1) = "z"
    Debug.Print "M11|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 2, -1) = "z"
    Debug.Print "M12|num=" & CStr(Err.Number) & "|[" & s & "]"

    ' Variant / Null 目标
    Err.Clear
    Dim v As Variant
    v = Null
    Mid(v, 1) = "z"
    Debug.Print "M13|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v)) & "|vt=" & CStr(VarType(v))
    Err.Clear
    v = ""
    Mid(v, 1) = ""
    Debug.Print "M14|num=" & CStr(Err.Number) & "|[" & CStr(v) & "]|vt=" & CStr(VarType(v))
End Sub