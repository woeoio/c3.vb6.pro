Option Explicit

' §11.231 (第二百四十九刀): Mid 语句余下的边角 —— 定长串目标 / Variant Empty / Null 替换串 /
' 显式长度与替换串长度不一致 / Mid$ 写法。
' 夹具 = 真机探针 .temp/k233p/P360 的同一份源码（只把 Print #f 换成 Debug.Print），9 格逐字对照。
Public Sub Main()
    Dim s As String
    Dim fx As String * 5
    Dim v As Variant
    Dim nv As Variant

    On Error Resume Next

    ' 定长串目标
    fx = "abcde"
    Mid(fx, 2) = "z"
    Debug.Print "F2|num=" & CStr(Err.Number) & "|[" & fx & "]|len=" & CStr(Len(fx))
    Err.Clear
    fx = "abcde"
    Mid(fx, 1) = ""
    Debug.Print "F3|num=" & CStr(Err.Number) & "|[" & fx & "]"
    Err.Clear
    fx = "abc"
    Mid(fx, 4) = "z"
    Debug.Print "F4|num=" & CStr(Err.Number) & "|[" & fx & "]"

    ' Variant Empty 目标
    Err.Clear
    v = Empty
    Mid(v, 1) = "z"
    Debug.Print "F5|num=" & CStr(Err.Number) & "|isempty=" & CStr(IsEmpty(v)) & "|vt=" & CStr(VarType(v))

    ' Null 替换串
    Err.Clear
    s = "abc"
    nv = Null
    Mid(s, 1) = nv
    Debug.Print "F6|num=" & CStr(Err.Number) & "|[" & s & "]"

    ' 显式长度 vs 替换串长度
    Err.Clear
    s = "abc"
    Mid(s, 2, 2) = ""
    Debug.Print "F7|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 3) = "xy"
    Debug.Print "F8|num=" & CStr(Err.Number) & "|[" & s & "]"
    Err.Clear
    s = "abc"
    Mid(s, 2, 1) = "xyz"
    Debug.Print "F9|num=" & CStr(Err.Number) & "|[" & s & "]"

    ' Mid$ 语句写法
    Err.Clear
    s = "abc"
    Mid$(s, 2) = "z"
    Debug.Print "F10|num=" & CStr(Err.Number) & "|[" & s & "]"
End Sub