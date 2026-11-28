Option Explicit

' §11.229 (第二百四十七刀): `On Error GoTo 0` 与 Err 的关系。
' 夹具 = 真机探针 .temp/k233p/P357 的同一份源码（只把 Print #f 换成 Debug.Print），
' 9 格逐字对照：GoTo 0 之后 Err.Number/Description/Source 全零；Err.Clear 作对照。
Public Sub Main()
    On Error Resume Next
    Err.Raise 5
    Debug.Print "A1|num=" & CStr(Err.Number) & "|[" & Err.Description & "]|src=" & Err.Source
    On Error GoTo 0
    Debug.Print "A2|num=" & CStr(Err.Number) & "|[" & Err.Description & "]|src=" & Err.Source
    ' 没发生过错误时再执行一次
    On Error GoTo 0
    Debug.Print "A3|num=" & CStr(Err.Number)
    ' GoTo 0 之后重新开 Resume Next 再报一个
    On Error Resume Next
    Err.Raise 6
    Debug.Print "A4|num=" & CStr(Err.Number)
    On Error GoTo 0
    Debug.Print "A5|num=" & CStr(Err.Number)
    ' 对照：Err.Clear
    On Error Resume Next
    Err.Raise 7
    Debug.Print "A6|num=" & CStr(Err.Number)
    Err.Clear
    Debug.Print "A7|num=" & CStr(Err.Number)
    ' 带标签的处理程序里，Err 在 handler 里还在
    WithHandler
End Sub

Private Sub WithHandler()
    On Error GoTo H
    Err.Raise 5
    Debug.Print "B1|num=" & CStr(Err.Number)
    Exit Sub
H:
    Debug.Print "B2|num=" & CStr(Err.Number) & "|[" & Err.Description & "]"
    Resume Next
End Sub