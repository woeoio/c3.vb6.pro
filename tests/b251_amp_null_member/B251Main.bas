Option Explicit

Private Type TU
    V As Variant
End Type

Private arr(0 To 1) As Variant

' §11.228 (第二百四十六刀): `&` 里 Null 的**数组元素 / UDT 成员 / With 成员**。
' 判据是真机 .temp/k233p/P358（VB6Mini，Print #f 落文件），夹具与孪生逐格对照
' （同一份源码，只把 Print #f 换成 Debug.Print）。
' 口径与 §11.225 一致：一侧 Null 当空串、两侧都 Null ⇒ 结果 Null（String 目标 94）。
Public Sub Main()
    Dim s As String, v As Variant, u As TU
    On Error Resume Next

    arr(0) = Null
    arr(1) = Null
    s = "x" & arr(0)
    Debug.Print "F1|num=" & CStr(Err.Number) & "|[" & s & "]"
    s = arr(0) & "y"
    Debug.Print "F2|num=" & CStr(Err.Number) & "|[" & s & "]"
    u.V = Null
    s = "x" & u.V
    Debug.Print "F3|num=" & CStr(Err.Number) & "|[" & s & "]"
    With u
        s = "x" & .V
    End With
    Debug.Print "F4|num=" & CStr(Err.Number) & "|[" & s & "]"
    v = arr(0) & arr(1)
    Debug.Print "F5|num=" & CStr(Err.Number) & "|isnull=" & CStr(IsNull(v))
    s = arr(0) & arr(1)
    Debug.Print "F6|num=" & CStr(Err.Number) & "|[" & Err.Description & "]"
    ' 非 Null 的对照（这些格不该变）
    Err.Clear
    arr(0) = "a"
    u.V = "b"
    s = "x" & arr(0)
    Debug.Print "F7|num=" & CStr(Err.Number) & "|[" & s & "]"
    s = "x" & u.V
    Debug.Print "F8|num=" & CStr(Err.Number) & "|[" & s & "]"
End Sub