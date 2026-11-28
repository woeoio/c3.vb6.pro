Option Explicit

' §11.223 (第二百四十一刀): `Error(n)` 函数面。
' 判据是真机 .temp/k231p/P340/P343/P345/P346（VB6Mini）。三条与 `Err.Raise` **不同**的规则：
'   1) `Error(0)` 是 vbNullString（StrPtr=0），不是空 BSTR，也不是上一格描述；
'   2) 负数号一律兜底句 —— 不是 `Err.Raise -1` 那枚 "Automation error"（P340 V4）；
'   3) 号门在另一处：`Error(65536)` 报 **6 Overflow**，而 `Err.Raise 65536` 报 5（P345 Z5）。
' 参数按 VB6 的 CLng 语义收：11.6→12、11.4→11、0.5→0、"11"→11、"abc"/""→13、Null→94、Empty→0。
' 注意 `Error(32768)` / `Error(40000)` **不报错**（给兜底句），别把它当成 Integer 越界。
Public Sub Main()
    On Error Resume Next

    Debug.Print "E1=[" & Error(11) & "]"
    Debug.Print "E2=[" & Error(9999) & "]"
    Debug.Print "E3=[" & Error(0) & "]/len=" & CStr(Len(Error(0)))
    Debug.Print "E4=[" & Error(-1) & "]"
    Debug.Print "E5=[" & Error(13) & "]"

    ' 上界：65536 报 6；40000 不报
    Err.Clear
    Dim s As String
    s = Error(65536)
    Debug.Print "E6=num=" & CStr(Err.Number) & "/[" & Err.Description & "]/s=[" & s & "]"

    Err.Clear
    Debug.Print "E7=[" & Error(40000) & "]/num=" & CStr(Err.Number)

    ' 字符串与 Null 的到达方式
    Err.Clear
    s = Error("abc")
    Debug.Print "E8=num=" & CStr(Err.Number) & "/[" & Err.Description & "]"

    Err.Clear
    Dim v As Variant
    v = Null
    s = Error(v)
    Debug.Print "E9=num=" & CStr(Err.Number) & "/[" & Err.Description & "]"

    ' 舍入：11.6→12（兜底），11.4→11（Division by zero）
    Err.Clear
    Debug.Print "E10=[" & Error(11.6) & "]/[" & Error(11.4) & "]/num=" & CStr(Err.Number)

    ' 裸 Error 仍是当前那一格描述
    Err.Clear
    Err.Raise 5
    Debug.Print "E11=[" & Error & "]/[" & Error(5) & "]"

    ' Empty → 0 → 空；数字 13 的描述正好是 "Type mismatch"
    Err.Clear
    v = Empty
    Debug.Print "E12=[" & Error(v) & "]/num=" & CStr(Err.Number) & "/[" & Error(13) & "]"
End Sub