Option Explicit

' §11.236 (第二百五十四刀): Like / Is 的优先级 = 比较档（不是比算术还紧），
' 以及非字符串操作数的 Like 按 CStr 折串。夹具 = 真机 .temp/k236p/P373 + P374 + P375。
Public Sub Main()
    Dim z As String
    z = Chr(0)

    ' --- P373: Like 与 = / & 的相对位置
    Debug.Print "Y1|" & CStr(("a" = "a") Like "a")
    Debug.Print "Y2|" & CStr("a" Like "a" = True)
    Debug.Print "Y3|" & CStr("ab" Like "a" & "b")
    Debug.Print "Y4|" & CStr(VarType("x" Like "x" And True))
    Debug.Print "Y5|" & CStr(1 = 1 Like "True")
    Debug.Print "Y6|" & CStr("a" Like "a" = True)
    Debug.Print "Y7|" & CStr((1 = 1) Like "True")
    Debug.Print "Y8|" & CStr(VarType("ab" Like "a" & "b"))
    Debug.Print "Y9|" & CStr(VarType(("a" & z) Like "a" & z))

    ' --- P374: Is 与 = 的相对位置
    Dim o As Object
    Dim p As Object

    On Error Resume Next
    Debug.Print "Z1|" & CStr(o Is Nothing = True)
    If Err.Number <> 0 Then Debug.Print "Z1err|" & CStr(Err.Number)
    Err.Clear
    Debug.Print "Z2|" & CStr(VarType(o Is Nothing))
    If Err.Number <> 0 Then Debug.Print "Z2err|" & CStr(Err.Number)
    Err.Clear
    Debug.Print "Z3|" & CStr(VarType(o Is Nothing And True))
    If Err.Number <> 0 Then Debug.Print "Z3err|" & CStr(Err.Number)
    Err.Clear
    Set p = o
    Debug.Print "Z4|" & CStr(p Is o And True)
    If Err.Number <> 0 Then Debug.Print "Z4err|" & CStr(Err.Number)
    On Error GoTo 0

    ' --- P375: 非字符串操作数按 CStr 折串
    Debug.Print "N1|" & CStr(1 Like "1")
    Debug.Print "N2|" & CStr(True Like "True")
    Debug.Print "N3|" & CStr(2.5 Like "2.5")
    Debug.Print "N4|" & CStr("abc" Like 1)
    Debug.Print "N5|" & CStr("1" Like 1)
    Debug.Print "N6|" & CStr(1 Like 1)
    Debug.Print "N7|" & CStr(False Like "False")
    Debug.Print "N8|" & CStr(VarType(1 Like "1"))
End Sub