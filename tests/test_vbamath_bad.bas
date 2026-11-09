Option Explicit

' 负控 (test_vbamath): 只放行**限定符位**的 `Math.<VBA.Math 成员>`。
' 值位的裸 `Math` 不是 VB6 里的合法名字，必须照旧报 VB3001。

Public Sub Main()
    Dim m As Object

    Set m = Math
    Debug.Print "VM-BAD PASS"
End Sub
