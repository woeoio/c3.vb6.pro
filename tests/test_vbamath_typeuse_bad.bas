Option Explicit

' VBA 标准模块出现在 As 位置的负控（PhotoDemon Math 定案，handoff §10.4）。
' 真 VB6 对 `Dim m As VBA.Math` / `Dim m As Math` 报「Visual Basic 中不支持自动化类型」
' （标准模块不是类，也不是对象）。C3 必须在语义层报 VB3046 并停止发码
' （修前静默当 Variant 编过）。

Public Sub Main()
    Dim mLib  As VBA.Math
    Dim mBare As Math
    Debug.Print "VMBAD-TYPE PASS"
End Sub
