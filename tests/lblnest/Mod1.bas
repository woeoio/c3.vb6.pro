Attribute VB_Name = "Mod1"
Option Explicit

' C2045 复现 (PhotoDemon 整编 50 条): For 方向拆分是**逐层**的, 嵌套时内层会把外层的
' 副本位覆盖/清零 ⇒ 同一标签在不同副本组合里拼出同一个 _d2 (label redefined),
' 外层循环体里内层 For 之后的标签更是两份都不带后缀。
Public Sub TestLoop()
    Dim i As Long, j As Long, n As Long
    For i = 1 To 10
        For j = 1 To 10
            If j = 5 Then GoTo SkipPixel
            n = n + 1
SkipPixel:
        Next j
        If i = 3 Then GoTo NextPixel
        n = n + 1
NextPixel:
    Next i
End Sub
