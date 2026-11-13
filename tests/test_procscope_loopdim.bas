Option Explicit

' 过程级作用域：VB6 里 `Dim` / `Const` 写在哪个块里都算**整个过程**的声明，
' 因此 Do 体里声明的量可以在 `Loop While` 条件里用（PhotoDemon pdXCF.cls 的
' ldResult / ERR_BUFF_TOO_SMALL、UserControl_Support.bas 的 positionFailed /
' failsafePosition 都是这一形）。修前语义层先分析条件、后分析体 ⇒ 假 VB3001。

Public Sub Main()
    Dim i As Long
    Dim result As Long

    i = 0
    Do
        Dim inner As Long
        Const TOO_SMALL As Long = 3

        i = i + 1
        inner = i * 2
        result = inner
    Loop While ((result <> TOO_SMALL * 2) Or (inner < 2)) And (i < 5)

    Debug.Print "PS-LOOP=" & i & "," & result & "," & inner
End Sub
