Option Explicit

' 负控：`While ... Wend` 与 `If ... Then` 的条件都文本上先于体，
' 条件里引用体内才声明的 `Const` / `Dim` 必须报未声明。
' 真 VB6 在这三处都报「变量未定义」；C3 方向要一致（只钉"有没有报"，
' 等级 warning/error 的差异不在本例判据内）。

Public Sub Main()
    Dim i As Long
    While i < LIMIT_WEND
        Const LIMIT_WEND As Long = 3
        i = i + 1
    Wend

    Dim j As Long
    If j = ON_VAL_IF Then
        Const ON_VAL_IF As Long = 1
        j = 1
    End If

    Dim k As Long
    Do While Not DONE_FLAG
        Dim DONE_FLAG As Boolean
        DONE_FLAG = True
        k = k + 1
    Loop

    Debug.Print "CR=" & i & "/" & j & "/" & k
End Sub
