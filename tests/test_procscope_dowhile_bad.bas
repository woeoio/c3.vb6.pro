Option Explicit

' 负控（调查里的 D1）：`Do While` 是**前置**条件，文本上先于循环体。
' 条件里引用体内才声明的 `stopFlag` 必须报未声明 —— 真 VB6 报「变量未定义」。
' 若这里静默通过，说明 visit(DoLoopStmt) 又被改成了无条件「先体后条件」
' （那会把 VB6 编不过的写法放进来）。

Public Sub Main()
    Dim i As Long
    Do While Not stopFlag
        Dim stopFlag As Boolean
        stopFlag = True
        i = i + 1
    Loop
    Debug.Print "DH=" & i
End Sub
