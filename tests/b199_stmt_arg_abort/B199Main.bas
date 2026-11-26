Option Explicit

' 账 #305 §11.174 (第一百九十一刀): 语句位置调用的实参中止面。
'
' 改前 (病): 实参守卫只在「实参含晚绑定 COM 读」时开闸 (第一百零二刀 的 I 族判据),
' 普通过程调用 / 内建 / 除零当暴雷源时压根不开闸 —— 于是 `Take2 FEscape(), LogTag(2)`
' 在 C3 里两个实参都求值、被调方也进去了 (真机: 一个都不该发生)。
'
' 判据以真机为准 (.temp/k191p/P217.bas 的 A7/A8 两格, 读数 p217.txt):
'   A7 `Take2 FEscape(), LogTag(2)` → hits=0: 坏实参之后的实参不求值, 被调方不进入;
'   A8 `Take2 LogTag(1), FEscape()` → hits=1: 坏实参之前的实参照常求值**恰好一次**。
'
' 这里钉的四格不重复探针, 只钉这一刀真正碰到的两个面:
'   B1/B2  上述两条口径 (改前分别是 1002 / 1001);
'   B3     无 raise 时 ByRef 写回照旧 —— 提升临时量的实现会把写回丢到 _agN 上,
'          所以 ByRef 实参一律不提升, 这一格就是那条纪律的哨兵;
'   B4     ByRef 在前、坏实参在后: 整句中止, 目标不动。
'
' 没钉的几格 (A1/A5/A9/A10/A11/A12) 属后面两刀: 赋值站开闸、被调方体内中止、
' Resume Next 的过程级作用域, 都还红着, 现在钉不进来。

Private dv As Long        ' 0 ⇒ `100 \ dv` 运行期 raise 11
Private hits As Long      ' 求值痕迹: 实参 ×10 + 序号, 被调方 +1000

' 无处理器 ⇒ 错误上抛
Private Function FEscape() As Long
    Dim z As Long
    z = 100 \ dv
    FEscape = 9
End Function

' 有副作用、会改动 hits 的普通实参
Private Function LogTag(ByVal n As Long) As Long
    hits = hits * 10 + n
    LogTag = n
End Function

' 两个 ByVal 形参: 被调方一旦进入就留痕
Private Sub Take2(ByVal a As Long, ByVal b As Long)
    hits = hits + 1000
    Print #1, "    NOTE take2-ran"
End Sub

' ByRef 形参: 写回必须落到调用方的变量上
Private Sub Bump(ByRef n As Long, ByVal t As Long)
    n = n + t
End Sub

Public Sub Main()
    Open App.Path & "\b199.txt" For Output As #1
    dv = 0
    Dim x As Long
    Dim e As Long

    On Error Resume Next

    ' B1: 坏实参在前 ⇒ 之后的实参不求值, 被调方不进入
    hits = 0: Err.Clear
    Take2 FEscape(), LogTag(2)
    e = Err.Number: Err.Clear
    Print #1, "B1=[" & hits & "," & e & "]"

    ' B2: 坏实参在后 ⇒ 之前的实参照常求值恰好一次
    hits = 0: Err.Clear
    Take2 LogTag(1), FEscape()
    e = Err.Number: Err.Clear
    Print #1, "B2=[" & hits & "," & e & "]"

    ' B3: 无 raise ⇒ ByRef 写回照旧 (提升临时量不许吃掉写回)
    x = 5: hits = 0: Err.Clear
    Bump x, LogTag(3)
    e = Err.Number: Err.Clear
    Print #1, "B3=[" & x & "," & hits & "," & e & "]"

    ' B4: ByRef 在前、坏实参在后 ⇒ 整句中止, 目标不动
    x = 5: hits = 0: Err.Clear
    Bump x, FEscape()
    e = Err.Number: Err.Clear
    Print #1, "B4=[" & x & "," & hits & "," & e & "]"

    Print #1, "B199-DONE"
    Close #1
End Sub
