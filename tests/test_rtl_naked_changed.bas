Option Explicit

' <vbeclipse> 账 #219: RTL 曾在文件作用域导出 int16_t Changed 这枚**外部链接的裸名 C
' 全局, 用来接住 .pag 里裸写的 Changed。工程里一发 Public Changed As Long 就与它撞成
' C2371 重定义, 连 exe 都不出 (改前实测 .build/b229out/pjChanged.bas: BUILD-RC=1 / no exe,
' 与账 #220 那两枚旗标同型)。脏标记从来只有一个带前缀的名字在承载 (语料 186 处、裸名 0 处),
' 所以那枚全局是纯负担 —— 撤掉之后这一枚名字整个归用户。
' NC219-* 钉的是「Changed 真归用户了」: 累加、装箱两形都按用户的意思走。

Public Changed As Long

Public Sub Main()
    Dim i As Long

    Changed = 3
    For i = 1 To 2
        Changed = Changed + i
    Next
    Debug.Print "NC219-CHANGED=" & Changed
    Debug.Print "NC219-VT=" & VarType(Changed)
    Debug.Print "NC219-DONE"
End Sub
