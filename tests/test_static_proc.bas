' Option: 账 #305 §B218 (甲) 第一百六十一刀的夹具 = Static 过程头那一形（parser 从前直接拒收）。
' 五格各钉一件事：
'   cnt  = Private Static Sub 里那枚标量跨三次调用累加（VB6 口径：Static 过程里的 Dim 就是静态存储）
'   acc  = Static Function 里那枚标量：3 / 6 / 9（返回值就是静态那枚的读数）
'   fa   = 同一枚 Static 过程里的定长数组：第一次写、第二次读（第一百六十刀那发首轮守卫的读者）
'   plain= 反面证人：普通（非 Static）Sub 里同形状的标量每次进来归零，两条臂共用同一位主人
'   mg   = 模块级 Static 数组那一档不受本刀影响（Static 后面跟的是标识符，仍走变量声明那条老路）
Option Explicit

Static mg(0 To 2) As Long

Private Static Sub Counter()
    Dim n As Long
    Dim fa(0 To 2) As Long
    n = n + 1
    If n = 1 Then
        fa(2) = 77
    End If
    Debug.Print "K161-cnt=" & n & " fa2=" & fa(2)
End Sub

Static Function Acc() As Long
    Dim t As Long
    t = t + 3
    Acc = t
End Function

Sub PlainReset()
    Dim n2 As Long
    n2 = n2 + 1
    Debug.Print "K161-plain=" & n2
End Sub

Public Sub Main()
    Counter
    Counter
    Counter
    Debug.Print "K161-a1=" & Acc
    Debug.Print "K161-a2=" & Acc
    Debug.Print "K161-a3=" & Acc
    PlainReset
    PlainReset
    mg(1) = 5
    Debug.Print "K161-mg=" & mg(1)
    Debug.Print "K161-DONE=1"
End Sub
