Option Explicit

' VB6 自带 VBA 类型库里的 VBA.Math 模块（PhotoDemon 的 Math.Log / Math.Abs / Math.Cos …，
' 剩余 VB3001 里的 18 条就是这么来的）。成员是同一批内建函数的限定写法；
' 值位的裸 `Math` 仍然不是合法名字（负控在 test_vbamath_bad.bas）。

Public Sub Main()
    Dim d As Double

    d = Math.Log(1#)
    Debug.Print "VM-LOG=" & d

    Debug.Print "VM-ABS=" & Math.Abs(-4.5)
    Debug.Print "VM-COS=" & Math.Cos(0#)
    Debug.Print "VM-SIN=" & Math.Sin(0#)
    Debug.Print "VM-SQR=" & Math.Sqr(9#)
    Debug.Print "VM-ROUND=" & Math.Round(2.6)

    Debug.Print "VM PASS"
End Sub
