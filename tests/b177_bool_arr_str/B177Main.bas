Option Explicit

' 账 #305 §11.157 (第一百七十四刀): Boolean 数组元素在**拼串**里的档位。
'
' 改前 inferExprType 对数组元素读一律答 Variant ⇒ 元素读走装箱 (VB6_SA_AT → int16
' → vb6_VariantFromValue → VT_I2)，拼串交的是数字 "-1"；VB6 交 "True"/"False"。
' 数值档 (Long/Integer/Byte/Double) 装箱后格式化结果碰巧相同，所以这一格只有
' Boolean 能看见差异。
'
' B177-2 是同一件事的另一形：`元素 And True` —— 元素答 Variant 时 And 降成 int32，
' 结果也成了 "-1"；元素答 Boolean 之后 And 自己仍是 Boolean。
' B177-4 是反向证人：Boolean 交给**数值**上下文仍是 -1 (VB6 的 True = -1)。
Public Sub Main()
    Dim f(1 To 2) As Boolean
    f(1) = True
    f(2) = False

    Debug.Print "B177-1=" & f(1)
    Debug.Print "B177-2=" & (f(1) And True)
    Debug.Print "B177-3=" & Not f(1)

    Dim n As Long
    n = f(1)
    Debug.Print "B177-4=" & n

    Dim v As Boolean
    v = True
    Debug.Print "B177-5=" & v

    Debug.Print "B177-6=" & f(2)

    Debug.Print "B177-DONE"
End Sub
