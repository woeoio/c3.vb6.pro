Option Explicit

' 账 §38: VB6 的 Variant 比较里, **String 与数值混比**有专有次序 ——
' 「数值侧恒小于字符串侧」, 与字符串内容能否转数无关。实测 (真 VB6 /make 编译):
'   "0" = 0 → False    "0" > 0 → True    "1" < 2 → False    0 < "x" → True
' 且与**类型化** String 不同: `Dim s As String: s = "0": (0 = s)` → True (数值转换)。
' Empty 是"随对方"档: 与数值比当 0, 与字符串比当 ""。
' 这一条是 ucTreeViewPlus.ctl 里 `If N.Image > 0` (Image 是 Variant 字符串键) 的上游:
' C3 曾按 wcstod 抹成 0 → 该分支不进 → 行框丢掉图片宽度 → Form3/Form5 文字被截断。
Public Sub Main()
    Dim v As Variant
    Dim n As Variant
    Dim e As Variant
    Dim b As Boolean

    v = "John_Perez"
    n = 0
    b = (v > n):  Debug.Print "VCS-sgt0="; b
    b = (v < n):  Debug.Print "VCS-slt0="; b
    b = (v = n):  Debug.Print "VCS-seq0="; b
    b = (v <> n): Debug.Print "VCS-sne0="; b
    b = (v >= n): Debug.Print "VCS-sge0="; b
    b = (v <= n): Debug.Print "VCS-sle0="; b

    v = "1"
    n = 2
    b = (v < n):  Debug.Print "VCS-str1lt2="; b
    v = "10"
    n = 9
    b = (v > n):  Debug.Print "VCS-str10gt9="; b

    n = "John_Perez"
    v = 0
    b = (v < n):  Debug.Print "VCS-n0lts="; b

    ' Variant vs 非 Variant 标量 (cgen 走 vb6_VarCmpLong* 快捷形, 与上面的 VarCmp* 是两条发码)
    v = "x"
    b = (v > 0):  Debug.Print "VCS-long-gt="; b
    v = "0"
    b = (v = 0):  Debug.Print "VCS-long-eq="; b
    v = "x"
    b = (0 < v):  Debug.Print "VCS-long-lt="; b

    ' Empty 是"随对方"档
    e = 0
    v = 0
    b = (v = e):    Debug.Print "VCS-empty-eq0="; b
    b = (e = "x"):  Debug.Print "VCS-empty-eqs="; b
    b = (e < "x"):  Debug.Print "VCS-empty-lts="; b

    Debug.Print "VCS-DONE"
End Sub
