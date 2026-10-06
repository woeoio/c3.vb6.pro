Option Explicit

' <vbeclipse> 账 #214 = 多维数组的**元素**访问收成一处带检查的出口（一维那一半在账 #209）。
' 三条读数全取自真实输出（本地 x86+x64 各真编真跑）：
'   · `Dim a2(1 To 2, 1 To 3)` 读 a2(3,1) 以前**静默拿到隔壁那格**（回 12），VB6 是错误 9；
'   · `Dim d() As Long` 从没 ReDim 就取 d(1,1) 以前读 (NULL)->data ⇒ 0xC0000005；
'   · 4 秩数组以前**在范围内也当场崩** —— 发码那一条把实参拼成 `(int[]){i0, i1}` 只塞两个下标
'     却按实际秩数交出去，偏移成了垃圾。这一条是本轮唯一"改坏了功能"的洞（1/2/3 秩一直是对的）。
' NA-in（第四个数就是 4 秩范围内的那一格）与 NA-dyn-in 两枚是**负控**：在范围内的读写必须照旧出数，
' 检查误报就红在它们身上。

Private Type NAPair
    A As Long
    B As Long
End Type

Public Sub Main()
    Dim i As Long, j As Long, k As Long, m As Long
    Dim got As Long
    Dim bad As Long
    Dim a2(1 To 2, 1 To 3) As Long
    Dim a3(1 To 2, 1 To 2, 1 To 2) As Long
    Dim a4(1 To 2, 1 To 2, 1 To 2, 1 To 2) As Long
    Dim s2(1 To 2, 1 To 2) As NAPair

    For i = 1 To 2
        For j = 1 To 3
            a2(i, j) = i * 10 + j
        Next
    Next
    For i = 1 To 2
        For j = 1 To 3
            If a2(i, j) <> i * 10 + j Then bad = bad + 1
        Next
    Next

    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                a3(i, j, k) = i * 100 + j * 10 + k
            Next
        Next
    Next
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                If a3(i, j, k) <> i * 100 + j * 10 + k Then bad = bad + 1
            Next
        Next
    Next

    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                For m = 1 To 2
                    a4(i, j, k, m) = i * 1000 + j * 100 + k * 10 + m
                Next
            Next
        Next
    Next
    For i = 1 To 2
        For j = 1 To 2
            For k = 1 To 2
                For m = 1 To 2
                    If a4(i, j, k, m) <> i * 1000 + j * 100 + k * 10 + m Then bad = bad + 1
                Next
            Next
        Next
    Next

    ' UDT 元素 + 多维（VBFlexGrid 那一族的形状）
    s2(1, 1).A = 5
    s2(2, 2).B = 6
    Debug.Print "NA-udt=" & s2(1, 1).A & "/" & s2(2, 2).B

    Debug.Print "NA-in=" & bad & "/" & a2(2, 3) & "/" & a3(2, 1, 1) & "/" & a4(2, 1, 1, 2)

    Call NaAbove(1)
    Call NaBelow(2)
    Call NaNullDyn(3)
    Call NaDynIn(4)
    Call NaRankMismatch(5)
    Debug.Print "NA-DONE"
End Sub

' 上界外: VB6 = 错误 9
Private Sub NaAbove(ByVal n As Long)
    Dim a2(1 To 2, 1 To 3) As Long
    Dim got As Long
    On Error GoTo H
    got = a2(3, 1)
    Debug.Print "NA-above=NOERR " & got
    Exit Sub
H:
    Debug.Print "NA-above=" & Err.Number
End Sub

' 下界外（秩 2 的下界是 1）
Private Sub NaBelow(ByVal n As Long)
    Dim a2(1 To 2, 1 To 3) As Long
    Dim got As Long
    On Error GoTo H
    got = a2(1, 0)
    Debug.Print "NA-below=NOERR " & got
    Exit Sub
H:
    Debug.Print "NA-below=" & Err.Number
End Sub

' 从没 ReDim 的动态数组: 以前读 (NULL)->data 当场 0xC0000005
Private Sub NaNullDyn(ByVal n As Long)
    Dim d() As Long
    Dim got As Long
    On Error GoTo H
    got = d(1, 1)
    Debug.Print "NA-null=NOERR " & got
    Exit Sub
H:
    Debug.Print "NA-null=" & Err.Number
End Sub

' 动态二维 ReDim 之后的范围内读写（负控: 不许误报）
Private Sub NaDynIn(ByVal n As Long)
    Dim d() As Long
    Dim i As Long, j As Long
    Dim sum As Long
    ReDim d(1 To 2, 1 To 3)
    For i = 1 To 2
        For j = 1 To 3
            d(i, j) = i * 10 + j
        Next
    Next
    For i = 1 To 2
        For j = 1 To 3
            sum = sum + d(i, j)
        Next
    Next
    Debug.Print "NA-dyn-in=" & sum & "/" & d(2, 3)
End Sub

' 秩数与声明不符: VB6 在**编译期**就拒（"Number of dimensions doesn't match"），
' 我们现在是编译得过、运行期交回错误 9 —— 这一针钉的是"不再拿越界的下标去寻址"，
' 不是"这就是 VB6 的口径"（发码期那条判死是另一件事，见台账 §B46 末段）。
Private Sub NaRankMismatch(ByVal n As Long)
    Dim d() As Long
    Dim got As Long
    ReDim d(1 To 2, 1 To 3)
    On Error GoTo H
    got = d(1, 1, 1)
    Debug.Print "NA-rank=NOERR " & got
    Exit Sub
H:
    Debug.Print "NA-rank=" & Err.Number
End Sub
