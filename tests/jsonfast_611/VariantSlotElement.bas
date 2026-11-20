Option Explicit

' Account #611 (#7 保真): 两个独立口径，以前坏在同一格。
'
' 1) `UBound(x, 1)` / `LBound(x, 1)` 的首参是 **Variant 槽**时，秩是运行期属性 ——
'    同一枚 ByRef Variant 形参这一趟可能装 1D、下一趟装 2D。改前：同一过程里只要有
'    一处 `UBound(x, 2)`，'按名字猜秩' 那条判据就把名字登记成 ND，连 `UBound(x, 1)`
'    也改发 ND 版；1D 描述符被按 ND 布局读（1D 的 uBound 正好落在 ND 的
'    bounds[0].cElements 上）⇒ UBound 交回 lBound + uBound - 1 = 0 ⇒ 循环只走一枚。
'    实测 JSonFast 按钮 7：roles 只输出 "admin"，丢了 "auditor"。
'
' 2) Variant() 的元素交给 `ByRef x As Variant` 形参时必须整枚 VARIANT 过去（vt 一起）。
'    改前被压成 `{.vt=VT_BSTR, .bstrVal=VariantToString(元素)}` ⇒ 元素是嵌套数组就
'    交 ""、是 Boolean 就交 "True"（vt 一律变成 8）。
'
' 期望（真 VB6 口径）：
'   ub=2 | vt11 | vt8 | vt8204
' 改前：
'   ub=0 | vt8            (只剩一枚, 且 Boolean 被压成串)
Sub Main()
    Dim a(0 To 2) As Variant
    a(0) = True
    a(1) = "admin"
    a(2) = Array(1, 2)
    Walk a, 1
End Sub

' rank 那一段既触发'按名字猜秩'的登记（`UBound(v, 2)` 在文本里），又在 1D 那一趟
' 不执行 —— 与 JSonFast 的 `If cDims = 2 Then ...` 同一个形状。
Private Sub Walk(ByRef v As Variant, ByVal rank As Long)
    Dim i As Long
    Dim s As String
    Dim d2 As String

    If rank > 1 Then
        d2 = CStr(LBound(v, 2)) & ".." & CStr(UBound(v, 2))
    End If

    s = "ub=" & UBound(v, 1) & " d2=" & d2
    For i = LBound(v, 1) To UBound(v, 1)
        s = s & " | " & Tag(v(i))
    Next i
    Debug.Print s
End Sub

' 元素经 ByRef Variant 形参过一道 —— vt 身份必须活下来。
Private Function Tag(ByRef x As Variant) As String
    Tag = "vt" & VarType(x)
End Function
