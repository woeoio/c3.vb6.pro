Option Explicit

' ============================================================
'  test_split_result_subscript.bas - 交接 §7.8 缺陷 A 的最小回归夹具
'
'  形状: `Split(s, ",")(1)` = **调用结果再下标**. 外层括号是对 Split 返回的
'  String 数组的元素下标, 不是 Split 的第 5 个实参.
'
'  修前发码 (喂给 vb6_Split 8 个实参, C2197):
'    vb6_Split(src, delim, -1, 0, 1, NULL, -1, 0)
'  修后发码 (外层下标落在括号层级之外):
'    VB6_SA_AT(BSTR, vb6_Split(src, delim, -1, 0), 1)
'
'  同时钉住同族的 `Array(...)(i)`: 匿名数组临时量 `_arr_N` 同理, 外层下标
'  不许并入 vb6_ArrayCreate 的实参表.
' ============================================================

Sub Main()
    Dim s As String
    Dim v As Variant
    Dim host As String

    s = Split("a,b,c", ",")(1)
    v = Array(10, 20, 30)(1)
    host = Split("x,y", ",")(0)

    Debug.Print s
    Debug.Print CStr(v)
    Debug.Print host
End Sub
