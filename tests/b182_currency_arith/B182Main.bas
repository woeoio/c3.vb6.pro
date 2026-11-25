Option Explicit

' 账 #305 §11.162 (第一百七十九刀): Currency 参与算术后的身份。
'
' 真 VB6 探针 .temp/k179p (P179.bas) 读数 —— Currency 与 Double 的规则**不对称**:
'   + - : Currency 吸收一切数值 (c+0 / c-0 / c+0.5 / c+d / c+s / -c 都是 Currency 6)
'   *   : 遇到 Single/Double 就在浮点里算 (c*1.5 / 1.5*c / c*d 都是 Double 5, 且不圈)
'         只与 Integer/Long/Currency 相乘才回到 Currency (c*1 / c*2 / c*i / c*10000&)
'   /   : 恒 Double (c/1 = 5)
' 改前 C3 一律答 Double (Currency 变量登记在 knownDoubleVars_ 里, inferExprType 先答
' Double), 于是 A1/A2/A3/A5/A6/A7/A10/B1/B3/C1/D1/V1/V2 全打 5|Double —— 其中 B3
' 还打 20.2469998703003 (Single 拓宽后的尾巴), 真机是 20.247。
'
' 每格四列: VarType | TypeName | CStr(x) | CStr(CLng(x*10000))
' E1/E2/E3 是 Double 侧的值证人 (不许被顺手圈成 4 位)。

Public Sub Main()
    Dim c As Currency
    Dim d As Double
    Dim s As Single
    Dim i As Long
    Dim v As Variant

    c = 10.1235
    d = 10.1235
    s = 10.1235
    i = 1

    v = c * 1:      PrintV "A1", v
    v = c + 0:      PrintV "A2", v
    v = c - 0:      PrintV "A3", v
    v = c / 1:      PrintV "A4", v
    v = -c:         PrintV "A5", v
    v = c * 2:      PrintV "A6", v
    v = c * i:      PrintV "A7", v
    v = c * 1.5:    PrintV "A8", v
    v = 1.5 * c:    PrintV "A9", v
    v = c + 0.5:    PrintV "A10", v
    v = c + d:      PrintV "B1", v
    v = c * d:      PrintV "B2", v
    v = c + s:      PrintV "B3", v
    v = c * 10000&: PrintV "C1", v
    v = (c * 1) * 1: PrintV "D1", v
    v = c / 3:      PrintV "E1", v
    v = c * 1.0001: PrintV "E2", v
    v = c * 2 * 1.0001: PrintV "E3", v

    Debug.Print "V1=" & VarType(c * 1) & "|" & TypeName(c * 1)
    Debug.Print "V2=" & VarType(-c) & "|" & TypeName(-c)
    Debug.Print "B182-DONE"
End Sub

Private Sub PrintV(ByVal tag As String, ByVal v As Variant)
    Debug.Print tag & "=" & VarType(v) & "|" & TypeName(v) & "|" & _
                  CStr(v) & "|" & CStr(CLng(v * 10000))
End Sub
