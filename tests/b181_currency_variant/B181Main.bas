Option Explicit

' 账 #305 §11.161 (第一百七十八刀): Currency 变量的身份是 VT_CY。
'
' 真 VB6 探针 .temp/k178p (P178.bas) 读数:
'   P1  v = c (局部 Currency)         -> type=6 Currency
'   P12 ShowV "P12", c (ByVal Variant) -> type=6
'   P13 v = m_c (模块级 Currency)      -> type=6
'   P6/P7/P8 Double 变量/字面量        -> type=5 (反向证人)
'   P9  v = 1 (Integer 字面量)         -> type=2
' Currency 在 C 侧就是 double, 所以 inferExprType 与 _Generic 都只能答 VT_R8 ——
' 改前 C3 的 P1 打 type=5 Double。
'
' 每格四列: VarType | TypeName | CStr(x) | CStr(CLng(x*10000))
' 第三列才是"值"的证人 (VT_CY 也要能原样回读 10.1235)。

Private m_c As Currency

Public Sub Main()
    Dim c As Currency
    Dim d As Double
    Dim v As Variant

    c = 10.1235
    v = c:              PrintV "P1", v     ' 局部 Currency 变量
    v = CCur(10.123456): PrintV "P2", v    ' CCur 表达式 (改前就对)
    v = CurFn():        PrintV "P3", v     ' 返回 Currency 的函数

    d = 10.1235
    v = d:              PrintV "P6", v     ' 反向证人: Double 必须 VT_R8
    v = 10.1235:        PrintV "P7", v
    v = 1.5:            PrintV "P8", v
    v = 1:              PrintV "P9", v     ' Integer 字面量

    Debug.Print "P11=" & TypeName(c) & "|" & VarType(c)

    PrintV "P12", c                        ' ByVal Variant 形参 (实参站)
    m_c = 5.4321
    v = m_c:            PrintV "P13", v    ' 模块级 Currency
    PrintV "P14", m_c

    Debug.Print "B181-DONE"
End Sub

Private Sub PrintV(ByVal tag As String, ByVal v As Variant)
    Debug.Print tag & "=" & VarType(v) & "|" & TypeName(v) & "|" & _
                  CStr(v) & "|" & CStr(CLng(v * 10000))
End Sub

Private Function CurFn() As Currency
    CurFn = 10.123456
End Function
