Option Explicit

' 账 #305 §11.164 (第一百八十一刀): Currency 数组元素存 double（与标量同口径）。
'
' 真 VB6 探针 .temp/k181p (P186.bas) 读数:
'   A1 a(1)=10.123456        -> 6 Currency 10.1235   （写入即圈）
'   A2 VarType(a(1))         -> 6
'   A3 a(1)=a(1)+0.00004     -> 6 Currency 10.1235   （每步圈，§11.163）
'   A4 ByRef 传元素 x=x+1.00004 -> 6 Currency 11.1235
'   A5 未赋值元素            -> 6 Currency 0
'   D1 赋给 Double           -> 11.1235
'   A6 a(1)*1                -> 6 Currency 11.1235
'   A7 ReDim 动态数组        -> 6 Currency 10.1235
' 改前 C3: 全部 20|Long 且值被截 (A1 打 10) —— 元素槽是 int64_t，赋值只过 vb6_CCur
' 不乘 10000，小数在存储层就丢了。
'
' 每格四列: VarType | TypeName | CStr(x) | CStr(CLng(x*10000))

Private Sub TakeByRef(x As Currency)
    x = x + 1.00004
End Sub

Public Sub Main()
    Dim a(1 To 3) As Currency
    Dim v As Variant
    Dim d As Double

    a(1) = 10.123456
    PrintV "A1", a(1)
    Debug.Print "A2=" & VarType(a(1)) & "|" & TypeName(a(1))
    a(1) = a(1) + 0.00004
    PrintV "A3", a(1)
    TakeByRef a(1)
    PrintV "A4", a(1)
    PrintV "A5", a(2)
    d = a(1)
    Debug.Print "D1=" & d
    v = a(1) * 1
    PrintV "A6", v

    Dim b() As Currency
    ReDim b(1 To 2)
    b(1) = 10.123456
    PrintV "A7", b(1)

    Debug.Print "B184-DONE"
End Sub

Private Sub PrintV(ByVal tag As String, ByVal v As Variant)
    Debug.Print tag & "=" & VarType(v) & "|" & TypeName(v) & "|" & _
                  CStr(v) & "|" & CStr(CLng(v * 10000))
End Sub
