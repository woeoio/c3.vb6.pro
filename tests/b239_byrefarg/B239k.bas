Attribute VB_Name = "B239k"
Option Explicit

' 账 #305 §B224 第二半 (第二百三十九刀): 晚绑定实参按引用打包的形状针 (只喂 --emit-c)。
' 真 VB6 的晚绑定对**带型变量**递 VT_BYREF|T (ai/probes A/B: `o.Bump(x)` ⇒ x=11);
' 这里钉五档 Ref 取址, 同时钉两枚反面: 字面量与表达式不是 lvalue, 仍按值。
Public Sub Main()
    Dim o As Object
    Dim n As Long
    Dim i As Integer
    Dim d As Double
    Dim s As Single
    Dim b As Boolean
    Set o = CreateObject("Scripting.Dictionary")
    n = 1
    i = 2
    d = 3
    s = 4
    b = True
    o.Add n, "k1"
    o.Add i, "k2"
    o.Add d, "k3"
    o.Add s, "k4"
    o.Add b, "k5"
    o.Add 7, "lit"
    o.Add n + 1, "expr"
End Sub
