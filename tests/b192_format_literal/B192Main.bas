Option Explicit

' 账 #305 §11.167 (第一百八十四刀): Format 数值段里的字面量与 Percent 的分组。
'
' 真 VB6 探针 .temp/k184p (P197.bas) 逐格读数:
'   H1 (0.00)@1234.5678        -> (1234.57)      字面量括号要在
'   H2 0.00@-1234.5678         -> -1234.57
'   H3 0.00;(0.00)@-1234.5678  -> (1234.57)      负号段走字面量、不再补 '-'
'   H4 0.00"x"                 -> 1234.57x       引号内原样
'   H5 0.00-                   -> 1234.57-
'   H6 0.00-@-1234.5678        -> -1234.57-
'   H7 \#0.00                  -> #1234.57      反斜杠转义
'   H8 0.00;0.0;zero@正        -> 1234.57
'   H9 0.00;0.0;zero@0         -> zero          零段没有占位符 ⇒ 只打字面量
'   H10 Percent                -> 123456.78%    不带千分位
'   H11 Percent@负             -> -123456.78%
'   H12 Percent@0.125          -> 12.50%
'   H13 Fixed                  -> 1234.57
'   H14 Standard@1.5           -> 1.50
'   H15 #.##@0.5               -> .5            整数段只有 '#' 且值为 0 ⇒ 不打前导 0
'   H16 0.0;;zero@1            -> 1.0
'
' 改前 C3: H1/H3/H4/H5/H6/H7 全丢字面量；H9 打 "0"；H10/H11 多千分位；H15 打 "0.5"。

Public Sub Main()
    Dim d As Double
    d = 1234.5678

    Show "H1", Format(d, "(0.00)")
    Show "H2", Format(-d, "0.00")
    Show "H3", Format(-d, "0.00;(0.00)")
    Show "H4", Format(d, "0.00""x""")
    Show "H5", Format(d, "0.00-")
    Show "H6", Format(-d, "0.00-")
    Show "H7", Format(d, "\#0.00")
    Show "H8", Format(d, "0.00;0.0;zero")
    Show "H9", Format(0, "0.00;0.0;zero")
    Show "H10", Format(d, "Percent")
    Show "H11", Format(-d, "Percent")
    Show "H12", Format(0.125, "Percent")
    Show "H13", Format(d, "Fixed")
    Show "H14", Format(1.5, "Standard")
    Show "H15", Format(0.5, "#.##")
    Show "H16", Format(1, "0.0;;zero")

    Debug.Print "B192-DONE"
End Sub

Private Sub Show(ByVal tag As String, ByVal s As String)
    Debug.Print tag & "=[" & s & "]"
End Sub
