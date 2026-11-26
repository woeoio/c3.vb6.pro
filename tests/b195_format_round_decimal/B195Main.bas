Option Explicit

' 账 #305 §11.170 (第一百八十七刀): FormatNumber / FormatCurrency / FormatPercent 的十进制半值进位。
'
' 病: 三家出口都用 C 运行库的 `%.*f` —— 那是**二进制**五成双, 于是
'   FormatNumber(0.5,0)="0"、FormatNumber(2.5,0)="2"、FormatNumber(1.005,2)="1.00"。
' 真机是**十进制**半值远离零 (先把值归到 15 位有效数字的十进制文本, 再逐位进位)。
'
' 真 VB6 读数 (.temp/k187p/P205~P208, 逐字抄自 p2xx.txt):
'   D1  0.5@0   -> 1        D2  1.5@0 -> 2        D3  2.5@0 -> 3
'   D4  3.5@0   -> 4        D5 -0.5@0 -> -1       D6 -1.5@0 -> -2
'   D7  0.05@1  -> 0.1      D8  0.15@1 -> 0.2     D9  0.25@1 -> 0.3
'   D10 0.125@2 -> 0.13     D11 FormatPercent(0.5,0) -> 50%
'   D13 1234.5@0 -> 1,235
'   R1  1.005@2 -> 1.01     (二进制展开是 1.00499… ⇒ %.*f 给 1.00)
'   R2  2.675@2 -> 2.68     R3  0.145@2 -> 0.15    R4  1.0049999@2 -> 1.00
'   R5  2.5@0   -> 3        R6  1234.565@2 -> 1,234.57
'   R7  FormatPercent(0.005,1) -> 0.5%            R8 FormatPercent(0.015,1) -> 1.5%
'   R9  -0.05@1 -> -0.1     R10 0.4999999999999999@0 -> 1     (15 位十进制 = 0.5)
'   R11 1.0005@3 -> 1.001   R12 0.5@0 incLeading=0 -> 1      R13 1.5@0 incLeading=0 -> 2
'   B1  1E+15@2 -> 1,000,000,000,000,000.00       (超大值不走科学记数)
'   B2  1E+20@2 -> 100,000,000,000,000,000,000.00
'   B3  123456789012345.5@0 -> 123,456,789,012,346
'   B4  0.0001@0 -> 0       B5  0@2 -> 0.00       B6  0.0000001@2 -> 0.00
'   B7  1234567890.5@0 -> 1,234,567,891          B8  0.9999999999999999@0 -> 1
'   B9  9.995@2 -> 10.00    (要连进位)            B10 -9.995@2 -> -10.00
'   B11 1E-20@2 -> 0.00     B12 123.456@0 -> 123
'   F1..F7 = FormatNumber 默认档: 1,234.57 / 1,234.6 / 1,234.567 / 1,234.57 /
'            1234.57 (groupDigits=0) / 1 (0.5@0) / -1,234.57
' 货币那一档只钉形状 (符号跟着区域设置走): C1 = Mid(FormatCurrency(0.5,0),2) = "1"

Public Sub Main()
    Show "D1", FormatNumber(0.5, 0)
    Show "D2", FormatNumber(1.5, 0)
    Show "D3", FormatNumber(2.5, 0)
    Show "D4", FormatNumber(3.5, 0)
    Show "D5", FormatNumber(-0.5, 0)
    Show "D6", FormatNumber(-1.5, 0)
    Show "D7", FormatNumber(0.05, 1)
    Show "D8", FormatNumber(0.15, 1)
    Show "D9", FormatNumber(0.25, 1)
    Show "D10", FormatNumber(0.125, 2)
    Show "D11", FormatPercent(0.5, 0)
    Show "D13", FormatNumber(1234.5, 0)

    Show "R1", FormatNumber(1.005, 2)
    Show "R2", FormatNumber(2.675, 2)
    Show "R3", FormatNumber(0.145, 2)
    Show "R4", FormatNumber(1.0049999, 2)
    Show "R5", FormatNumber(2.5, 0)
    Show "R6", FormatNumber(1234.565, 2)
    Show "R7", FormatPercent(0.005, 1)
    Show "R8", FormatPercent(0.015, 1)
    Show "R9", FormatNumber(-0.05, 1)
    Show "R10", FormatNumber(0.4999999999999999, 0)
    Show "R11", FormatNumber(1.0005, 3)
    Show "R12", FormatNumber(0.5, 0, 0)
    Show "R13", FormatNumber(1.5, 0, 0)

    Show "B1", FormatNumber(1E+15, 2)
    Show "B2", FormatNumber(1E+20, 2)
    Show "B3", FormatNumber(123456789012345.5, 0)
    Show "B4", FormatNumber(0.0001, 0)
    Show "B5", FormatNumber(0, 2)
    Show "B6", FormatNumber(0.0000001, 2)
    Show "B7", FormatNumber(1234567890.5, 0)
    Show "B8", FormatNumber(0.9999999999999999, 0)
    Show "B9", FormatNumber(9.995, 2)
    Show "B10", FormatNumber(-9.995, 2)
    Show "B11", FormatNumber(1E-20, 2)
    Show "B12", FormatNumber(123.456, 0)

    Show "F1", FormatNumber(1234.567)
    Show "F2", FormatNumber(1234.567, 1)
    Show "F3", FormatNumber(1234.567, 3, 0)
    Show "F4", FormatNumber(1234.567, 2, 0, 0)
    Show "F5", FormatNumber(1234.567, 2, -1, 0, 0)
    Show "F6", FormatNumber(0.5, 0)
    Show "F7", FormatNumber(-1234.567, 2)

    Show "C1", Mid(FormatCurrency(0.5, 0), 2)

    Debug.Print "B195-DONE"
End Sub

Private Sub Show(ByVal tag As String, ByVal s As String)
    Debug.Print tag & "=[" & s & "]"
End Sub
