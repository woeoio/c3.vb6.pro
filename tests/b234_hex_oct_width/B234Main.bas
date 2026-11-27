Option Explicit

' 11.210 (knife 227): the *width* of Hex/Oct follows the carrier: Byte -> 2 digits,
' Integer/Boolean -> 4, everything else -> 8. Positive values are not zero padded, so
' the width only shows on negatives (two's complement truncated to that many bits).
'
' Oracle .temp/k225p/P308: Hex(CByte(255))="FF", Hex(CInt(-1))="FFFF", Hex(True)="FFFF",
' Hex(CLng(-1))="FFFFFFFF", Oct(CInt(-32768))="100000", and the Variant carriers follow
' their VT (VT_I2 -> 4, VT_UI1 -> 2, VT_I4 -> 4/8 per type).
Public Sub Main()
    Dim b As Byte
    Dim i As Integer
    Dim l As Long
    Dim bo As Boolean
    Dim v As Variant

    ' 静态载体：宽度由声明档决定（负值才看得出）
    b = 255
    i = -1
    l = -1
    bo = True
    Debug.Print "W1=" & Hex(b) & " " & Hex(i) & " " & Hex(l) & " " & Hex(bo)
    Debug.Print "W2=" & Oct(b) & " " & Oct(i) & " " & Oct(l) & " " & Oct(bo)

    ' 正值不补零
    Debug.Print "W3=" & Hex(CByte(16)) & " " & Hex(CInt(16)) & " " & Hex(CLng(16))
    Debug.Print "W4=" & Hex(CByte(0)) & " " & Hex(CInt(0)) & " " & Hex(CLng(0))

    ' 各档上下界
    Debug.Print "W5=" & Hex(CByte(128)) & " " & Hex(CInt(-32768)) & " " & Hex(CLng(-2147483648))
    Debug.Print "W6=" & Oct(CByte(128)) & " " & Oct(CInt(-32768)) & " " & Oct(CLng(-2147483648))
    Debug.Print "W7=" & Hex(CInt(-2)) & " " & Hex(CByte(255)) & " " & Hex(CLng(-2))

    ' 变体载体：宽度跟 VT 走
    v = CInt(-1):   Debug.Print "V1=" & VarType(v) & "/" & Hex(v)
    v = CByte(255): Debug.Print "V2=" & VarType(v) & "/" & Hex(v)
    v = CLng(-1):   Debug.Print "V3=" & VarType(v) & "/" & Hex(v)
    v = True:       Debug.Print "V4=" & VarType(v) & "/" & Hex(v)
    v = CInt(-1):   Debug.Print "V5=" & Oct(v)
    v = CByte(255): Debug.Print "V6=" & Oct(v)
    v = CLng(-1):   Debug.Print "V7=" & Oct(v)

    ' 浮点/日期/单精度载体：先收窄到 Long，宽度 8
    v = 1.5:        Debug.Print "V8=" & Hex(v) & " " & VarType(v)
    v = CDate(-1.4): Debug.Print "V9=" & Hex(v) & " " & VarType(v)
    v = CSng(-1):   Debug.Print "V10=" & Hex(v) & " " & VarType(v)

    ' 字面量档（Integer 字面量与 Long 字面量同形，因为正值）
    Debug.Print "L1=" & Hex(1) & " " & Hex(40000)
    Debug.Print "L2=" & Oct(1) & " " & Oct(40000)
End Sub