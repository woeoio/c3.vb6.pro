Option Explicit

' 11.209 (knife 226): Hex/Oct narrow their argument to Long through the *same* station
' as CLng -- half-to-even, and error 6 on overflow -- instead of letting C truncate at
' the int32_t boundary.
'
' Oracle .temp/k225p/P305 + P306: Hex(44259.62)=ACE4 (C3 gave ACE3), Hex(1.5)=2 (gave 1),
' Hex("2.5")=2 (gave a pointer value), Hex(2147483647.6) -> Error 6 with the target
' unchanged (gave 7FFFFFFF). Values are all hex/octal/decimal integers or True/False, so
' nothing here depends on the locale.
'
' Not covered on purpose (their own ledgers): the *width* of the answer for Integer/Byte
' carriers (Hex(-1) = "FFFF" in VB6, 8 digits for Long) and Hex/Oct of a whole array.
Public Sub Main()
    Dim sv As String
    Dim dv As Date
    Dim cv As Currency
    Dim s As String

    ' 载体档：浮点/日期/货币 -> 先四入（半到偶）后进制
    Debug.Print "H1=" & Hex(44259.62)
    Debug.Print "H2=" & Hex(44259.49)
    Debug.Print "H3=" & Hex(44259.5)
    Debug.Print "H4=" & Hex(44259.25)
    Debug.Print "H5=" & Hex(DateSerial(2021, 3, 1))
    Debug.Print "H6=" & Hex(CDate("2021-03-01 12:00:00"))

    ' 平局往偶：1.5->2 / 0.5->0 / 3.5->4 / -2.5->-2
    Debug.Print "H7=" & Hex(2.5)
    Debug.Print "H8=" & Hex(1.5)
    Debug.Print "H9=" & Hex(0.5)
    Debug.Print "H10=" & Hex(3.5)
    Debug.Print "H11=" & Hex(-2.5)
    Debug.Print "H12=" & Hex(CSng(2.5))
    Debug.Print "H13=" & Hex(CCur(2.5))
    Debug.Print "H14=" & Hex(CLng(44259.62))

    ' 变量路（类型来自声明，不来自字面量）
    dv = 44259.62
    Debug.Print "H15=" & Hex(dv)
    cv = 2.5
    Debug.Print "H16=" & Hex(cv)
    sv = "2.5"
    Debug.Print "H17=" & Hex(sv)
    sv = "16"
    Debug.Print "H18=" & Hex(sv)

    ' 非数值载体
    Debug.Print "H19=" & Hex(Empty)
    Debug.Print "H20=" & IsNull(Hex(Null))

    ' 负日期：序列号先四入到 Long 再取补
    Debug.Print "H21=" & Hex(CDate(-1.6))
    Debug.Print "H22=" & Hex(CDate(-1.4))

    ' 边界不越界（2^31-1 附近仍合法）
    Debug.Print "H23=" & Hex(CDbl(2147483647))
    Debug.Print "H24=" & Hex(CDbl(2147483647.4))

    ' Oct 同族
    Debug.Print "O1=" & Oct(44259.62)
    Debug.Print "O2=" & Oct(44259.49)
    Debug.Print "O3=" & Oct(2.5)
    Debug.Print "O4=" & Oct(1.5)
    Debug.Print "O5=" & Oct(3.5)
    Debug.Print "O6=" & Oct(-2.5)
    Debug.Print "O7=" & Oct(Empty)
    Debug.Print "O8=" & IsNull(Oct(Null))
    Debug.Print "O9=" & Oct(CDate("2021-03-01 12:00:00"))

    ' 越界 = Error 6，且目标保持原值
    On Error Resume Next
    Err.Clear
    s = "x"
    s = Hex(CDbl(2147483647.6))
    Debug.Print "Y1=" & Err.Number & "/" & s
    Err.Clear
    s = "x"
    s = Oct(CDbl(2147483647.6))
    Debug.Print "Y2=" & Err.Number & "/" & s
    Err.Clear
    s = "x"
    s = Hex(CDbl(-2147483648.6))
    Debug.Print "Y3=" & Err.Number & "/" & s
    Err.Clear
    Debug.Print "Y4=" & Err.Number & "/" & Hex(CDbl(-2147483648#))
    On Error GoTo 0
End Sub