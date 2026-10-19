Attribute VB_Name = "ModB434"
Option Explicit

' 账 #434 (§B208, 第一百零六刀): 「数值形参吃 As String 实参」—— §B207 的另一半。
'
' 病形与 §B207 同形 (裸 BSTR 直递 = cl C2440), 但**答案不同**, 判别格就写在下面两行里:
'   Sqr("2026/1/5")   = err 13   ← 日期串进数值形参要**报错**
'   DateAdd("d","3","2026/1/5") = 2026/1/8   ← 同一枚 builtin 的位 1 走**数值**折算
' 所以这一族绝不许复用 vb6_CDate(会把日期串折成序列号再开方), 出口是 vb6_CDbl。
' `vb6_Val` 也不行: 它是 Val 的宽松语义(遇非数字就停、永不报错), 与引擎要的 13 相反。
'
' 承面 = 0 (真语料里没有一处把文本喂给数值形参) —— 这一格是 §B207 的**负控**顺手量出来的,
' 当时写成「另一笔账」。之所以仍然补: 它与 §B207 是同一格问句的两半, 承面量测显示 0 处,
' 所以形状门 467 行不动, 风险面只在**新增**那一行夹具。

Sub Main()
    Dim s4 As String
    Dim s37 As String
    Dim sDate As String
    Dim sNum As String
    s4 = "4"
    s37 = "3.7"
    sDate = "2026/1/5"
    sNum = "46027"

    ' --- 一元数值族 (引擎 cscript 尺 .build/k434_oracle.vbs 逐格相同) ---
    Debug.Print "SQR=" & Sqr(s4)                 ' 2
    Debug.Print "ABS=" & Abs(s4)                 ' 4
    Debug.Print "INT=" & Int(s37)                ' 3
    Debug.Print "SGN=" & Sgn(s4)                 ' 1
    Debug.Print "SQRlit=" & Sqr("16")            ' 4  (字面量那一格)

    ' --- 判别格: 日期串进数值形参必须报错, 不许折成序列号 ---
    Debug.Print "SQRdate-err=" & ErrOfSqr(sDate)      ' 13
    Debug.Print "SQRabc-err=" & ErrOfSqr("abc")        ' 13
    ' 引擎**收**的两格: 区域化数字 + 前后空白 —— 证明这一族折算宽松, 不是 Val 的"遇非数字就停"
    Debug.Print "SQRfmt-err=" & ErrOfSqr("1,234.5")    ' 0
    Debug.Print "SQRspc-err=" & ErrOfSqr(" 42 ")       ' 0

    ' --- 同一枚 DateAdd 的两个位答案不同 (逐位登记的证据) ---
    Debug.Print "ADDnum=" & Format$(DateAdd("d", s4, sDate), "yyyy/mm/dd")   ' 2026/01/09
    Debug.Print "ADDdate=" & Format$(DateAdd("d", 3, sNum), "yyyy/mm/dd")     ' 2026/01/08

    ' --- 负控: §B207 那一族仍是**日期**解析, 没被本刀改成数值折算 ---
    Debug.Print "FDT-ok=" & (IIf(Year(sDate) = 2026 And Month(sDate) = 1 And Day(sDate) = 5, 1, 0))  ' 1

    Debug.Print "B434-DONE"
End Sub

Private Function ErrOfSqr(s As String) As Long
    Dim d As Double
    Dim n As Long
    On Error Resume Next
    d = Sqr(s)
    n = Err.Number
    Err.Clear
    On Error GoTo 0
    ErrOfSqr = n
End Function