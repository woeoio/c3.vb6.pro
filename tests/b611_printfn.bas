Attribute VB_Name = "T611"
Option Explicit

' 账 #609 (台账 §B11): Print # / Write # 的实参是「返回某档类型的工程 Function」时，
' 该以什么形状交给写文件的出口。修前那半自留一条 `vb6_Str((int32_t)(val))`，一次塌三档：
'   1) As String  的函数 ⇒ BSTR 指针被截成 int32，文件里只剩一枚指针数 (§B11)；
'   2) As Variant 的函数 ⇒ **整棵工程编不过** (C2440: 无法从 vb6_VARIANT 转换到 int32_t)；
'   3) As Double / As Date / As Boolean ⇒ 被 (int32_t) 截成整数 (1.5 打成 1、日期打成序列数)。
' 第331刀把这一半改走 Form Print 早已在用的那位主人 (wrapToBSTR 按档分诊)，三户一条口。
' 判据一律从**文件回读**拿；第 4 行那枚 Long 是**数值档中立性证人** —— 改前改后同值，
' 带动它就算回归。载荷全 ASCII ⇒ 结果与码页无关 (ai/029 的 ASCII 载荷纪律)。

Public Function StrFn() As String
    StrFn = "StrFn-611"
End Function

Public Function VarFn() As Variant
    VarFn = 7
End Function

Public Function NumFn() As Long
    NumFn = 4242
End Function

Public Function DblFn() As Double
    DblFn = 1.5
End Function

Public Function BoolFn() As Boolean
    BoolFn = True
End Function

Public Function DateFn() As Date
    DateFn = DateSerial(2026, 11, 2)
End Function

Sub Main()
    Dim fn As Integer

    fn = FreeFile
    Open "t611_out.txt" For Output As #fn
    Print #fn, StrFn()
    Print #fn, StrFn
    Print #fn, VarFn()
    Print #fn, NumFn()
    Print #fn, DblFn()
    Print #fn, BoolFn()
    Print #fn, DateFn()
    Write #fn, StrFn()
    Close #fn

    Debug.Print "611-strfn=[" & ReadLine(1) & "] ok=" & (ReadLine(1) = "StrFn-611")
    Debug.Print "611-strfn-bare ok=" & (ReadLine(2) = "StrFn-611")
    Debug.Print "611-varfn=[" & ReadLine(3) & "] ok=" & (Trim(ReadLine(3)) = "7")
    Debug.Print "611-num=[" & ReadLine(4) & "] neutral=" & (Trim(ReadLine(4)) = "4242")
    Debug.Print "611-dbl-not-truncated ok=" & (Trim(ReadLine(5)) <> "1")
    Debug.Print "611-bool=[" & ReadLine(6) & "] ok=" & (Trim(ReadLine(6)) = "True")
    Debug.Print "611-date-in-year ok=" & (InStr(ReadLine(7), "2026") > 0)
    Debug.Print "611-date-not-serial ok=" & Not IsDigits(Trim(ReadLine(7)))

    ' Write # 那一枚：串该带引号（同一户的另一支，形状针与真跑各钉一次）
    Dim all As String
    Dim rb As String
    all = ""
    fn = FreeFile
    Open "t611_out.txt" For Input As #fn
    Do While Not EOF(fn)
        Line Input #fn, rb
        all = all & rb & "|"
    Loop
    Close #fn
    Debug.Print "611-write-quoted ok=" & (InStr(all, Chr(34) & "StrFn-611" & Chr(34)) > 0)

    Kill "t611_out.txt"
    Debug.Print "611-DONE"
End Sub

' 回读第 n 行 (1 起) —— 与 test_fileio_modstr 同一把手，每次从头数
Private Function ReadLine(ByVal n As Long) As String
    Dim fn As Integer
    Dim i As Long
    ReadLine = ""
    fn = FreeFile
    Open "t611_out.txt" For Input As #fn
    For i = 1 To n
        If EOF(fn) Then Exit For
        Line Input #fn, ReadLine
    Next i
    Close #fn
End Function

Private Function IsDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then
        IsDigits = False
        Exit Function
    End If
    For i = 1 To Len(s)
        If Mid(s, i, 1) < "0" Or Mid(s, i, 1) > "9" Then
            IsDigits = False
            Exit Function
        End If
    Next i
    IsDigits = True
End Function
