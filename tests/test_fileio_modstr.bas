Attribute VB_Name = "T608"
Option Explicit

' 账 #608: Print # 与 Write # 对**模块级** String 实参认不认得。修前发成
' vb6_Print(fn, vb6_Str((int32_t)(gAll))) —— 把 BSTR 指针当数打印，文件里只剩一枚指针数
' (根因 = knownBstrVars_ 每过程入口 clear 后只由局部声明与形参重新填，模块级那句进不了表)。
' 判据一律从**文件回读**拿；模块级 Long 那一行是**数值档中立性证人** —— 第331刀把那条 int32 兜底
' 撤了 (改走 wrapToBSTR 按档分诊)，这一枚钉的是「Long 档的读数一字不变」，带动它就算回归。
' 载荷全 ASCII ⇒ 结果与码页无关 (ai/029 的 ASCII 载荷纪律)。

Public gAll As String
Public gNum As Long
Public gFix As String * 5

Sub Main()
    gAll = "ModLevel-String-608"
    gNum = 4242
    gFix = "ab"

    Dim fn As Integer
    Dim rb As String
    Dim loc As String
    loc = "Local-String-608"

    fn = FreeFile
    Open "t608_out.txt" For Output As #fn
    Print #fn, gAll
    Print #fn, loc
    Print #fn, gNum
    Print #fn, gAll & "!"
    Print #fn, gFix
    Write #fn, gAll
    Close #fn

    ' 逐行只数到第 n 行 (第 6 行是 Write 那一枚，带引号，故意不回读)
    Debug.Print "608-modstr=[" & ReadLine(1) & "] ok=" & (ReadLine(1) = gAll)
    Debug.Print "608-local=[" & ReadLine(2) & "] ok=" & (ReadLine(2) = loc)
    ' Print # 数值按 VB6 Str 语义写盘 (正数一个前导空格、恒一个尾空格), 见
    ' .temp/m4/vb6num/p05_long_var.txt = " 123456 " —— 故回读比对用 Trim, 那一枚仍是
    ' 「数值读数不变」的证人 (4242 不许被折成指针数)。
    Debug.Print "608-num=[" & ReadLine(3) & "] ok=" & (Trim(ReadLine(3)) = CStr(gNum))
    Debug.Print "608-concat=[" & ReadLine(4) & "] ok=" & (ReadLine(4) = gAll & "!")
    Debug.Print "608-fixstr-not-digits ok=" & Not IsDigits(Trim(ReadLine(5)))

    ' Write # 那一枚：串该带引号（修前那一行也是一串数字，引号那条判据自己会红）
    Dim all As String
    all = ""
    fn = FreeFile
    Open "t608_out.txt" For Input As #fn
    Do While Not EOF(fn)
        Line Input #fn, rb
        all = all & rb & "|"
    Loop
    Close #fn
    Debug.Print "608-write-quoted ok=" & (InStr(all, Chr(34) & gAll & Chr(34)) > 0)

    Kill "t608_out.txt"
    Debug.Print "608-DONE"
End Sub

' 回读第 n 行 (1 起) —— 每次从头数，省掉把整个文件读进数组那一套
Private Function ReadLine(ByVal n As Long) As String
    Dim fn As Integer
    Dim i As Long
    ReadLine = ""
    fn = FreeFile
    Open "t608_out.txt" For Input As #fn
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
