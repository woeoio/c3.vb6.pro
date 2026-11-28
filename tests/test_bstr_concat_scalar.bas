' Test: string concatenation with scalar function results
' Covers bugfix doc 260919: A-1 (unanchored substring check), A-2 (prefix-block
' fallback -> type inference), B3 (numeric list split CStrLong / CStrDbl).
' Prints BCS markers to stdout (Debug.Print) and writes result.txt (App.Path)
' so the output can be diffed against a native VB6 build.

Option Explicit

Sub Main()
    Dim s As String
    Dim e1 As Long

    On Error Resume Next
    Err.Raise 5
    On Error GoTo 0

    ' A-2: member scalar WITHOUT string arguments (was passed through bare)
    ' §11.229 (第二百四十七刀): 真机在这里是 **a=0** —— `On Error GoTo 0` 也清 Err
    ' (VB6Mini 实测 .temp/k233p/P357 A2/A3/A5=0，同段 VBScript 亦为 0)。
    ' 旧针脚 a=5 是当年把“GoTo 0 不清”当成了真机口径。
    s = "a=" & Err.Number

    ' A-1: scalar functions WITH string arguments (substring hit bypassed them)
    s = s & vbCrLf & "b=" & Len("abcd")
    s = s & vbCrLf & "c=" & Val("2.5")
    s = s & vbCrLf & "d=" & InStr("hello", "l")

    ' B3: double/float returning functions (were truncated via CStrLong)
    s = s & vbCrLf & "e=" & Sqr(4)
    s = s & vbCrLf & "f=" & Abs(-3.5)
    s = s & vbCrLf & "g=" & Int(7.9)
    s = s & vbCrLf & "h=" & Fix(-7.1)
    s = s & vbCrLf & "i=" & CSng(1.5)

    ' B3: long returning functions (unchanged group)
    s = s & vbCrLf & "j=" & CLng(3.49)
    s = s & vbCrLf & "k=" & CInt(2.9)
    s = s & vbCrLf & "l=" & Sgn(-4)
    s = s & vbCrLf & "m=" & VarType(e1)

    ' A-2: UBound without string arguments
    Dim arr(1 To 3) As Long
    s = s & vbCrLf & "n=" & UBound(arr)

    ' plain variable / literal concatenation (unchanged paths)
    Dim t As String
    t = "txt"
    s = s & vbCrLf & "o=" & t & "!"
    s = s & vbCrLf & "p=" & "lit"

    ' result file for VB6-native diff
    Dim h As Integer
    h = FreeFile
    Open App.Path & "\result.txt" For Output As #h
    Print #h, s
    Close #h

    ' stdout markers for run_tests.ps1
    Dim okAll As Boolean
    okAll = InStr(s, "a=0") > 0
    okAll = okAll And InStr(s, "b=4") > 0
    okAll = okAll And InStr(s, "c=2.5") > 0
    okAll = okAll And InStr(s, "d=3") > 0
    okAll = okAll And InStr(s, "e=2") > 0
    okAll = okAll And InStr(s, "f=3.5") > 0
    okAll = okAll And InStr(s, "g=7") > 0
    okAll = okAll And InStr(s, "h=-7") > 0
    okAll = okAll And InStr(s, "i=1.5") > 0
    okAll = okAll And InStr(s, "j=3") > 0
    okAll = okAll And InStr(s, "k=3") > 0
    okAll = okAll And InStr(s, "l=-1") > 0
    okAll = okAll And InStr(s, "m=3") > 0
    okAll = okAll And InStr(s, "n=3") > 0
    okAll = okAll And InStr(s, "o=txt!") > 0
    okAll = okAll And InStr(s, "p=lit") > 0

    If okAll Then
        Debug.Print "BCS:16/16"
    Else
        Debug.Print "BCS:MISMATCH"
    End If
End Sub
