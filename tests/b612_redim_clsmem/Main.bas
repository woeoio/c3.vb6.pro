Attribute VB_Name = "Main"
Option Explicit

' 账 #612 夹具: 类实例的一维动态数组成员, 从类外 ReDim / Preserve / Erase。
' 改前这一形发 `o_Digits`(从没声明) => C2065, 整台 exe 编不出来。
' 对照组两枚 (模块级数组 / UDT 成员数组) 钉「没改过头」。

Private Type THost
    Items() As Long
End Type

Dim mo As K612
Dim g() As Long

Sub Main()
    Dim o As K612
    Dim h As THost
    Dim v As Variant
    Dim w As Variant
    Dim i As Long

    Set o = New K612
    Set mo = New K612

    ' 1) 局部类实例: ReDim + 逐元素写读
    ReDim o.Digits(1 To 4)
    For i = 1 To 4
        o.Digits(i) = i * 100
    Next i
    Debug.Print "612-redim=[" & o.Digits(4) & "]"

    ' 2) Preserve: 旧值留住, 新槽可写
    ReDim Preserve o.Digits(1 To 6)
    o.Digits(6) = 600
    Debug.Print "612-preserve=[" & o.Digits(4) & "]"
    Debug.Print "612-preserve-new=[" & o.Digits(6) & "]"

    ' 3) Erase 之后载体可再用
    Erase o.Digits
    ReDim o.Digits(1 To 2)
    o.Digits(2) = 77
    Debug.Print "612-erase-reuse=[" & o.Digits(2) & "]"

    ' 4) 字符串成员 (元素档 BSTR, 不是 Long)
    ReDim o.Tags(1 To 2)
    o.Tags(2) = "T612"
    Debug.Print "612-tags=[" & o.Tags(2) & "]"

    ' 5) 类形参那一档 owner
    o.Grow o, 5
    Debug.Print "612-formal=[" & o.Digits(5) & "]"
    Debug.Print "612-sum=[" & o.Sum(o) & "]"

    ' 6) 模块级类实例
    ReDim mo.Digits(1 To 3)
    mo.Digits(3) = 33
    Debug.Print "612-module=[" & mo.Digits(3) & "]"

    ' 7) 整数组读进 Variant: 与模块级对照组同一个 VarType
    '    (刻意排在 With 那一档之前: With 块 ReDim 的元素档另有一格缺口, 见台账)
    v = o.Digits
    Debug.Print "612-vartype=[" & VarType(v) & "]"

    ' 8) With 块成员 (左值这一档改前就对, 当对照)
    With o
        ReDim .Digits(1 To 3)
        .Digits(2) = 55
    End With
    Debug.Print "612-with=[" & o.Digits(2) & "]"

    ' 对照组 A: 模块级动态数组
    ReDim g(1 To 3)
    g(2) = 77
    Debug.Print "612-ctl-mod=[" & g(2) & "]"
    w = g
    Debug.Print "612-ctl-vartype=[" & VarType(w) & "]"

    ' 对照组 B: UDT 成员数组 (改前就走 knownUdtVars_ 那条腿)
    ReDim h.Items(1 To 2)
    h.Items(2) = 88
    Debug.Print "612-ctl-udt=[" & h.Items(2) & "]"

    Debug.Print "612-DONE"
End Sub
