Option Explicit

Private gb As Byte
Private gc As Currency
Private gt As Date

Public Sub Main()
    Dim box As B165Box
    Dim o As Object
    Dim ba(1 To 2) As Byte
    Dim ca(1 To 2) As Currency
    Dim ta(1 To 2) As Date

    Set box = New B165Box
    Set o = box

    ' Y1: Byte 数组元素
    ba(1) = 1
    o.SetB ba(1)
    Debug.Print "B165-Y1=" & ba(1)

    ' Y2: Byte 类字段
    box.B = 2
    o.SetB box.B
    Debug.Print "B165-Y2=" & box.B

    ' Y3: Byte 模块级裸变量（基线）
    gb = 3
    o.SetB gb
    Debug.Print "B165-Y3=" & gb

    ' C1: Currency 数组元素
    ca(1) = 10
    o.IncC ca(1)
    Debug.Print "B165-C1=" & ca(1)

    ' C2: Currency 类字段
    box.C = 20
    o.IncC box.C
    Debug.Print "B165-C2=" & box.C

    ' C3: Currency 模块级裸变量（基线）
    gc = 30
    o.IncC gc
    Debug.Print "B165-C3=" & gc

    ' T1: Date 数组元素（用比较判，避开日期格式化差异）
    ta(1) = DateSerial(2020, 1, 1)
    o.NextD ta(1)
    If ta(1) = DateSerial(2020, 1, 2) Then
        Debug.Print "B165-T1=ok"
    Else
        Debug.Print "B165-T1=bad"
    End If

    ' T2: Date 类字段
    box.T = DateSerial(2021, 6, 30)
    o.NextD box.T
    If box.T = DateSerial(2021, 7, 1) Then
        Debug.Print "B165-T2=ok"
    Else
        Debug.Print "B165-T2=bad"
    End If

    ' T3: Date 模块级裸变量（基线）
    gt = DateSerial(2022, 12, 31)
    o.NextD gt
    If gt = DateSerial(2023, 1, 1) Then
        Debug.Print "B165-T3=ok"
    Else
        Debug.Print "B165-T3=bad"
    End If

    ' Y4: 同一元素连发两次
    ba(2) = 0
    o.SetB ba(2)
    o.SetB ba(2)
    Debug.Print "B165-Y4=" & ba(2)

    Debug.Print "B165-DONE"
    Set box = Nothing
    Set o = Nothing
End Sub
