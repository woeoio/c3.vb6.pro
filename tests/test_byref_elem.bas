' test_byref_elem.bas - 账 #305 第二百五十三刀: ByRef 形参吃「数组元素」那一形
' 病灶: 从前把元素当前值拷进一枚 BSTR 复合字面量再取址递出去 (形状见台账 §B164, 别在本文件里
'       照抄那串字节 —— 判据的反面针要靠它在产物里数), 于是被调方那句"先退旧再存新"退掉的
'       就是数组元素自己那一枚串: 写回落进没人看的临时量 (VB6 该填进 arr(i)) + 元素悬空
'       (下次读它 use-after-free、下次写它双放 = 堆损坏同形)。
' 两头都钉: 正面 = 写回真的落进元素; 反面 = 同一格反复调用不许把同一枚串放两次。
Option Explicit

Private Sub FillStr(ByRef s As String)
    s = "FILLED"
End Sub

Private Sub FillNum(ByRef n As Long)
    n = n + 7
End Sub

Sub Main()
    Dim ok As Long
    ok = 0
    Dim i As Long
    Dim bad As Long

    ' W1 一维动态 String 数组元素
    Dim arrS() As String
    ReDim arrS(0 To 2) As String
    arrS(1) = "OLD"
    FillStr arrS(1)
    If arrS(1) = "FILLED" Then
        ok = ok + 1
    End If
    Debug.Print "W1=" & arrS(1)

    ' W2 二维动态 String 数组元素 (靶工程 VBFlexGrid.ctl 那一形)
    Dim arrM() As String
    ReDim arrM(0 To 1, 0 To 1) As String
    FillStr arrM(0, 1)
    If arrM(0, 1) = "FILLED" Then
        ok = ok + 1
    End If
    Debug.Print "W2=" & arrM(0, 1)

    ' W3 数值档: ByRef Long 吃一维 Long 数组元素
    Dim arrN() As Long
    ReDim arrN(0 To 2) As Long
    arrN(1) = 10
    FillNum arrN(1)
    If arrN(1) = 17 Then
        ok = ok + 1
    End If
    Debug.Print "W3=" & arrN(1)

    ' W4 变量那一形不许被本刀改坏 (本来就直递地址)
    Dim s As String
    s = "V"
    FillStr s
    If s = "FILLED" Then
        ok = ok + 1
    End If
    Debug.Print "W4=" & s

    ' W5 反面证人: 同一格反复 2000 次 —— 改前每次都把元素自己那枚串再放一次
    bad = 0
    For i = 1 To 2000
        FillStr arrS(1)
        If arrS(1) <> "FILLED" Then
            bad = bad + 1
        End If
    Next
    If bad = 0 Then
        ok = ok + 1
    End If
    Debug.Print "W5=bad " & bad

    ' W6 元素再赋一次 (退旧退自己那一枚, 不许崩)
    arrS(1) = "AGAIN"
    If arrS(1) = "AGAIN" Then
        ok = ok + 1
    End If
    Debug.Print "W6=" & arrS(1)

    Debug.Print "BRW-1:COUNT ok=" & ok & "/6"
    If ok = 6 Then
        Debug.Print "BRW:6/6"
    Else
        Debug.Print "BRW:FAIL " & ok
    End If
End Sub
