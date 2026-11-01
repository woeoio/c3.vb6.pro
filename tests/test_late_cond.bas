' test_late_cond.bas - 账 #305 第二百四十四刀: 「晚绑定调用的结果当 VB6 的真值用」
' 病灶: 裸 If / ElseIf / Do While / Do Until / Loop While / Loop Until / While...Wend 那七站
'       从前发成 if (vb6_ComCall(...)) —— 测的是返回的那枚 VARIANT 指针, 恒真。
' 两头都钉: 同一个成员分别用「在场的键 A」(真) 与「缺席的键 K-…」（假）各跑一遍,
'          按 VB6 的循环语义数轮数: While 假=0 轮 / Until 真=0 轮 / 后测形至少 1 轮。
Option Explicit

Sub Main()
    Dim ok As Long
    ok = 0
    Dim n As Long

    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "A", "1"

    If d.Exists("K-IF") Then
        n = 0
    Else
        ok = ok + 1
    End If

    If d.Exists("A") Then
        ok = ok + 1
    End If

    If Not d.Exists("K-NOT") Then
        ok = ok + 1
    End If

    Dim q As Long
    q = 0
    If d.Exists("K-IF") Then
        q = 1
    ElseIf d.Exists("K-ELSEIF") Then
        q = 2
    Else
        q = 3
    End If
    If q = 3 Then
        ok = ok + 1
    End If

    n = 0
    Do While d.Exists("K-DOWHILE")
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop
    If n = 0 Then
        ok = ok + 1
    End If

    n = 0
    Do Until d.Exists("A")
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop
    If n = 0 Then
        ok = ok + 1
    End If

    n = 0
    Do Until d.Exists("K-DOUNTIL")
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop
    If n = 6 Then
        ok = ok + 1
    End If

    n = 0
    Do
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop While d.Exists("K-LOOPWHILE")
    If n = 1 Then
        ok = ok + 1
    End If

    n = 0
    Do
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop While d.Exists("A")
    If n = 6 Then
        ok = ok + 1
    End If

    n = 0
    Do
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop Until d.Exists("A")
    If n = 1 Then
        ok = ok + 1
    End If

    n = 0
    Do
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop Until d.Exists("K-LOOPUNTIL")
    If n = 6 Then
        ok = ok + 1
    End If

    n = 0
    While d.Exists("K-WHILE")
        n = n + 1
    Wend
    If n = 0 Then
        ok = ok + 1
    End If

    ' 第二百四十七刀补的四臂: Not 那一支从前发成 ~(int32_t)(裸调用) = 拿堆 VARIANT 指针截成
    ' int32 再取反 => 恒非零。真值位与值位两头都要钉, 而条件里带 Not 的那一头 244 修不到。
    If Not d.Exists("A") Then
        q = 1
    Else
        ok = ok + 1
    End If
    Dim x As Long
    x = Not d.Exists("A")
    If x = 0 Then
        ok = ok + 1
    End If
    x = Not d.Exists("K-VAL")
    If x = -1 Then
        ok = ok + 1
    End If
    n = 0
    Do While Not d.Exists("A")
        n = n + 1
        If n > 5 Then
            Exit Do
        End If
    Loop
    If n = 0 Then
        ok = ok + 1
    End If
    Debug.Print "COND-1:COUNT ok=" & ok & "/16"
    If ok = 16 Then
        Debug.Print "COND:16/16"
    Else
        Debug.Print "COND:FAIL " & ok
    End If

    Set d = Nothing
End Sub
