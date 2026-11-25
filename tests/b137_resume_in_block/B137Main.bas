Option Explicit

' 账 #305 §11.120 (第一百三十八刀): Resume Next 在块内的落点位置。
' Resume 的语义是"回到出错语句的下一条", 与嵌套深浅无关 —— 以前只有过程顶层发 resume 点,
' 块内报错时 Resume Next 落到块尾之后 (块内剩下的语句被跳过, 块尾那句放手也被跳过)。

Public gTerms As Long
Public gFlag As Boolean

' A: 在册边界 —— handler 里一句 Resume 都没有 (控制权被带出块, §11.119 那一档还没落码)。
Private Sub A_NoResume()
    On Error GoTo H
    With New Collection
        .Add New B137Item
        Debug.Print "B137-A1-count=" & .Count
        Error 5000
    End With
    Debug.Print "B137-A1c-unreachable"
H:
    Debug.Print "B137-A2-terms=" & gTerms
End Sub

' B: 块内报错、handler Resume Next —— 回到出错语句的下一条 (仍在块内), 块尾照常放手。
Private Sub B_ResumeNextInWith()
    On Error GoTo H
    With New Collection
        .Add New B137Item
        Debug.Print "B137-B1-count=" & .Count
        Error 5000
        Debug.Print "B137-B1b-after-error"
    End With
    Debug.Print "B137-B2-terms=" & gTerms
    Exit Sub
H:
    Resume Next
End Sub

' C: 同一条错, 但嵌在 If 里再套 With (两层) —— 落点与深浅无关。
Private Sub C_ResumeNextDeep()
    On Error GoTo H
    If gFlag Then
        With New Collection
            .Add New B137Item
            Debug.Print "B137-C1-count=" & .Count
            Error 5000
            Debug.Print "B137-C1c-after-error"
        End With
    End If
    Debug.Print "B137-C2-terms=" & gTerms
    Exit Sub
H:
    Resume Next
End Sub

Sub Main()
    gFlag = True
    A_NoResume
    B_ResumeNextInWith
    C_ResumeNextDeep
    Debug.Print "B137-Z-terms=" & gTerms
    Debug.Print "B137-DONE"
End Sub