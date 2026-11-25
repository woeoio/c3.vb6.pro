Option Explicit

' 账 #305 §11.119 八.3 (第一百四十刀): 出错语句是块内**最后一条** + handler Resume Next。
' 这一形以前落点落到块外、块尾那句放手被跨过 —— 块尾补一个 resume 点 (补在放手**之前**) 关掉它。

Public gTerms As Long
Public gColl As Collection

' A: 晚绑定载体, 末句报错。
Private Sub A_LastStmt()
    On Error GoTo H
    With New Collection
        .Add New B139Item
        Error 5000
    End With
    Debug.Print "B139-A2-terms=" & gTerms
    Exit Sub
H:
    Resume Next
End Sub

' B: 工程类载体 (§11.117), 末句报错。
Private Sub B_LastStmtClass()
    On Error GoTo H
    With New B139Item
        .Tag = 7
        Error 5000
    End With
    Debug.Print "B139-B2-terms=" & gTerms
    Exit Sub
H:
    Resume Next
End Sub

' C: 对照 —— 报错语句后面还有一句 (第一百三十八刀那一档), 落点在块内。
Private Sub C_MidStmt()
    On Error GoTo H
    With New Collection
        .Add New B139Item
        Error 5000
        Debug.Print "B139-C1-after-error"
    End With
    Debug.Print "B139-C2-terms=" & gTerms
    Exit Sub
H:
    Resume Next
End Sub

' D: 反面证人 —— With <模块字段> 借来的那一枚, 块尾没有清理可跳, 也不许退。
Private Sub D_Borrowed()
    On Error GoTo H
    With gColl
        .Add New B139Item
        Error 5000
    End With
    Debug.Print "B139-D2-count=" & gColl.Count
    Exit Sub
H:
    Resume Next
End Sub

Sub Main()
    Set gColl = New Collection
    gColl.Add New B139Item
    A_LastStmt
    B_LastStmtClass
    C_MidStmt
    D_Borrowed
    Debug.Print "B139-Z-terms=" & gTerms
    Debug.Print "B139-Z-count=" & gColl.Count
    Debug.Print "B139-DONE"
End Sub