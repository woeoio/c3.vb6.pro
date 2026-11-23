Option Explicit

' 账 #305 §11.116 五 (第一百四十一刀): GoTo 跳出 With 块时, 本块自己造的载体要先退。
' 跳出去之后块尾那句放手永远跑不到 —— 不退就活到进程结束。

Public gTerms As Long
Public gColl As Collection

' A: 晚绑定自有载体, GoTo 跳出。
Private Sub A_GoToOut()
    With New Collection
        .Add New B140Item
        GoTo ADone
    End With
ADone:
    Debug.Print "B140-A2-terms=" & gTerms
End Sub

' B: 两层嵌套, GoTo 一次跳出两层 —— 逐层退。
Private Sub B_Nested()
    With New Collection
        .Add New B140Item
        With New Collection
            .Add New B140Item
            GoTo BDone
        End With
    End With
BDone:
    Debug.Print "B140-B2-terms=" & gTerms
End Sub

' C: 对照 —— 标签在同一个 With 体内, 一层都不出, 一枚都不许退。
Private Sub C_InnerLabel()
    With New Collection
        .Add New B140Item
        GoTo CIn
        Debug.Print "B140-C0-not-printed"
CIn:
        Debug.Print "B140-C1-terms=" & gTerms
    End With
    Debug.Print "B140-C2-terms=" & gTerms
End Sub

' D: 工程类载体 (§11.117), GoTo 跳出。
Private Sub D_Class()
    With New B140Item
        .Tag = 1
        GoTo DDone
    End With
DDone:
    Debug.Print "B140-D2-terms=" & gTerms
End Sub

' E: 反面证人 —— With gColl 借来的那一枚, GoTo 出去不许退。
Private Sub E_Borrowed()
    With gColl
        .Add New B140Item
        GoTo EDone
    End With
EDone:
    Debug.Print "B140-E2-count=" & gColl.Count
End Sub

Sub Main()
    Set gColl = New Collection
    gColl.Add New B140Item
    A_GoToOut
    B_Nested
    C_InnerLabel
    D_Class
    E_Borrowed
    Debug.Print "B140-Z-terms=" & gTerms
    Debug.Print "B140-Z-count=" & gColl.Count
    Debug.Print "B140-DONE"
End Sub