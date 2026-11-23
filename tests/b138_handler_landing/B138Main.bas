Option Explicit

' 账 #305 §11.119 (第一百三十九刀): 错误落点的活载体名册。
' 只有"本过程一句 Resume 都没有"时才启用 —— 那种 handler 回不到块内, 落点可以逐枚放掉。

Public gTerms As Long
Public gColl As Collection

Private Function MakeColl() As Collection
    Set MakeColl = New Collection
End Function

' A: With New Collection, 句柄没有 Resume —— 落点放手。
Private Sub A_Landing()
    On Error GoTo H
    With New Collection
        .Add New B138Item
        Debug.Print "B138-A1-count=" & .Count
        Error 5000
    End With
H:
    Debug.Print "B138-A2-terms=" & gTerms
End Sub

' B: 同上, handler 里 Exit Sub —— 同样回不到块内, 落点放手。
Private Sub B_LandingExit()
    On Error GoTo H
    With New Collection
        .Add New B138Item
        Debug.Print "B138-B1-count=" & .Count
        Error 5000
    End With
H:
    Debug.Print "B138-B2-terms=" & gTerms
    Exit Sub
End Sub

' C: 反面证人 —— With <模块字段> 借来的那一枚, 落点一枚都不许动。
Private Sub C_Borrowed()
    On Error GoTo H
    With gColl
        .Add New B138Item
        Debug.Print "B138-C1-count=" & .Count
        Error 5000
    End With
H:
    Debug.Print "B138-C2-count=" & gColl.Count
    Debug.Print "B138-C3-terms=" & gTerms
End Sub

' D: 两个 With 先后 —— 前一个块尾退, 后一个落点退 (名册槽位互不串)。
Private Sub D_TwoBlocks()
    On Error GoTo H
    With New Collection
        .Add New B138Item
    End With
    With New Collection
        .Add New B138Item
        Error 5000
    End With
H:
    Debug.Print "B138-D2-terms=" & gTerms
End Sub

' E: §11.118 的工程函数那一枚, 也是名册里的客人。
Private Sub E_FnRet()
    On Error GoTo H
    With MakeColl()
        .Add New B138Item
        Debug.Print "B138-E1-count=" & .Count
        Error 5000
    End With
H:
    Debug.Print "B138-E2-terms=" & gTerms
End Sub

Sub Main()
    Set gColl = New Collection
    gColl.Add New B138Item
    A_Landing
    B_LandingExit
    C_Borrowed
    D_TwoBlocks
    E_FnRet
    Debug.Print "B138-Z-terms=" & gTerms
    Debug.Print "B138-Z-count=" & gColl.Count
    Debug.Print "B138-DONE"
End Sub