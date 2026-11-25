Option Explicit

' 账 #305 §11.119 四.4 (第一百四十二刀): 工程类载体进错误落点名册。
' 名册槽里只有 void**, 落点分不出类名 —— 登记时把类自己的 vb6_cls_X_Release
' 地址一并存进 _vb6_ownRel[k], 落点按它放手。

Public gTerms As Long
Public gColl As Collection

' A: 一枚工程类载体, 落点放手。
Private Sub A_ClassCarrier()
    On Error GoTo H
    With New B141Item
        .Tag = 1
        Error 5000
    End With
    Exit Sub
H:
    Debug.Print "B141-A1-terms=" & gTerms
End Sub

' B: 混合 —— 前一块是晚绑定载体 (块尾已退), 后一块是工程类 (落点退)。
Private Sub B_Mixed()
    On Error GoTo H
    With New Collection
        .Add New B141Item
    End With
    With New B141Item
        .Tag = 2
        Error 5000
    End With
    Exit Sub
H:
    Debug.Print "B141-B1-terms=" & gTerms
End Sub

' C: 反面证人 —— With gColl 借来的那一帧不进名册, 落点一枚都不许动。
Private Sub C_Borrowed()
    On Error GoTo H
    With gColl
        .Add New B141Item
        Error 5000
    End With
    Exit Sub
H:
    Debug.Print "B141-C1-count=" & gColl.Count
End Sub

' D: 两枚工程类载体同时在册 (嵌套), 落点逐槽退。
Private Sub D_TwoClass()
    On Error GoTo H
    With New B141Item
        .Tag = 3
        With New B141Item
            .Tag = 4
            Error 5000
        End With
    End With
    Exit Sub
H:
    Debug.Print "B141-D1-terms=" & gTerms
End Sub

Sub Main()
    Set gColl = New Collection
    gColl.Add New B141Item
    A_ClassCarrier
    B_Mixed
    C_Borrowed
    D_TwoClass
    Debug.Print "B141-Z-terms=" & gTerms
    Debug.Print "B141-Z-count=" & gColl.Count
    Debug.Print "B141-DONE"
End Sub