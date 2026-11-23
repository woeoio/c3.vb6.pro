Option Explicit

Public gTerms As Long
Public gBox As B135Box

' 账 #305 §11.117 (第一百三十五刀): `With New <工程类>` 的载体在 End With / Exit 出块时放手。
' 这一档发的是 vb6_cls_X_New(), 用的是类自己的 vb6_cls_X_Release (不是 vb6_ReleaseObject:
' 裸工程实例在 Fix 191 那道判据里只清槽)。修前 A2/B2/C2/Z 全 0。

' A: 正常走到 End With。
Private Sub NewCarrier()
    With New B135Item
        .Tag = 11
        Debug.Print "B135-A1-tag=" & .Tag
    End With
    Debug.Print "B135-A2-terms=" & gTerms
End Sub

' B: Exit Sub 从 With New <工程类> 体内跳出。
Private Sub NewCarrierExit()
    With New B135Item
        .Tag = 22
        Debug.Print "B135-B1-tag=" & .Tag
        Exit Sub
    End With
    Debug.Print "B135-B1b-unreachable"
End Sub

' C: 反面证人 —— With <字段里的实例> 借来的那一份, 两块口都不许退。
Private Sub FieldBorrow()
    With gBox.It
        .Tag = 33
        Debug.Print "B135-C1-tag=" & .Tag
    End With
    Debug.Print "B135-C2-terms=" & gTerms
End Sub

' D: 反面证人二 —— With <局部> + Exit Sub, 局部那一份归过程出口 (§11.112 那档), 不归 With。
Private Sub LocalBorrowExit()
    Dim it As B135Item
    Set it = gBox.It
    With it
        .Tag = 44
        Debug.Print "B135-E1-tag=" & .Tag
        Exit Sub
    End With
End Sub

Sub Main()
    Set gBox = New B135Box
    Set gBox.It = New B135Item
    NewCarrier
    NewCarrierExit
    Debug.Print "B135-B2-terms=" & gTerms
    FieldBorrow
    LocalBorrowExit
    Debug.Print "B135-E2-terms=" & gTerms
    Debug.Print "B135-Z-terms=" & gTerms
    Debug.Print "B135-D1-field=" & gBox.It.Tag
    Debug.Print "B135-DONE"
End Sub