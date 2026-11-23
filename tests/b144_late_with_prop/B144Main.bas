Option Explicit

' 账 #305 §11.125 (第一百四十四刀 · 设计刀): `With <晚绑定对象>.<属性>` 交回的引用。
'
' 本刀**只量不改**: 四口 (块尾 / Exit / 落点 / GoTo) 今天一句放手都不发, 读数为 0;
' 原因不在 With 这一站, 而在晚绑定读数在 COM 桥里就多交了两枚 (见 ai/039 §11.125 的账)。
' 这一格读数就是下一刀修桥之后的 oracle: 桥修好、With 归位, A2/C1/D1/E1 各应变成 1/2/3/4。

Public gTerms As Long
Public gKeep As B144Item

' A: 块尾 —— 交回的那一枚应当在 `End With` 放手。
Private Sub A_BlockTail()
    Dim o As Object
    Set o = New B144Box
    With o.Fresh
        .Tag = 1
    End With
    Debug.Print "B144-A2-terms=" & gTerms
End Sub

' B: 反面证人 —— 交回的是模块级共用实例, 放手一枚之后它必须还活着、值还读得出。
Private Sub B_Shared()
    Dim o As Object
    Set o = New B144Box
    With o.SameObj
        .Tag = 8
    End With
    Debug.Print "B144-B2-tag=" & gKeep.Tag
    Debug.Print "B144-B2-terms=" & gTerms
End Sub

' C: 体内 Exit Sub —— §11.116 的 Exit 口。
Private Sub C_Exit()
    Dim o As Object
    Set o = New B144Box
    With o.Fresh
        .Tag = 3
        Exit Sub
    End With
End Sub

' D: handler 无 Resume —— §11.119 四.4 的落点。
Private Sub D_Landing()
    Dim o As Object
    On Error GoTo H
    Set o = New B144Box
    With o.Fresh
        .Tag = 4
        Error 5000
    End With
    Exit Sub
H:
    Debug.Print "B144-D1-terms=" & gTerms
End Sub

' E: GoTo 跳出 With 块 —— §11.122 的逐层退。
Private Sub E_GoToOut()
    Dim o As Object
    Set o = New B144Box
    With o.Fresh
        .Tag = 5
        GoTo EDone
    End With
EDone:
    Debug.Print "B144-E1-terms=" & gTerms
End Sub

Sub Main()
    Set gKeep = New B144Item
    gKeep.Tag = 7
    A_BlockTail
    B_Shared
    C_Exit
    Debug.Print "B144-C1-terms=" & gTerms
    D_Landing
    E_GoToOut
    Debug.Print "B144-Z-terms=" & gTerms
    Debug.Print "B144-Z-tag=" & gKeep.Tag
    Debug.Print "B144-DONE"
End Sub
