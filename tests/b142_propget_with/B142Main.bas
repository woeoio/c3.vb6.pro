Option Explicit

' 账 #305 §11.124 (第一百四十三刀): With <对象>.<Property Get> 交回的载体, 块尾/Exit/落点要退。

Public gTerms As Long
Public gBox As B142Box

' A: 属性取回来的那一枚, 走块尾放手。
Private Sub A_PropGet()
    With gBox.Fresh
        .Tag = 1
    End With
    Debug.Print "B142-A2-terms=" & gTerms
End Sub

' B: 反面证人 —— 对象字段 (`With gBox.Bare`) 只是借, 块尾不许退, 之后还要读得出来。
Private Sub B_Field()
    With gBox.Bare
        .Tag = 2
    End With
    Debug.Print "B142-B2-tag=" & gBox.Bare.Tag
End Sub

' C: 属性载体 + 体内 Exit Sub —— §11.116 的 Exit 口要退。
Private Sub C_Exit()
    With gBox.Fresh
        .Tag = 3
        Exit Sub
    End With
End Sub

' D: 属性载体 + handler 无 Resume —— §11.119 四.4 的落点要按类自己的 Release 退。
Private Sub D_Landing()
    On Error GoTo H
    With gBox.Fresh
        .Tag = 4
        Error 5000
    End With
    Exit Sub
H:
    Debug.Print "B142-D1-terms=" & gTerms
End Sub

Sub Main()
    Set gBox = New B142Box
    Set gBox.Bare = New B142Item
    A_PropGet
    B_Field
    C_Exit
    Debug.Print "B142-C1-terms=" & gTerms
    D_Landing
    Debug.Print "B142-Z-terms=" & gTerms
    Debug.Print "B142-DONE"
End Sub