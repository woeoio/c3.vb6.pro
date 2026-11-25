Option Explicit

Public gTerms As Long
Public gBox As B134Box

' 账 #305 §11.116 (第一百三十四刀): 从 With 体内 Exit 出去, 本块自己造的那一枚要先放手。
' §11.115 只发了块尾那一句; 体内 Exit 走 goto, 一出块就没人再问它 ⇒ 修前 F2/H2/I2/Z 全 0。

' A: 对照组 —— 正常走到 End With (第一百三十三刀那一发)。
Private Sub WithNew()
    With New Collection
        .Add New B134Item
        Debug.Print "B134-A1-count=" & .Count
    End With
    Debug.Print "B134-A2-terms=" & gTerms
End Sub

' F: Exit Sub 从 With New 体内跳出。
Private Sub ExitInside()
    With New Collection
        .Add New B134Item
        Debug.Print "B134-F1-count=" & .Count
        Exit Sub
    End With
    Debug.Print "B134-F1b-unreachable"
End Sub

' G: 反面证人 —— With <字段> 里 Exit Sub, 不许退字段手里那一份。
Private Sub ExitInsideField()
    Set gBox.C = New Collection
    With gBox.C
        .Add New B134Item
        Debug.Print "B134-G1-count=" & .Count
        Exit Sub
    End With
    Debug.Print "B134-G1b-unreachable"
End Sub

' H: Exit Function 从 With New 体内跳出 (返回值仍要落位)。
Private Function ExitInsideFunc() As Long
    With New Collection
        .Add New B134Item
        ExitInsideFunc = 7
        Exit Function
    End With
    ExitInsideFunc = 99
End Function

' I: 两层自家载体 + Exit Sub —— 内层先退、外层后退, 两枚都要退。
Private Sub ExitNested()
    With New Collection
        .Add New B134Item
        With New Collection
            .Add New B134Item
            Exit Sub
        End With
    End With
    Debug.Print "B134-I1b-unreachable"
End Sub

Sub Main()
    Set gBox = New B134Box
    WithNew
    ExitInside
    Debug.Print "B134-F2-terms=" & gTerms
    ExitInsideField
    Debug.Print "B134-G2-terms=" & gTerms
    Debug.Print "B134-H1=" & ExitInsideFunc
    Debug.Print "B134-H2-terms=" & gTerms
    ExitNested
    Debug.Print "B134-I2-terms=" & gTerms
    Debug.Print "B134-Z-terms=" & gTerms
    Debug.Print "B134-DONE"
End Sub