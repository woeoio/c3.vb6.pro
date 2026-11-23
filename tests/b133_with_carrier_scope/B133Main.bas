Option Explicit

Public gTerms As Long
Public gBox As B133Box

' 账 #305 §11.115 (第一百三十三刀): With 自己造的那一枚载体, `End With` 放手。
' 修前四条读数 A2/B2/C2/D2 与 Z 全 0 —— 载体连同体内元素活到进程结束。

' A: With New Collection —— 发码 vb6_NewObject, 所有权交过来, 块尾退一枚。
Private Sub WithNew()
    With New Collection
        .Add New B133Item
        Debug.Print "B133-A1-count=" & .Count
    End With
    Debug.Print "B133-A2-terms=" & gTerms
End Sub

' B: With CreateObject(...) —— 发码 vb6_CreateObject, 同上。
Private Sub WithCreate()
    With CreateObject("Scripting.Dictionary")
        .Add "k", New B133Item
        Debug.Print "B133-B1-count=" & .Count
    End With
    Debug.Print "B133-B2-terms=" & gTerms
End Sub

' C: 反面证人一 —— With 一个**字段** (借来的那一份, 块尾不许退)。
Private Sub WithField()
    Set gBox.C = New Collection
    With gBox.C
        .Add New B133Item
        Debug.Print "B133-C1-count=" & .Count
    End With
    Debug.Print "B133-C2-terms=" & gTerms
End Sub

' D: 反面证人二 —— With 一个**局部** (Set 已取过引用; 那一份归局部, 不归 With)。
Private Sub WithLocal()
    Dim c As Collection
    Set gBox.D = New Dictionary
    gBox.D.Add "k", New B133Item
    Set c = gBox.C
    With c
        .Add New B133Item
        Debug.Print "B133-D1-count=" & .Count
    End With
    Debug.Print "B133-D2-terms=" & gTerms
End Sub

' E: 在册边界 —— With New <工程类> 走的是类自己的 _New/_Release 那一档, 不是晚绑定两座工厂,
'    这一格本刀不发 (tempType 是 vb6_cls_*, 闸把它挡在外面)。
Private Sub WithProjectClassNew()
    With New B133Item
        .Tag = 11
        Debug.Print "B133-E1-tag=" & .Tag
    End With
    Debug.Print "B133-E2-terms=" & gTerms
End Sub

Sub Main()
    Set gBox = New B133Box
    WithNew
    WithCreate
    WithField
    WithLocal
    WithProjectClassNew
    Debug.Print "B133-Z-terms=" & gTerms
    Debug.Print "B133-DONE"
End Sub