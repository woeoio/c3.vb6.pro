Option Explicit

Public gTerms As Long
Public gBox As B131Box

' 腿④ 宿主类局部档第二格: Dim dc As Dictionary 出作用域放手 (C 侧是类型化接口)。
Private Sub NewLocal()
    Dim dc As Dictionary
    Set dc = New Dictionary
    dc.Add "k", New B131Item
    Debug.Print "B131-E1-count=" & dc.Count
End Sub

' 反面证人一: 局部是从字段借的 (Set 那一发 AddRef), 出口只退自己那一份。
Private Sub FieldAlias()
    Dim dc As Dictionary
    Set dc = gBox.Dict
    Debug.Print "B131-A1-count=" & dc.Count
End Sub

' 反面证人二: Static 槽跨调用活着, 出口不许放手。
Private Sub StaticSlot()
    Static d As Dictionary
    Static n As Long
    If n = 0 Then
        n = 1
        Set d = New Dictionary
        d.Add "k", New B131Item
    End If
    Debug.Print "B131-S1-count=" & d.Count
End Sub

' 反面证人三: 局部交给字段之后, 出口那一发不许把字段手里的东西带走。
Private Sub AssignOut()
    Dim dc As Dictionary
    Set dc = New Dictionary
    dc.Add "k", New B131Item
    Set gBox.Dict = dc
End Sub

Sub Main()
    NewLocal
    Debug.Print "B131-E2-terms=" & gTerms
    Set gBox = New B131Box
    Set gBox.Dict = New Dictionary
    gBox.Dict.Add "k", New B131Item
    FieldAlias
    Debug.Print "B131-A2-field=" & gBox.Dict.Count
    StaticSlot
    StaticSlot
    AssignOut
    Debug.Print "B131-O1-alive=" & gBox.Dict.Count
    Set gBox.Dict = Nothing
    Debug.Print "B131-Z-terms=" & gTerms
    Debug.Print "B131-DONE"
End Sub