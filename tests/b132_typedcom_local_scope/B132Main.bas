Option Explicit

Public gTerms As Long
Public gBox As B132Box

' 腿④ 第三档: 类型化 COM 接口局部 (C 侧 vb6_ComIface_I<X>*) 出作用域放手。
' 第一路: 别名接口名 + New (所有权直接交过来)。
Private Sub ViaAlias()
    Dim d As IDictionary
    Set d = New Dictionary
    d.Add "k", New B132Item
    Debug.Print "B132-A1-count=" & d.Count
End Sub

' 第二路: CreateObject (C 侧 vb6_ComCreateTyped, 所有权交过来)。
Private Sub ViaCreate()
    Dim d As IDictionary
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", New B132Item
    Debug.Print "B132-C1-count=" & d.Count
End Sub

' 反面证人一: 从字段借 (Set 那一发 AddRef), 出口只退自己那一份。
Private Sub ViaField()
    Dim d As IDictionary
    Set d = gBox.D
    Debug.Print "B132-F1-count=" & d.Count
End Sub

' 反面证人二: Static 槽跨调用活着。
Private Sub StaticSlot()
    Static d As IDictionary
    Static n As Long
    If n = 0 Then
        n = 1
        Set d = New Dictionary
        d.Add "k", New B132Item
    End If
    Debug.Print "B132-S1-count=" & d.Count
End Sub

' 反面证人三: 局部交给字段之后, 出口那一发不许把字段手里的东西带走。
Private Sub AssignOut()
    Dim d As IDictionary
    Set d = New Dictionary
    d.Add "k", New B132Item
    Set gBox.K = d
End Sub

' 真外部类型库服务器 (scrrun): 出口那一发不许把进程带崩。
Private Sub ExternalServer()
    Dim f As FileSystemObject
    Set f = CreateObject("Scripting.FileSystemObject")
    Debug.Print "B132-X1-name=" & f.GetFileName("C:\Windows\win.ini")
End Sub

Sub Main()
    ViaAlias
    Debug.Print "B132-A2-terms=" & gTerms
    ViaCreate
    Debug.Print "B132-C2-terms=" & gTerms
    Set gBox = New B132Box
    Set gBox.D = New Dictionary
    gBox.D.Add "k", New B132Item
    ViaField
    Debug.Print "B132-F2-field=" & gBox.D.Count
    StaticSlot
    StaticSlot
    Debug.Print "B132-S2-terms=" & gTerms
    AssignOut
    Debug.Print "B132-K1-alive=" & gBox.K.Count
    Set gBox.K = Nothing
    Set gBox.D = Nothing
    Debug.Print "B132-Z-terms=" & gTerms
    ExternalServer
    Debug.Print "B132-DONE"
End Sub