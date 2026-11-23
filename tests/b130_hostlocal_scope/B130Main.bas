Option Explicit

Public gTerms As Long
Public gBox As B130Box

' 腿④ 宿主类局部档: 载体自己计数 (第一百二十九刀) + 出口那一句 (第一百三十刀)。
Private Sub NewLocal()
    Dim c As New Collection
    c.Add New B130Item
    Debug.Print "B130-E1-count=" & c.Count
End Sub

' 反面证人: 局部是从**字段**借来的引用 (Set 那一发会 AddRef), 出口放手只退自己那一份,
' 字段手里那一枚不许跟着走。
Private Sub Aliased()
    Dim c As Collection
    Set c = gBox.Col
    Debug.Print "B130-A1-count=" & c.Count
End Sub

' 反面证人: Static 槽跨调用活着, 出口不许放手 (下一次进来还要读它)。
Private Sub StaticSlot()
    Static s As Collection
    Static n As Long
    If n = 0 Then
        n = 1
        Set s = New Collection
        s.Add New B130Item
    End If
    Debug.Print "B130-S1-count=" & s.Count
End Sub

Sub Main()
    NewLocal
    Debug.Print "B130-E2-terms=" & gTerms
    Set gBox = New B130Box
    Set gBox.Col = New Collection
    gBox.Col.Add New B130Item
    Aliased
    Debug.Print "B130-A2-field=" & gBox.Col.Count
    StaticSlot
    StaticSlot
    Debug.Print "B130-S2-terms=" & gTerms
    Debug.Print "B130-DONE"
End Sub