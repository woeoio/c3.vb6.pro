Option Explicit

Public gMade As Long
Public gTerms As Long

Public Sub Main()
    Dim it As B154Item
    Dim box As B154Box
    Dim o As Object

    Set it = New B154Item
    it.Tag = 11
    Set box = New B154Box
    Set o = box

    ' A: 早绑定写口 (box.Item = it) —— 对照组
    Set box.Item = it
    Debug.Print "B154-A1-nothing=" & box.ItemIsNothing
    Debug.Print "B154-A2-held=" & box.Held

    ' B: 晚绑定写口 (o.Item = it) —— 同一枚属性, 经 COM 派发
    Set o.Item = it
    Debug.Print "B154-B1-nothing=" & box.ItemIsNothing
    Debug.Print "B154-B2-held=" & box.Held

    ' C: 放掉调用方那一枚, 盒子还握着 ⇒ item 仍活着
    Set it = Nothing
    Debug.Print "B154-C1-nothing=" & box.ItemIsNothing
    Debug.Print "B154-C2-held=" & box.Held

    ' D: 盒子放手 ⇒ 字段出厂放手 ⇒ item Terminate
    Set box = Nothing
    Set o = Nothing
    Debug.Print "B154-Z-terms=" & gTerms
    Debug.Print "B154-Z-made=" & gMade
    Debug.Print "B154-DONE"
End Sub
