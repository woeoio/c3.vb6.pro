Option Explicit

Public gMade As Long
Public gTerms As Long

Public Sub Main()
    Dim it As B155Item
    Dim box As B155Box
    Dim o As Object

    Set it = New B155Item
    it.Tag = 11
    Set box = New B155Box
    Set o = box

    ' P1: 早绑定置空 (box.Item = Nothing) —— 写口那一站
    Set box.Item = it
    Debug.Print "B155-P1-held=" & box.Held
    Set box.Item = Nothing
    Debug.Print "B155-P1-cleared=" & box.ItemIsNothing

    ' P2: 晚绑定置空 (o.Item = Nothing) —— 经 COM 派发, 对照组
    Set box.Item = it
    Set o.Item = Nothing
    Debug.Print "B155-P2-cleared=" & box.ItemIsNothing

    ' P3: 置空后再写, 槽还能用
    Set box.Item = it
    Debug.Print "B155-P3-held=" & box.Held

    ' P4: 放掉调用方那一枚, 盒子还握着
    Set it = Nothing
    Debug.Print "B155-P4-nothing=" & box.ItemIsNothing

    ' Z: 盒子放手 ⇒ 字段出厂放手 ⇒ item Terminate (置空那一站没留下多算的引用)
    Set box = Nothing
    Set o = Nothing
    Debug.Print "B155-Z-terms=" & gTerms
    Debug.Print "B155-Z-made=" & gMade
    Debug.Print "B155-DONE"
End Sub
