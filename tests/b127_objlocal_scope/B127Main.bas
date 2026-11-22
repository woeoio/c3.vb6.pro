Option Explicit

' 第一百二十七刀 (§11.109): 局部泛对象槽 (`Dim o As Object`) 在过程出口放手那一格。
Public gTerms As Long

' 甲: 现造现丢 —— 出口那一发交回包装器自己那一枚 (包装器再交回类引用) ⇒ 每调用一枚出厂。
Private Sub ViaLocal(seed As Long)
    Dim o As Object
    Dim it As B127Item
    Set it = New B127Item
    it.Seed = seed
    Set o = it
End Sub

' 乙 (反面): For Each 的循环变量每轮由循环自己写, 元素那一枚归容器 ⇒ 出口不许退。
Private Sub ViaEach(c As Collection)
    Dim oEach As Object
    For Each oEach In c
        Debug.Print "  each " & TypeName(oEach)
    Next
End Sub

' 丙: 声明住在 If 括号里 (提升到序言) 的那一格: 出口那一发要编得过, 且照样放手。
Private Sub InBlock(flag As Long, seed As Long)
    If flag <> 0 Then
        Dim o As Object
        Dim it As B127Item
        Set it = New B127Item
        it.Seed = seed
        Set o = it
    End If
End Sub

Sub Main()
    Dim i As Long
    Dim c As Collection
    Dim v As Variant
    Dim it As B127Item

    For i = 1 To 3
        ViaLocal i
    Next
    Debug.Print "C3ProbeB127-A1-terms=" & CStr(gTerms)

    Set it = New B127Item
    it.Seed = 77
    Set c = New Collection
    Set v = it
    c.Add v
    Set it = Nothing
    ViaEach c
    Debug.Print "C3ProbeB127-A2-live=" & TypeName(c.Item(1))

    InBlock 1, 9
    InBlock 0, 9
    Debug.Print "C3ProbeB127-A3-terms=" & CStr(gTerms)

    Debug.Print "C3ProbeB127-DONE"
End Sub