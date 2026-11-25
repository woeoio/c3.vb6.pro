Option Explicit

' 账 #305 §11.127 (第一百四十六刀): `Set x = <晚绑定对象>.<属性>` 交回的引用。
'
' 桥在 §11.126 已经对平 (适配器转移 + 服务端清 tmpRet): 一次读数恰好交出 1 枚。
' 本刀把 setlet 站那一步多余的 `vb6_ComAddRefDispatch` 撤掉 —— RTL 契约明写
' 「对象引用已转移给调用方」, 槽拿走的就是那一枚; 再 claim 一次 = 每发净 +1。
'
' A 路: 新造实例 —— Set c = Nothing 之后必须恰好 Terminate 一次 (改前 = 0)。
' B 路: 反面证人 —— 交回的是模块级共用实例, 退读数那一枚之后它必须还活着。
' C 路: **也**是反面证人, 方向相反 —— `Set d = c` 是裸槽拷贝 (拷的是持有者的那一枚),
'       那一步的 claim 因此**必须留着**, 否则 Set c = Nothing 会把 d 也带走。

Public gTerms As Long
Public gMade As Long
Public gKeep As B146Item

Sub Main()
    Set gKeep = New B146Item
    gKeep.Tag = 7

    Dim o As Object
    Set o = New B146Box

    Dim c As Object
    Dim d As Object

    ' A: 新造一枚
    Dim tA As Long
    tA = gTerms
    Set c = o.Fresh
    c.Tag = 11
    Set c = Nothing
    Debug.Print "B146-A1-terms=" & (gTerms - tA)

    ' B: 共用实例 (退读数那一枚, 实例自己那一枚还在)
    Dim tB As Long
    tB = gTerms
    Set c = o.SameObj
    Set c = Nothing
    Debug.Print "B146-B1-tag=" & gKeep.Tag
    Debug.Print "B146-B2-terms=" & (gTerms - tB)

    ' C: 裸槽拷贝仍要取引用 (方向相反的证人)
    '    判据不看属性读 (d 是 As Object, 读裸工程实例的属性会走晚绑定, 读数不可信),
    '    直接看 Terminate 计数: Set c = Nothing 之后**还该是 0** (d 手里那一枚仍在),
    '    直到 Set d = Nothing 才变 1。
    Dim tC As Long
    tC = gTerms
    Set c = o.Fresh
    c.Tag = 13
    Set d = c
    Set c = Nothing
    Debug.Print "B146-C1-alive=" & (gTerms - tC)
    Set d = Nothing
    Debug.Print "B146-C2-terms=" & (gTerms - tC)

    Debug.Print "B146-Z-terms=" & gTerms
    Debug.Print "B146-Z-made=" & gMade
    Debug.Print "B146-DONE"
End Sub
