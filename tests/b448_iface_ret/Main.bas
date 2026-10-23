Option Explicit

' 账 #305 §B221/#449 (§B222): 「这一枚表达式交出来的是视图还是实例」—— 四个站点同台。
' 修复前三形都在编译期就红 (cl C2039 / C2198), 第四形 (被调方函数体内的 Set) 交出去的
' 值本身就是坏的: 类实例指针被直接塞进接口视图槽。

Public Function MakeNew() As IB448
    Set MakeNew = New CB448
End Function

Public Function MakeLocal() As IB448
    Dim o As IB448
    Set o = New CB448
    Set MakeLocal = o
End Function

Public Sub TakeByVal(ByVal o As IB448)
    Debug.Print "P3-arg twice=" & o.Twice(30)
End Sub
' 账 #305 §B218 第一百三十七刀: 形参是工程类 / 视图的**实参那一站**。移交值 (`New <类>` 与被调方
' 返回槽交回来的那一枚) 自带一枚对价, 而修前没有任何一站退它 => 每调用净漏一枚, 那一枚永不销毁。
' W23 = 五枚 `New` 现造现交; W24 = 同一枚借来的实例分别当 ByVal 与 ByRef 实参 (反面证人:
' 两头都不许多退, 只有过程自己那枚局部在出口退); W25 = 视图那一档的移交值 (与 P3 同形, 循环四次)。
Public Function MakeCls48() As CB448
    Set MakeCls48 = New CB448
End Function

Public Sub TakeCls48(ByVal c As CB448)
    c.Twice 1
End Sub

Public Sub TakeCls48Ref(c As CB448)
    c.Twice 2
End Sub

Public Sub TakeView48(ByVal v As IB448)
    v.Twice 3
End Sub

Sub W23ArgHandOver()
    Dim k23 As Long
    For k23 = 1 To 5
        TakeCls48 New CB448
    Next
End Sub

Sub W24ArgBorrow()
    Dim c24 As CB448
    Set c24 = New CB448
    TakeCls48 c24
    TakeCls48Ref c24
    Debug.Print "W24 hits=" & c24.Hits()
End Sub

Sub W25ArgViewHandOver()
    Dim k25 As Long
    For k25 = 1 To 4
        TakeView48 MakeNew()
    Next
End Sub


' 账 #305 §B218 第一百三十三刀: `Static` 那一档的工程接口视图。视图腿登记「出口放手」时缺了
' 工程类那两腿都有的 `!var.isStatic && !inStaticProc_` 那道闸 ⇒ 槽 (C 侧 static 存储) 跨调用还
' 活着, 出口却把本站 QueryInterface 那一枚退掉: 第二次出口就退到 0, 第三次进来读的是已释放块。
' gMark48 = 只算这一格自己的命中 (前面 P1..P5 也吃 Twice)。
Public gHits48 As Long
Public gTerm48 As Long
Public gMark48 As Long

Sub W20StaticView()
    Static sv20 As IB448
    If sv20 Is Nothing Then
        Set sv20 = MakeNew()
    End If
    Debug.Print "W20 twice=" & sv20.Twice(2) & " hits=" & (gHits48 - gMark48) & " term=" & gTerm48
End Sub

' 账 #305 §B218 第一百三十四刀: 视图腿② (`Set <工程接口槽> = <返回该接口的工程内过程>`) 那一站
' 铸的临时量 __ivrhsN 自带被调方移交来的那一枚引用 (它的返回槽在 腿④ 不退), 而本站 QI 完只退旧槽值、
' 从不交回它 ⇒ 每调用净 +1。W21 = 一枚普通 Dim 视图槽被同一枚工厂喂 5 次: 每次覆写退掉槽里那一枚,
' 修前那枚接管值留下 ⇒ 五枚实例一枚都不销毁; 修后当场交回 ⇒ 五枚都销毁。
' W22 = 借来的那一档 (`Set dst = src`, 右值是裸视图变量): 不许交回 —— 交回就是多退一枚。
Sub W21ViewLeak()
    Dim v21 As IB448
    Dim k21 As Long
    For k21 = 1 To 5
        Set v21 = MakeNew()
    Next
End Sub

Sub W22ViewBorrow()
    Dim src22 As IB448
    Dim dst22 As IB448
    Set src22 = MakeNew()
    Set dst22 = src22
    Debug.Print "W22 twice=" & dst22.Twice(7) & " src=" & src22.Twice(1)
End Sub

Sub Main()
    Dim s As IB448
    Set s = MakeNew()
    Debug.Print "P1-setslot twice=" & s.Twice(21)
    Debug.Print "P2-inline twice=" & MakeLocal().Twice(21)
    TakeByVal MakeNew()
    Dim c As CB448
    Set c = New CB448
    Debug.Print "P4-plain twice=" & c.Twice(5)
    Debug.Print "P5-hits=" & c.Hits()
    Set s = Nothing
    gMark48 = gHits48
    gTerm48 = 0
    W20StaticView
    W20StaticView
    W20StaticView
    Debug.Print "W20-end term=" & gTerm48
    gTerm48 = 0
    W21ViewLeak
    Debug.Print "W21-end term=" & gTerm48
    W22ViewBorrow
    Debug.Print "W22-end term=" & gTerm48
    gTerm48 = 0
    W23ArgHandOver
    Debug.Print "W23-end term=" & gTerm48
    gTerm48 = 0
    W24ArgBorrow
    Debug.Print "W24-end term=" & gTerm48
    gTerm48 = 0
    W25ArgViewHandOver
    Debug.Print "W25-end term=" & gTerm48
    Debug.Print "B448-DONE"
End Sub
