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
    Debug.Print "B448-DONE"
End Sub
