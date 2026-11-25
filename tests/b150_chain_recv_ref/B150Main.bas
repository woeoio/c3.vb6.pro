Option Explicit

' 账 #305 §11.131 (第一百五十刀): 链写**接收者**那一枚临时引用。
'
' `vb6_ComGetObjectProp(o, L"Child")` 交回的那一枚归调用方（§11.126），而 `Set ... .Prop = v`
' 这一站只拿它递一次派发、没人放 ⇒ 每发链写漏一枚：中间那枚对象永不 Terminate。
' 改前发码（一条语句，无临时量）：
'     vb6_ComSetRef(vb6_ComGetObjectProp(p, L"Child"), L"Held", it);
' 改后：
'     void* _rcv0 = vb6_ComGetObjectProp(p, L"Child");
'     vb6_ComSetRef(_rcv0, L"Held", it);
'     vb6_ReleaseObject((void**)&_rcv0);
'
' 尺子是 Terminate 计数（探针 .temp/k150/p2 把读链那一路分出去）：
'   A2: 写进去的确实是 it 那一枚（改 h.Tag 后 it.Tag 跟着变）
'   Z : 两枚 box（根 + 子）都该 Terminate = 2；改前停在 1（子对象被那一枚夹住）
'   Z-item: 子对象放掉 mHold 之后 item 也 Terminate = 1
'
' 刻意**写完之后**才取子对象（验收用），不在链写期间另外握一枚 —— 那个形状另有读数：
' 同一枚子对象若在链写时还握在别的变量里，链写之后仍多留一枚（在册，见 §11.132）。

Public gBoxTerms As Long
Public gItemTerms As Long
Public gItemMade As Long

Sub Main()
    Dim p As Object
    Set p = New B150Box
    p.Tag = 5

    Dim it As Object
    Set it = New B150Item
    it.Tag = 7

    Set p.Child.Held = it        ' ← 靶

    Dim c As Object
    Set c = p.Child              ' 写完之后才取（单层，不碰读链那一格）
    Dim h As Object
    Set h = c.Held
    If h Is Nothing Then
        Debug.Print "B150-A1-held=empty"
    Else
        h.Tag = 9
        Dim t As Long
        t = it.Tag
        Debug.Print "B150-A2-held-tag=" & t
    End If

    Set h = Nothing
    Set c = Nothing
    Set it = Nothing
    Debug.Print "B150-C1-item-terms=" & gItemTerms
    Set p = Nothing
    Debug.Print "B150-Z-box-terms=" & gBoxTerms
    Debug.Print "B150-Z-item-made=" & gItemMade
    Debug.Print "B150-Z-item-terms=" & gItemTerms
    Debug.Print "B150-DONE"
End Sub
