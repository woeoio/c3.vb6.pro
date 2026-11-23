Option Explicit

' 账 #305 §11.128 (第一百四十七刀): `Set <晚绑定对象>.<属性> = <对象>` 的写口。
'
' 改前两处病同时挡住这一发写:
'   (1) 分发表里 Get/Set 各拿一枚 dispid (Held=1 / Held=2), 而客户端按名查只回**第一条**
'       ⇒ 写永远打在 Get 那一格上;
'   (2) 服务端认档写成 `(ik == 4 || ik == 8) && (wFlags & DISPATCH_PROPERTYPUT)` ——
'       客户端递的是 PROPERTYPUTREF(8), `wFlags & 4` 为假 ⇒ 按档一位都对不上, 落进
'       「同 dispid 第一条」的兜底 = 读臂 ⇒ 参数整个丢掉, 返回 S_OK。
' 两头合起来的读数: 写静默无效, 而且对象当场收回 0 (W1=1 / W2=1)。
'
' W 路: Set o.Held = it —— 写进去之后 gHold 持有, it 交回自己那一枚; 清 gHold 才 Terminate。
' R 路: 同名 Let/Get、普通 Sub、带参 Function 三条反面证人照旧。

Public gTerms As Long
Public gMade As Long
Public gHold As Object

Sub Main()
    Dim o As Object
    Set o = New B147Box

    Dim it As Object
    Set it = New B147Item
    Debug.Print "B147-W0-made=" & gMade

    ' W1/W2: 写一次，交回自己那一枚。W1 该是 0（gHold 拿着），W2 该是 1。
    Set o.Held = it
    Set it = Nothing
    Debug.Print "B147-W1-terms=" & gTerms
    Set gHold = Nothing
    Debug.Print "B147-W2-terms=" & gTerms

    ' W3: 连写两次（同一个槽覆写），最后清空 => 该是 2。
    Set it = New B147Item
    Set o.Held = it
    Set o.Held = it
    Set it = Nothing
    Set gHold = Nothing
    Debug.Print "B147-W3-terms=" & gTerms

    ' W4: 写 Nothing（那一档不取引用）
    Set o.Held = Nothing
    Debug.Print "B147-W4-terms=" & gTerms

    ' R 路（反面证人）: 同名 Let/Get 与普通方法照旧。
    '   判据先落到 Long 再打 —— `"..." & o.Num` 会发成 vb6_ComGetStringProp,
    '   而 Num 是 VT_I4, 那条读臂取的是 bstrVal ⇒ 打印出来是垃圾 (量具的错)。
    Dim n As Long
    o.Num = 41
    n = o.Num
    Debug.Print "B147-R1-num=" & n
    o.Bump 1
    n = o.Num
    Debug.Print "B147-R2-num=" & n
    Debug.Print "B147-R3-tag=" & o.TagOf(it)

    ' R4/R5: 读回写进去的那一枚（Get 交回的那一枚归调用方，见 §11.127）
    Set it = New B147Item
    it.Tag = 55
    Set o.Held = it
    Set gHold = Nothing
    Dim back As Object
    Set back = o.Held
    Debug.Print "B147-R4-terms=" & gTerms
    Set back = Nothing
    Set o.Held = Nothing
    Debug.Print "B147-R5-terms=" & gTerms

    Debug.Print "B147-Z-made=" & gMade
    Debug.Print "B147-DONE"
End Sub
