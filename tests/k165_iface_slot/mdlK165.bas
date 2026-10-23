Attribute VB_Name = "mdlK165"
Option Explicit

' <vbeclipse> 第一百六十五刀 回归夹子: 模块级**接口视图槽**的覆写与置空。
' 病灶 = 腿②/腿③ 那一站先问类名 (inferClassTypeOfExpr 把接口名也答出来),
' 于是铸出一枚声明侧从没造过的 `vb6_cls_K165Iface*` = cl C2440 (十二形实测,
' 门 #832 之后由 #407 的重做探针抓到)。修法 = 先问视图名册那一位 (slotIsLegacyIfaceView),
' 视图那一档落回裸拷贝 / vb6_ReleaseObject, 与局部那一档今天发的一字相同。
Dim mv As K165Iface
Dim mi As K165Impl

Sub Main()
    Set mi = New K165Impl
    Set mv = New K165Impl
    Debug.Print "K165-new=" & mv.Bar(3)
    Set mv = mi
    Debug.Print "K165-over=" & mv.Bar(4)
    Dim lv As K165Iface
    Set lv = New K165Impl
    Debug.Print "K165-local=" & lv.Bar(6)
    Set mv = Nothing
    Set mv = mi
    Debug.Print "K165-afternothing=" & mv.Bar(7)
    Debug.Print "K165-mi=" & mi.Twice(8)
    Debug.Print "K165-DONE=1"
End Sub
