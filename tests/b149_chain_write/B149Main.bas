Option Explicit

' 账 #305 §11.130 (第一百四十九刀): `Set <晚绑定 o>.<Child>.<Prop> = <对象>`。
'
' 改前两槽双双串位, 发码是
'   vb6_B148Box_prop_set_Held(o, vb6_ComGetObjectProp(o, L"Child"))
' —— 接收者写成裸 `o`（该是 o.Child 交回的那一枚），值写成 o.Child 的交回（该是 it）。
' 两个病: (1) 早绑定直呼那一支只把**裸标识符**接收者判成晚绑定，成员链接收者漏网;
'        (2) `emitExpr(o.Child)` 只交出基础对象、把成员名挂在 marker 上，该 marker 漏进
'            值槽 ⇒ 值被求成 o.Child 的 getter。
'
' A 路: 根对象 o 的 Held 必须还是空（没写到根上）
' B 路: 子对象 o.Child 的 Held 里必须是 it 那一枚
' C 路: 同一枚的证人 —— 改 h1.Tag 后 it.Tag 跟着变
' Z 路: 收尾。box 停在 1 而非 2 是**在册**的: 本夹具里 `Set h1 = o.Child.Held` 是一发
'      **嵌套读链**，中间那枚 `o.Child` 交回的引用没人放（§11.132 那把尺）。
'      链写那一站（`Set p.Child.Held = it`）的接收者已在 §11.131 对平 ——
'      只留链写的形状见 tests/b150_chain_recv_ref（box-terms 1→2）。

Public gBoxTerms As Long
Public gItemTerms As Long
Public gItemMade As Long

Sub Main()
    Dim o As Object
    Set o = New B149Box
    o.Tag = 5

    Dim it As Object
    Set it = New B149Item
    it.Tag = 7

    Set o.Child.Held = it          ' ← 靶

    Dim h0 As Object
    Set h0 = o.Held
    Debug.Print "B149-A1-root-held=" & ObjName(h0)

    Dim h1 As Object
    Set h1 = o.Child.Held
    Debug.Print "B149-A2-child-held=" & ObjName(h1)

    If h1 Is Nothing Then
        Debug.Print "B149-A3-tag=-1"
    Else
        h1.Tag = 9
        Dim t As Long
        t = it.Tag
        Debug.Print "B149-A3-tag=" & t
    End If

    Set h1 = Nothing
    Set h0 = Nothing
    Set it = Nothing
    Debug.Print "B149-C1-item-terms=" & gItemTerms
    Set o = Nothing
    Debug.Print "B149-Z-box-terms=" & gBoxTerms
    Debug.Print "B149-Z-item-made=" & gItemMade
    Debug.Print "B149-Z-item-terms=" & gItemTerms
    Debug.Print "B149-DONE"
End Sub

Public Function ObjName(ByVal v As Object) As String
    If v Is Nothing Then
        ObjName = "empty"
    Else
        ObjName = TypeName(v)
    End If
End Function
