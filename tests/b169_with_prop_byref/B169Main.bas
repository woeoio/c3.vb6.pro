Option Explicit

' 账 #305 §11.151 (第一百六十九刀): 「属性读当 ByRef 实参」的两处收口。
'   W1 = With 块里的属性读 (改前一个都不成立: `.Tag` 在 AST 上是 WithMemberExpr，
'        既没有 object 可推型，我们那一问又只认 MemberAccessExpr)
'   W2 = 型不一致那一格的证人 (Long 属性接 Integer 形参) —— §11.151 时仍按值 ⇒ 40；
'        §11.152 (第一百七十刀) 开了换算 ⇒ 41。
'   W3 = 早绑定直喂 (与前几刀同形，证人)
'   W4 = With 块 + String 属性 (临时量的 BSTR 照旧归调用方退)
'   W5 = With 块 + 带索引属性
Public Sub Main()
    Dim b As B169Box

    Set b = New B169Box

    b.Tag = 40
    With b
        .Inc .Tag
    End With
    Debug.Print "B169-W1=" & b.Tag

    b.Tag = 40
    b.IncI b.Tag
    Debug.Print "B169-W2=" & b.Tag

    b.Tag = 40
    b.Inc b.Tag
    Debug.Print "B169-W3=" & b.Tag

    b.S = "ab"
    With b
        .App .S
    End With
    Debug.Print "B169-W4=" & b.S

    b.Slot(2) = 10
    With b
        .Inc .Slot(2)
    End With
    Debug.Print "B169-W5=" & b.Slot(2)

    Debug.Print "B169-DONE"
End Sub
