Option Explicit

' 账 #305 §11.133 (第一百五十二刀): `As Object` 字段出厂放手。
'
' 尺子: item 的 Terminate 计数。
'   A1: 写进字段、放掉局部变量之后, item 还被字段握着 ⇒ 0
'   Z : 放掉持有者 ⇒ 字段那一枚被退 ⇒ item Terminate = 1（改前恒 0）
'   B1/B2: 裸实例 (Me) 进了 As Object 字段, 出厂放手不许越权 Release ⇒ 不崩, 走到 DONE

Public gItemTerms As Long
Public gItemMade As Long

Sub Main()
    Dim it As Object
    Set it = New B152Item
    it.Tag = 7

    Dim h As Object
    Set h = New B152Holder
    Set h.Held = it
    Set it = Nothing
    Debug.Print "B152-A1-items=" & gItemTerms

    Dim got As Object
    Set got = h.Held
    If got Is Nothing Then
        Debug.Print "B152-A2-held=empty"
    Else
        Dim tg As Long
        tg = got.Tag
        Debug.Print "B152-A2-held-tag=" & tg
    End If
    Set got = Nothing

    Set h = Nothing
    Debug.Print "B152-Z-items=" & gItemTerms

    ' ② 裸实例进 As Object 字段: 出厂放手只许清指针
    Dim h2 As Object
    Set h2 = New B152Holder
    h2.KeepMe
    Debug.Print "B152-B1-self-nothing=" & CStr(h2.SelfIsNothing)
    Set h2 = Nothing
    Debug.Print "B152-B2-ok"

    Debug.Print "B152-DONE"
End Sub
