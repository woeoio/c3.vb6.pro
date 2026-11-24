Option Explicit

' 账 #305 §11.134 (第一百五十三刀): `As Variant` 字段出厂放手。
'
' 尺子: item 的 Terminate 计数。
'   A1: 写进 Variant 槽、放掉局部变量之后, item 还被槽握着 ⇒ 0
'   A2: 槽里读得回 (放手不是"先放后读")
'   Z : 放掉持有者 ⇒ 槽里那一枚被 VariantClear 退掉 ⇒ item Terminate = 1 (改前恒 0)
'   B1: 槽里放标量再读回 —— VariantClear 不许把标量弄坏 (Empty 槽清起来是空操作)
'   B2: 槽里放字符串, 读回长度
'   C1: 同一枚 item 写进两个持有者的槽, 两枚都放 ⇒ 1 次 Terminate (两次 AddRef 两退)

Public gItTerms As Long
Public gItMade As Long

Sub Main()
    Dim it As Object
    Set it = New B153Item
    it.Tag = 5

    Dim h As Object
    Set h = New B153Holder
    Set h.V = it
    Set it = Nothing
    Debug.Print "B153-A1-terms=" & gItTerms

    Dim got As Object
    Set got = h.V
    If got Is Nothing Then
        Debug.Print "B153-A2-held=empty"
    Else
        Dim tg As Long
        tg = got.Tag
        Debug.Print "B153-A2-held-tag=" & tg
    End If
    Set got = Nothing
    Debug.Print "B153-A3-terms=" & gItTerms   ' 改前 1（读出去那枚已经把槽里的拆了）, 改后 0

    Set h = Nothing
    Debug.Print "B153-Z-terms=" & gItTerms

    ' 标量/串那一臂
    Dim h2 As Object
    Set h2 = New B153Holder
    h2.V = 7
    Dim l1 As Long
    l1 = h2.TextLen
    Debug.Print "B153-B1-len=" & l1
    h2.PutText "abcd"
    Dim l2 As Long
    l2 = h2.TextLen
    Debug.Print "B153-B2-len=" & l2
    Set h2 = Nothing

    ' 两个持有者各握一枚
    Dim it2 As Object
    Set it2 = New B153Item
    it2.Tag = 6
    Dim p As Object
    Dim q As Object
    Set p = New B153Holder
    Set q = New B153Holder
    Set p.V = it2
    Set q.V = it2
    Set it2 = Nothing
    Set p = Nothing
    Set q = Nothing
    Debug.Print "B153-C1-terms=" & gItTerms

    Debug.Print "B153-DONE"
End Sub
