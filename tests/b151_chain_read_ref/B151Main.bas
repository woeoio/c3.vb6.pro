Option Explicit

' 账 #305 §11.132 (第一百五十一刀): 链读**接收者**那一枚临时引用。
'
' `vb6_ComGetObjectProp(o, L"Child")` 交回的那一枚归调用方 (§11.126)，被当接收者再递一层
' 时没人放 ⇒ 每多一层链漏一枚，中间那枚对象永不 Terminate。改前发码（一条语句、嵌套）：
'     vb6_ComGetIntProp(vb6_ComGetObjectProp(p, L"Child"), L"Tag")
'     vb6_ComGetObjectProp(vb6_ComGetObjectProp(p, L"Child"), L"Held")
' 改后：
'     void* _rrv0 = vb6_ComGetObjectProp(p, L"Child");
'     ... <用 _rrv0> ...
'     vb6_ReleaseObject((void**)&_rrv0);   /* 语句末 */
'
' 尺子是 box 的 Terminate 计数（根 + 每层子对象各一枚）：
'   A: 嵌套对象读 `Set h = p.Child.Held` —— 值对不对
'   B: 嵌套标量读 `v = p.Child.Tag`     —— 值对不对（子对象预置 99）
'   C: 两层链 `w = p.Child.Child.Tag`   —— 两层中间量都要放
'   Z: box-terms = 3（改前 1：两层中间量各夹住一枚，只剩根 Terminate）

Public gBoxTerms As Long
Public gBoxMade As Long
Public gItemTerms As Long
Public gItemMade As Long

Sub Main()
    Dim p As Object
    Set p = New B151Box
    p.Tag = 5

    Dim it As Object
    Set it = New B151Item
    it.Tag = 7

    Set p.Held = it                     ' 单层写(基线, §11.130 已对平)
    Set p.Child.Held = it               ' 链写(基线, §11.131 已对平)

    Dim h As Object
    Set h = p.Child.Held                ' ← 嵌套对象读
    If h Is Nothing Then
        Debug.Print "B151-A1-held=empty"
    Else
        Dim t As Long
        t = h.Tag
        Debug.Print "B151-A2-held-tag=" & t
    End If

    Dim v As Long
    v = p.Child.Tag                     ' ← 嵌套标量读
    Debug.Print "B151-B1-child-tag=" & v

    Dim w As Long
    w = p.Child.Child.Tag               ' ← 两层链
    Debug.Print "B151-C1-two-level-tag=" & w

    Set h = Nothing
    Set it = Nothing
    Debug.Print "B151-Y-box-made=" & gBoxMade
    Set p = Nothing
    Debug.Print "B151-Z-box-terms=" & gBoxTerms
    Debug.Print "B151-Z-item-made=" & gItemMade
    Debug.Print "B151-Z-item-terms=" & gItemTerms
    Debug.Print "B151-DONE"
End Sub
