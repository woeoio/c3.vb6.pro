' 账 #305 §B218 (甲) 第一百六十四刀的夹具 = 嵌在 If / For / Select 括号里那几档数组载体。
' 从前那一档的声明发在**那个块的括号内**，而出口那一发销毁在括号外 ⇒ 闸① 只能答否 ⇒ 每调用漏一枚；
' 现在声明提到过程序言（创建留在原位，带「拿指针本身当首建守卫」那一发），出口那一发就够得着了。
' 七格各钉一件事：
'   twice = 块进两次只建一枚，第二次读到第一次写进去的值（守卫那一格；没有守卫就每次归零）
'   nd    = 多维那一档嵌在 For 体里：同样的第一次写、第二次读
'   dyn   = 动态那一档嵌在 For 体里：名字本来就在序言（Fix 086），本刀补的是闸① 那一问 ⇒ 出口收尾
'   never = 块一次都不进：序言那枚指针还是 NULL，出口那一发空转不崩（RTL 两头第一句都吃 NULL）
'   multi = 逗号声明里嵌在块里那两档数组，名字各自都提到序言
'   top   = 反面证人：体顶层那一条 Dim 形状一个字不动（声明 + 创建同一行，不带守卫，出口照旧收尾）
'   esc   = 反面证人：整枚交给 UBound 的那一档算逃逸 ⇒ 出口仍旧不碰它（闸② 一条没松）
Option Explicit

Private Sub Twice()
    Dim i As Long
    For i = 1 To 2
        Dim q(0 To 2) As Long
        If q(1) = 0 Then
            q(1) = 7
        Else
            q(1) = q(1) * 2
        End If
        q(2) = i
    Next i
    Debug.Print "K164-twice=" & q(1) & "/" & q(2)
End Sub

Private Sub NdOnce()
    Dim j As Long
    For j = 1 To 2
        Dim m(0 To 1, 0 To 2) As Long
        If m(1, 2) = 0 Then
            m(1, 2) = 9
        Else
            m(1, 2) = m(1, 2) + 1
        End If
    Next j
    Debug.Print "K164-nd=" & m(1, 2)
End Sub

Private Sub DynOnce()
    Dim k As Long
    For k = 1 To 2
        Dim d() As Long
        ReDim d(0 To 3) As Long
        d(3) = d(3) + k
    Next k
    Debug.Print "K164-dyn=" & d(3)
End Sub

Private Sub NeverEntered()
    If 1 = 2 Then
        Dim z(0 To 3) As Long
        z(0) = 5
    End If
    Debug.Print "K164-never=1"
End Sub

Private Sub MultiOnce()
    If 1 = 1 Then
        Dim p(0 To 1) As Long, r(0 To 1) As Long
        p(1) = 4
        r(1) = 6
        Debug.Print "K164-multi=" & p(1) & "/" & r(1)
    End If
End Sub

Private Sub TopOnce()
    Dim t(0 To 2) As Long
    t(0) = 8
    Debug.Print "K164-top=" & t(0)
End Sub

Private Sub EscOnce()
    Dim s As Long
    If 1 = 1 Then
        Dim e2(0 To 2) As Long
        e2(0) = 3
        s = UBound(e2)
    End If
    Debug.Print "K164-esc=" & s
End Sub

Public Sub Main()
    Twice
    NdOnce
    DynOnce
    NeverEntered
    MultiOnce
    TopOnce
    EscOnce
    Debug.Print "K164-DONE=1"
End Sub
