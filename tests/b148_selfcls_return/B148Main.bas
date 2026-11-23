Option Explicit

' 账 #305 §11.129 (第一百四十八刀): 服务端 getter 适配器里**对象返回**那两档。
'
' 改前 `Public Property Get Self() As B148Box`（返回类型 = 所在类自己）和
' `Public Property Get Held() As Object` 的 C 返回型都是 `void*`，于是 dll_entry 里那条
' 「工程类实例返回值先包一层 IDispatch」的口径不命中，两条都落到最后一个
' else: `vt = VT_I4; lVal = (int32_t)(intptr_t)_r` —— 对象被当整数交出去，x64 上指针当场截断。
' 实测 trace: GetProp 'Self' → vt=3 val=150397680（而 As B148Item 那条是 vt=9）。
'
' A 路: 本类返回必须交回对象（改前 = Nothing）。
' B 路: 异类返回是对照组，改前改后都该是对象。
' C 路: As Object 返回必须交回对象，且交回的确实是写进去那一枚（刀147 写口 + 本刀读口）。
' Z 路: 引用账 —— 交回的那一枚归调用方（§11.126），清 s 之后 o 必须还活着。

Public gBoxTerms As Long
Public gItemTerms As Long
Public gItemMade As Long
Public gHold As Object

Sub Main()
    Dim o As Object
    Set o = New B148Box
    o.Tag = 5

    ' A: 本类返回
    Dim s As Object
    Set s = o.Self
    If s Is Nothing Then
        Debug.Print "B148-A1-self=Nothing"
    Else
        Dim t As Long
        t = s.Tag
        Debug.Print "B148-A2-self-tag=" & t
    End If
    Set s = Nothing
    o.Tag = 6
    Dim t2 As Long
    t2 = o.Tag
    Debug.Print "B148-A3-tag=" & t2

    ' B: 异类返回（对照组）
    Dim f As Object
    Set f = o.Fresh
    If f Is Nothing Then
        Debug.Print "B148-B1-fresh=Nothing"
    Else
        f.Tag = 11
        Dim u As Long
        u = f.Tag
        Debug.Print "B148-B2-fresh-tag=" & u
    End If
    Set f = Nothing
    Debug.Print "B148-B3-item-terms=" & gItemTerms

    ' C: As Object 返回
    Dim it As Object
    Set it = New B148Item
    Set o.Held = it
    Dim h As Object
    Set h = o.Held
    If h Is Nothing Then
        Debug.Print "B148-C1-held=Nothing"
    Else
        h.Tag = 77
        Dim w As Long
        w = it.Tag
        Debug.Print "B148-C2-held-tag=" & w
    End If
    Set h = Nothing
    Set o.Held = Nothing
    Set it = Nothing
    Debug.Print "B148-C3-item-terms=" & gItemTerms

    ' Z: 清空之后 —— box 这一格停在 0, 是**在册**的那一枚残留:
    '    `Set Self = Me` 在返回槽上 claim 了一枚 (服务端 getter 的泛对象返回槽, §11.125 步骤 4
    '    同族), 适配器再从同一枚包装器上取一枚交给 VARIANT; 客户端拿走那一枚、放手之后,
    '    实例自己的 __refcount 仍多 1 ⇒ Class_Terminate 不触发。本刀不接它 (那是另一把尺),
    '    这一行的 0 就是它的读数。item 那一枚 (异类返回) 是 1, 说明收放照样对得上。
    Set o = Nothing
    Debug.Print "B148-Z-box-terms=" & gBoxTerms
    Debug.Print "B148-Z-item-made=" & gItemMade
    Debug.Print "B148-Z-item-terms=" & gItemTerms
    Debug.Print "B148-DONE"
End Sub
