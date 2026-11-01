Attribute VB_Name = "mdl3138"
Option Explicit

' 三档槽位一起钉: 类字段 (C3138.m_iv, 本刀的靶子) / 模块级视图变量 (控制档 —— 按名
' 那张表本来就认得它) / 类实例变量 (FromImpl 的右值)。计数全放这里, Class_Terminate 才读得到。
' 每一位实现者的 Twice 带自己的常数 (H3138=+100 / H3138B=+200) ⇒ 「这格视图槽把方法打给谁」
' 是一个数, 不是一句断言。
Public g_initA As Long
Public g_termA As Long
Public g_initB As Long
Public g_termB As Long

Public g_impl As H3138
Public g_iv As IK3138

Sub Main()
    Dim c As New C3138

    ' 类字段 <- New: 初始引用移交, 旧值 (空) 不退。
    c.FromNew
    Debug.Print "K3138-1=" & g_initA & "/" & g_termA & "/" & c.TwiceViaField(5)

    ' 覆写: 上一枚当场退掉 => termA 恰 1; 槽里换成了 B。
    c.FromNewB
    Debug.Print "K3138-2=" & g_termA & "/" & g_initB & "/" & g_termB & "/" & c.TwiceViaField(5)

    ' 类字段 <- 类实例变量: 适配器 &实例->__iv_I, 实例仍由 g_impl 持有 => termA 不动。
    Set g_impl = New H3138
    c.FromImpl
    Debug.Print "K3138-3=" & g_termA & "/" & g_termB & "/" & g_initA & "/" & c.TwiceViaField(5)

    ' 类字段 <- 模块级视图变量: 视图 -> 视图 = QueryInterface (另一枚 B, g_iv 还拿着它)。
    Set g_iv = New H3138B
    c.FromGlobalView
    Debug.Print "K3138-4=" & g_termA & "/" & g_termB & "/" & g_initB & "/" & c.TwiceViaField(5)

    ' 账 #305 第二百一十刀: 视图接收者的**属性读**那一发要交回同一位主人 —— 两档接收者各钉一发:
    ' 模块级视图变量 (裸名 g_iv, 站S1 那一格) 与类字段 (me->m_iv, 本刀补的第二格)。字段那一格
    ' 改前实测 K3138-7=1200/0: 接收者交进来的是**发射文本** `me->m_iv`, 名册那一问只认裸名
    ' ⇒ 答不出视图 ⇒ 落 IDispatch helper 把薄指针当接口去问 ⇒ 恒 0 (静默: 不崩也不报号,
    ' 只在 stderr 留一声 property "Kind" not found)。
    Dim k1 As Long
    k1 = g_iv.Kind
    Dim k2 As Long
    k2 = c.KindViaField
    Debug.Print "K3138-7=" & k1 & "/" & k2

    ' 置空只放手**这一格**的那一枚引用; g_iv 那一枚还在 => 对象不死。
    c.ClearIt
    Debug.Print "K3138-5=" & g_termB

    ' 控制档 (按名那张表认得的模块级视图变量) 置空: 这里才是最后一枚引用 => B 死。
    Set g_iv = Nothing
    Debug.Print "K3138-6=" & g_termB
    Debug.Print "K3138-DONE"
End Sub
