Attribute VB_Name = "mdl2947"
Option Explicit

' 账 #305 站F 写侧 (第一百八十五~一百八十七刀): Set / 多入口参数 / 对象属性读 在视图上都走槽表。
Function MakeI() As IK2947
    Dim t As IK2947
    Set t = New H2947
    Set t.Owner = New O2947
    Set MakeI = t
End Function

Sub Main()
    Dim x As IK2947
    Dim s As String
    Dim o As Object
    Set x = New H2947
    Set x.Owner = New O2947
    Debug.Print "K2947-a=" & x.Has()
    Set o = x.Owner
    Debug.Print "K2947-o=" & (o Is Nothing)
    Set x.Owner = Nothing
    Debug.Print "K2947-b=" & x.Has()
    Set o = x.Owner
    Debug.Print "K2947-p=" & (o Is Nothing)
    x.Item(3) = "a"
    s = x.Item(3)
    Debug.Print "K2947-i=" & s
    ' 第一百八十七刀补: With 体内的对象档读 —— 与块外同一位读者 (今天之前没有证人)
    Set x.Owner = New O2947
    With x
        Set o = .Owner
        Debug.Print "K2947-w=" & (o Is Nothing)
        Set .Owner = Nothing
        Set o = .Owner
        Debug.Print "K2947-w0=" & (o Is Nothing)
    End With
    ' 第一百八十八刀: 视图作返回类型 (Function 与 Property Get 两条路)
    Dim y As IK2947
    Dim h As H2947
    Set y = MakeI()
    Debug.Print "K2947-f=" & y.Has()
    Set h = New H2947
    Set x = h
    Set x.Owner = New O2947
    Set y = h.OwnView
    Debug.Print "K2947-g=" & y.Has()
    ' 第一百八十九刀: 接收者是返回视图的**表达式** (Property Get 调用) —— 按值适配器
    Debug.Print "K2947-h=" & h.OwnView.Has()
    ' 第一百九十刀: 视图表达式上的**裸**零参属性读 (无括号 = 成员访问路)
    Debug.Print "K2947-j=" & h.OwnView.Has
    ' 第一百九十一刀: With Me 体内读属性返回的视图 (Me 是 MeExpr, 按本类实例处理)
    Debug.Print "K2947-k=" & h.ViaMe
    ' 第一百九十二刀: With <返回视图的调用> —— 临时量按视图值拷贝, 体内读走同一族
    With MakeI()
        Set o = .Owner
        Debug.Print "K2947-l=" & (o Is Nothing)
        ' 第一百九十三刀: 同块内的**零参属性读** (视图临时量 → 按值适配器, 不是 IDispatch)
        Debug.Print "K2947-m=" & .Has
    End With
    ' 第一百九十三刀: 接收者是 MemberAccess 视图属性 (h.OwnView) 时的体内零参读
    With h.OwnView
        Debug.Print "K2947-n=" & .Has
    End With
    Debug.Print "K2947-DONE"
End Sub
