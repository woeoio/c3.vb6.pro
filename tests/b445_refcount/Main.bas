Option Explicit

' 账 #305 §B218 (甲) 刀1 的行为尺: 每格窗口先记下 Made/Terms 的起点, 只打差值,
' 这样各格之间的存活实例不会互相把对方的出厂次数算进自己头上。
' C445: Class_Initialize 里 Main.Made +1, Class_Terminate 里 Main.Terms +1, N 恒 7。
Public Made As Long
Public Terms As Long
Public MadeH As Long
Public TermsH As Long
Private m_auto As New C445
Public m_plain As C445

' W1 = 腿③ 置空: 局部那一枚交出去之前自己先放手。
Sub W1SetNothing()
    Dim a As C445
    Dim t0 As Long
    Dim isNull As Long
    t0 = Terms
    Set a = New C445
    Set a = Nothing
    If a Is Nothing Then isNull = 1
    Debug.Print "W1-terms=" & (Terms - t0) & " null=" & isNull
End Sub

' W2 = R3 别名双退: Set b = a 之后 a 置空, b 必须还活着 (只多一枚引用, 不当场拆)。
Sub W2Alias()
    Dim a As C445
    Dim b As C445
    Dim t0 As Long
    t0 = Terms
    Set a = New C445
    Set b = a
    Set a = Nothing
    Debug.Print "W2-terms=" & (Terms - t0) & " n=" & b.N
End Sub

' W3 = R2 自赋值: 先 claim 后 release, 同一枚不会把自己拆掉。
Sub W3SelfSet()
    Dim a As C445
    Dim t0 As Long
    Dim n0 As Long
    t0 = Terms
    Set a = New C445
    Set a = a
    n0 = a.N
    Set a = Nothing
    Debug.Print "W3-terms=" & (Terms - t0) & " n=" & n0
End Sub

' W4 = 腿② 覆写模块级槽: 每次覆写放手旧的那一枚 (今天这条腿压根没有)。
Sub W4ModuleSlot()
    Dim t0 As Long
    Dim isNull As Long
    t0 = Terms
    Set m_plain = New C445
    Set m_plain = New C445
    Set m_plain = Nothing
    If m_plain Is Nothing Then isNull = 1
    Debug.Print "W4-terms=" & (Terms - t0) & " null=" & isNull
End Sub

' W5 = R6 As New 的守卫: 用到才造, 读两次只造一枚, 而且守卫那条腿不许 claim。
Sub W5AutoNew()
    Dim m0 As Long
    Dim x As Long
    Dim y As Long
    m0 = Made
    x = m_auto.N
    y = m_auto.N
    Debug.Print "W5-made=" & (Made - m0) & " sum=" & (x + y)
End Sub

' W6 = 形⑤ 工厂返回: 被调方交出的是自带对价的那一枚 (claim 在返回槽那一格), 调用方接手后必须还活着。
Public Function MakeOne() As C445
    Dim t As C445
    Set t = New C445
    Set MakeOne = t
End Function

Sub W6Factory()
    Dim c As C445
    Dim t0 As Long
    t0 = Terms
    Set c = MakeOne()
    Debug.Print "W6-terms=" & (Terms - t0) & " n=" & c.N
    Set c = Nothing
End Sub

' W7 = 形参那一档 (今天刻意留在计数系统外面): 被调方里的 Set 形参 = Nothing
' 不许拆掉调用方还在用的那一枚 —— 这条就是 procClassFormals_ 那面夹的行为针。
Public Sub DropByVal(ByVal oF As C445)
    Set oF = Nothing
End Sub

Sub W7ParamBorrow()
    Dim a As C445
    Dim t0 As Long
    t0 = Terms
    Set a = New C445
    DropByVal a
    Debug.Print "W7-terms=" & (Terms - t0) & " n=" & a.N
    Set a = Nothing
End Sub

' W8 = 腿④ 过程出口: 每一轮现造的那一枚在自己那一格出口放手 (改前一条都没有)。
'      这一形就是 (甲) 的靶 = 泄漏尺 w-newonly 那一路: 循环 200 次, 出厂恰好 200 次。
Private Sub W8One()
    Dim a As C445
    Set a = New C445
End Sub

Sub W8ScopeExit()
    Dim t0 As Long
    Dim i As Long
    t0 = Terms
    For i = 1 To 200
        W8One
    Next
    Debug.Print "W8-terms=" & (Terms - t0)
End Sub

' W9 = 腿④ 的夹: 返回槽那一枚交出去, 出口不许放手 (放了 = 调用方刚接手就读已释放块)。
'      第一百三十一刀之前调用方**还多 claim 一枚** (实测 BASE 发码 `if (c) ... _AddRef(c);` =
'      每调用净 +1, 那一枚永远没人退), 现在这一格改成接手 —— 本窗口打印在放手之前, 所以两态都
'      答 terms=0 (它钉的是「接手后还活着」); 漏与不漏的行为面在 W15 / W16 那两格上。
Sub W9ReturnSlot()
    Dim c As C445
    Dim t0 As Long
    t0 = Terms
    Set c = MakeOne()
    Debug.Print "W9-terms=" & (Terms - t0) & " n=" & c.N & " twice=" & c.N
    Set c = Nothing
End Sub

' W10 = 腿②' 出参槽 (ByRef 那座桥的写回站): 被调方现造一枚写进调用方那一格, 而自己那一格
'       的局部会在被调方出口放手 —— 所以那一格必须自己取走一枚引用, 否则调用方还看着它时
'       它已经拆了 (门 #763 的 b187 第五臂 0xC0000409 = 就是这个形状)。
'       旧值那一格 (Set o = New 造的第一枚) 今天没人退 = 出参写回的泄漏, 与下标槽同在册。
Public Sub GiveFreshRef(p As Object)
    Dim g As C445
    Set g = New C445
    Set p = g
End Sub

Sub W10OutSlot()
    Dim held10 As C445
    Dim t0 As Long
    t0 = Terms
    Set held10 = New C445
    GiveFreshRef held10
    Debug.Print "W10-live-terms=" & (Terms - t0) & " n=" & held10.N
    Set held10 = Nothing
    Debug.Print "W10-gone-terms=" & (Terms - t0)
End Sub

' W11 = 腿④ 的第二道夹: 把一格工程类局部的**地址**交给外部 Declare (VB6 运行时的
'       __vbaObjSetAddref 那一族) ⇒ 这一格由外面写。C3 的那枚桩对工程内实例刻意不 AddRef
'       (结构体首字段不是 vtable), 所以本过程手里没有这一枚引用的对价, 出口就不许退它。
'       门 #763 的 FlexGridX86 = 就是这个形状把还活着的控件当场拆掉。
Public Declare PtrSafe Function SetAddRefAny Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" _
    (ByRef Destination As Any, ByVal lpObject As LongPtr) As Long

Sub W11ForeignSlot()
    Dim borrowed As C445
    Dim own11 As C445
    Dim t0 As Long
    Dim n As Long
    Dim termsMid As Long
    Dim termsEnd As Long
    t0 = Terms
    Set own11 = New C445
    SetAddRefAny borrowed, ObjPtr(own11)
    n = borrowed.N
    termsMid = Terms - t0
    Set own11 = Nothing
    termsEnd = Terms - t0
    Debug.Print "W11-n=" & n & " terms=" & termsMid & " owner-gone=" & termsEnd
End Sub

' W12 = 腿⑤ 出厂那一档: 字段里那一枚随持有者一起出厂 (改前没人退 ⇒ 内层永远走不到 Terms)。
Sub W12FieldTeardown()
    Dim h12 As C445H
    Dim m0 As Long
    Dim t0 As Long
    Dim hm0 As Long
    Dim ht0 As Long
    m0 = Made
    t0 = Terms
    hm0 = MadeH
    ht0 = TermsH
    Set h12 = New C445H
    h12.Hold12
    Set h12 = Nothing
    Debug.Print "W12-inner-made=" & (Made - m0) & " inner-terms=" & (Terms - t0)
    Debug.Print "W12-holder-made=" & (MadeH - hm0) & " holder-terms=" & (TermsH - ht0)
End Sub

' W13 = 覆写 + 出厂两档同台: 造两枚, 覆写退掉第一枚, 出厂再退掉第二枚 = 恰好两枚出厂。
Sub W13FieldOverwriteThenTeardown()
    Dim h13 As C445H
    Dim t1 As Long
    t1 = Terms
    Set h13 = New C445H
    h13.HoldTwice12
    Set h13 = Nothing
    Debug.Print "W13-terms=" & (Terms - t1)
End Sub

' W14 = 反面证人 (不许提前拆): 别名进字段时 claim 一枚, 持有者先出厂, 调用方那一枚还得活着。
Sub W14FieldAliasKeepsAlive()
    Dim o14 As C445
    Dim h14 As C445H
    Dim t2 As Long
    t2 = Terms
    Set o14 = New C445
    Set h14 = New C445H
    h14.AliasInto12 o14
    Set h14 = Nothing
    Debug.Print "W14-alive-terms=" & (Terms - t2) & " n=" & o14.N
    Set o14 = Nothing
    Debug.Print "W14-gone-terms=" & (Terms - t2)
End Sub

'
' W15 / W16 = 账 #305 §B218 (甲) 刀3 的两形: 工厂交回来的那一枚引用**已经带对价**, 接手这一格
'       不许再 claim (改前每调用净 +1, 那一枚永远没人退 => 下面两格 terms 恒 0)。
'       W15 = 模块级 Function 工厂循环 50 次: 每轮接手一枚、下一轮退掉上一轮那枚, 出口退最后一枚
'       => made=50 terms=50; W16 = 同一件事换 Property Get 那一形 (问句必须两形同答)。
Sub W15FactoryLoop()
    Dim c As C445
    Dim m0 As Long
    Dim t0 As Long
    Dim i As Long
    m0 = Made
    t0 = Terms
    For i = 1 To 50
        Set c = MakeOne()
    Next
    Set c = Nothing
    Debug.Print "W15-made=" & (Made - m0) & " terms=" & (Terms - t0)
End Sub

Sub W16PropFactory()
    Dim hh As C445H
    Dim m0 As Long
    Dim t0 As Long
    Set hh = New C445H
    m0 = Made
    t0 = Terms
    hh.PropFactory16
    Debug.Print "W16-made=" & (Made - m0) & " terms=" & (Terms - t0)
End Sub

'
' W17 / W18 / W19 = 账 #305 §B218 (甲) 刀4 的两形 + 反面那一档:
'       W17 = 工厂挂在**实例**上 (`hh.Make17()`) 循环 20 次现造现放手 => made=20 terms=20;
'       W18 = 属性读**不带括号** (`hh.Next18`) 同一契约 => made=1 terms=1;
'       W19 = 反面: `hh.mField19` 是**数据字段**, 那一枚仍由持有者握着 => 读出来必须 claim,
'             放手只退自己那一枚, 持有者出厂时才走第二次 (alive=0 / gone=1)。
Sub W17MemberFactoryLoop()
    Dim hh17 As C445H
    Dim c17 As C445
    Dim m0 As Long
    Dim t0 As Long
    Dim i As Long
    Set hh17 = New C445H
    m0 = Made
    t0 = Terms
    For i = 1 To 20
        Set c17 = hh17.Make17()
    Next
    Set c17 = Nothing
    Debug.Print "W17-made=" & (Made - m0) & " terms=" & (Terms - t0)
End Sub

Sub W18PropReadNoParens()
    Dim hh18 As C445H
    Dim c18 As C445
    Dim m0 As Long
    Dim t0 As Long
    Set hh18 = New C445H
    m0 = Made
    t0 = Terms
    Set c18 = hh18.Next18
    Set c18 = Nothing
    Debug.Print "W18-made=" & (Made - m0) & " terms=" & (Terms - t0)
End Sub

Sub W19FieldReadKeepsClaim()
    Dim hh19 As C445H
    Dim f19 As C445
    Dim t3 As Long
    Set hh19 = New C445H
    hh19.Hold19
    t3 = Terms
    Set f19 = hh19.mField19
    Set f19 = Nothing
    Debug.Print "W19-field-alive-terms=" & (Terms - t3) & " n=" & hh19.mField19.N
    Set hh19 = Nothing
    Debug.Print "W19-field-gone-terms=" & (Terms - t3)
End Sub

Sub Main()
    W1SetNothing
    W2Alias
    W3SelfSet
    W4ModuleSlot
    W5AutoNew
    W6Factory
    W7ParamBorrow
    W8ScopeExit
    W9ReturnSlot
    W10OutSlot
    W11ForeignSlot
    W12FieldTeardown
    W13FieldOverwriteThenTeardown
    W14FieldAliasKeepsAlive
    W15FactoryLoop
    W16PropFactory
    W17MemberFactoryLoop
    W18PropReadNoParens
    W19FieldReadKeepsClaim
    Debug.Print "B445-DONE"
End Sub
