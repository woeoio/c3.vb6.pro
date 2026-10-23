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

' 账 #305 §B218 (甲) 第一百四十刀 = 下标槽那一档: 目标是 VB6_SA_AT(...) 那一形 —— 带括号**但是左值**,
'   所以「带括号就当函数调用」那条老判据对它不成立。三站各一枚靶 (W20/W21/W22)。
Private mArr445(1 To 3) As C445
' 账 #305 §B218 (甲) 第一百四十六刀 = 载体销毁那一族: Erase / ReDim 不带 Preserve 换掉整块载体时,
'   元素那一枚引用由这一格握着 (腿②③ 已认同一位主人)，销毁前必须逐枚交回。靶两枚: 固定 + 动态。
Private mDyn445() As C445
' 第一百四十六刀的两枚反面证人: W33 = 载体里的元素**不是**工程类 (泛对象包装器 / 裸值)，
' W34 = 多维载体那一档本刀刻意没接 (RTL 还没有 ND 版出口)。
Private mObjArr445() As Object
Private mLngArr445() As Long
Private mND445() As C445
' 第一百五十四刀的反面证人: 同样是多维 Preserve, 但元素是裸值 (As Long) ⇒ 那一发必须仍走不带回调的转调
Private mNDLng445() As Long
Public Side445 As Long

Public Function Pick445() As Long
    Side445 = Side445 + 1
    Pick445 = 3
End Function

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

' W20 = 覆写那一格: 连设两次 ⇒ 前一枚当场出厂 (修前 = 裸赋值, 旧值没人退 ⇒ 0)。
Sub W20ArrElemOverwrite()
    Dim t0 As Long
    t0 = Terms
    Set mArr445(1) = New C445
    Set mArr445(1) = New C445
    Debug.Print "W20-over-terms=" & (Terms - t0)
    Set mArr445(1) = Nothing
End Sub

' W21 = 置空那一格: 放手 + 真清槽 (修前 = 把复合字面量 & 的槽交给 vb6_ReleaseObject, 而它认不出
'   工程类裸结构体 ⇒ 既不放手也不清槽, 整条语句是 no-op)。
Sub W21ArrElemNothing()
    Dim t0 As Long
    t0 = Terms
    Set mArr445(2) = New C445
    Set mArr445(2) = Nothing
    Debug.Print "W21-terms=" & (Terms - t0) & " cleared=" & (mArr445(2) Is Nothing)
End Sub

' W22 = 「下标算术算几次」: 每条语句的复合下标只许算一次 (两发 = 2)，同时那一格照样退旧。
Sub W22IndexSideEffectOnce()
    Dim t0 As Long
    Side445 = 0
    t0 = Terms
    Set mArr445(Pick445()) = New C445
    Set mArr445(Pick445()) = New C445
    Debug.Print "W22-side=" & Side445 & " terms=" & (Terms - t0)
    Set mArr445(3) = Nothing
End Sub

' W23 = 数组字段那一档的出厂 (腿⑤ 数组形) + 字段里的覆写: 改前两处都 0 枚 ⇒ 内层永不走 Class_Terminate。
Sub W23ArrFieldTeardown()
    Dim hh As C445H
    Dim t0 As Long
    Dim m0 As Long
    t0 = Terms
    m0 = Made
    Set hh = New C445H
    hh.FillKids23
    hh.FillKids23
    Debug.Print "W23-over-terms=" & (Terms - t0)
    Set hh = Nothing
    Debug.Print "W23-out-made=" & (Made - m0) & " terms=" & (Terms - t0)
End Sub

' ==== 账 #305 §B218 (甲) 第一百四十六刀 = 载体销毁那一族 (Erase / ReDim 不带 Preserve) ====
' 修前: 这两发都是裸销毁 —— 动态那一支 `vb6_SafeArrayDestroy1D` 对 vb6_sa_ptr 元素压根没有那一支,
'       固定那一支逐元素走 `vb6_ReleaseObject`, 而它的 Fix 191 判据认不出工程类裸结构体 ⇒ 一枚不退。
' 修后: 销毁载体之前先问一处「这一枚数组握着的元素是不是工程类」，是 ⇒ 逐枚交回 (RTL 那头先清槽
'       再放手 ⇒ Terminate 回头写这一格不会退第二遍)。

' W29 = 固定类数组的 Erase: 两枚当场出厂 + 槽真清 (修前 = terms 0、槽却被 memset 清成 Nothing)。
Sub W29EraseFixedArr()
    Dim t0 As Long
    t0 = Terms
    Set mArr445(1) = New C445
    Set mArr445(2) = New C445
    Erase mArr445
    Debug.Print "W29-terms=" & (Terms - t0) & " cleared=" & (mArr445(1) Is Nothing)
End Sub

' W30 = 动态类数组的 Erase (同一枚 RTL 出口的 isDynamic 那一支): 两枚当场出厂。
Sub W30EraseDynArr()
    Dim t0 As Long
    t0 = Terms
    ReDim mDyn445(1 To 2)
    Set mDyn445(1) = New C445
    Set mDyn445(2) = New C445
    Erase mDyn445
    Debug.Print "W30-terms=" & (Terms - t0)
End Sub

' W31 = ReDim 不带 Preserve 换掉旧载体: 旧的两枚当场出厂 (修前 = 0 枚)。
Sub W31RedimNoPreserve()
    Dim t0 As Long
    t0 = Terms
    ReDim mDyn445(1 To 2)
    Set mDyn445(1) = New C445
    Set mDyn445(2) = New C445
    ReDim mDyn445(1 To 5)
    Debug.Print "W31-terms=" & (Terms - t0)
End Sub

' W32 = 反面证人: Preserve 那一支元素是**被搬走**的, 一格都不许多退 (留下那两枚还得照样能用)。
Sub W32PreserveKeepsElems()
    Dim t0 As Long
    t0 = Terms
    ReDim Preserve mDyn445(1 To 6)
    Set mDyn445(1) = New C445
    Set mDyn445(2) = New C445
    ReDim Preserve mDyn445(1 To 6)
    Debug.Print "W32-terms=" & (Terms - t0) & " kept=" & (Not (mDyn445(1) Is Nothing))
    Erase mDyn445
End Sub

' W33 = 反面证人: 载体里躺的**不是**工程类实例 (As Object = 包装器 / As Long = 裸值) ⇒ 三道闸拦下，
'   这一站一声不出。今天这两发照样一枚不退 = 泛对象档那一格还欠口径 (在册)，钉成证词不是愿望。
Sub W33NonClassCarriers()
    Dim t0 As Long
    t0 = Terms
    ReDim mObjArr445(1 To 2)
    Set mObjArr445(1) = New C445
    ReDim mLngArr445(1 To 2)
    mLngArr445(1) = 7
    ReDim Preserve mLngArr445(1 To 3)
    Erase mObjArr445
    Erase mLngArr445
    Debug.Print "W33-terms=" & (Terms - t0)
End Sub

' W35 = Preserve **收缩**那一格 (第一百四十八刀): VB6 把尾那两枚当场放手 (尺寸变小 = 截断,
'   被截掉的元素移除), 修前 = 整批裸丢一枚不退。保留那一枚必须照常能用 = 同一条语句的另一半。
Sub W35PreserveShrink()
    Dim t0 As Long
    t0 = Terms
    ReDim mDyn445(1 To 3)
    Set mDyn445(1) = New C445
    Set mDyn445(2) = New C445
    Set mDyn445(3) = New C445
    ReDim Preserve mDyn445(1 To 1)
    Debug.Print "W35-terms=" & (Terms - t0) & " kept=" & (Not (mDyn445(1) Is Nothing))
    Set mDyn445(1) = Nothing
End Sub

' W34 = 多维载体那一档 (第一百四十九刀量、第一百五十刀补): 销毁 ND 载体之前，两格元素各交回一枚 ⇒ 2。
'   发码走的是 vb6_SA_ReleaseElementsND 那枚兄弟出口，形状那一问由主人自己问 (站点不判维度)。
Sub W34NdCarrierUntouched()
    Dim t0 As Long
    t0 = Terms
    ReDim mND445(1 To 2, 1 To 2)
    Set mND445(1, 1) = New C445
    Set mND445(2, 2) = New C445
    Erase mND445
    Debug.Print "W34-terms=" & (Terms - t0)
End Sub

' ==== 账 #305 §B218 (甲) 第一百四十四刀 = 形A 那座桥 (ByRef 工程类形参的写回站) ====
' 修前: 调用侧交出去的是一枚**副本** (vb6_argtmp_*)，调用行之后没人写回 ⇒ 被调方那一发 `Set <形参> = ...`
'       整条丢掉，而且那一枚 New 从此没人退；链形 (实参本身是 ByRef 工程类形参) 更糟 —— 交出去的是一枚
'       vb6_VARIANT*，里面还替调用者那枚实例多握了一个没人放手的包装器。
' 修后: 持有者那一格把**自己的地址**交出去 ⇒ 写回当场落地 + 当场退旧；不持有那一格 (ByVal / 移交值)
'       仍走副本，但副本先付一枚 (copy-in)、调用行之后交回 ⇒ 计数守恒，调用者那一格一个字节不动。
Public Sub OutOver24(p As C445)
    Set p = New C445
End Sub

' W24 = 持有者直传: 写回落地 (same 翻false) + 旧值当场出厂 (mid-terms=1) + 两台计数守恒。
Sub W24OutParamWriteBack()
    Dim a As C445
    Dim d As C445
    Dim t0 As Long
    Dim m0 As Long
    Set d = New C445
    Set a = d
    t0 = Terms
    m0 = Made
    OutOver24 a
    Debug.Print "W24-same=" & (a Is d) & " mid-terms=" & (Terms - t0)
    Set a = Nothing
    Set d = Nothing
    Debug.Print "W24-made=" & (Made - m0) & " end-terms=" & (Terms - t0)
End Sub

' W25 = 链形 (实参是本过程的 ByRef 工程类形参): 同一格一路传到底 ⇒ 写回照样落地，
'       而且不再替调用者多铸一枚包装器 (修前 = 调用者那一枚再也拆不掉, end-terms=0)。
Public Sub ChainOut25(q As C445)
    OutOver24 q
End Sub

Sub W25ChainCarriesBack()
    Dim b As C445
    Dim e As C445
    Dim t0 As Long
    Set e = New C445
    Set b = e
    t0 = Terms
    ChainOut25 b
    Debug.Print "W25-same=" & (b Is e) & " mid-terms=" & (Terms - t0)
    Set b = Nothing
    Set e = Nothing
    Debug.Print "W25-end-terms=" & (Terms - t0)
End Sub

' W26 = 出参那一格置空: 落地 + 退旧 (修前 = 裸写一枚 NULL 进副本, 调用者手里一切如旧)。
Public Sub OutNil26(p As C445)
    Set p = Nothing
End Sub

Sub W26OutParamNil()
    Dim c As C445
    Dim t0 As Long
    Set c = New C445
    t0 = Terms
    OutNil26 c
    Debug.Print "W26-null=" & (c Is Nothing) & " terms=" & (Terms - t0)
    Set c = Nothing
End Sub

' W27 = 反面证人: **ByVal** 那一格不持有它交出去的那一枚 ⇒ 刻意不落地，但副本先付一枚、调用行之后交回。
Public Sub ByValProbe27(ByVal x As C445)
    OutOver24 x
End Sub

Sub W27ByValDoesNotLand()
    Dim f As C445
    Dim g As C445
    Dim t0 As Long
    Dim m0 As Long
    Set g = New C445
    Set f = g
    t0 = Terms
    m0 = Made
    ByValProbe27 f
    Debug.Print "W27-same=" & (f Is g) & " mid-terms=" & (Terms - t0)
    Set f = Nothing
    Set g = Nothing
    Debug.Print "W27-made=" & (Made - m0) & " end-terms=" & (Terms - t0)
End Sub

' W28 = 移交值当实参 (MakeOne() 那一枚自带对价): 不补 copy-in，只交回；被调方退旧退的正是移交来的那一枚。
Sub W28HandedOverArg()
    Dim t0 As Long
    Dim m0 As Long
    t0 = Terms
    m0 = Made
    OutOver24 MakeOne()
    Debug.Print "W28-mid-terms=" & (Terms - t0)
    Debug.Print "W28-made=" & (Made - m0) & " end-terms=" & (Terms - t0)
End Sub

' ==== 账 #305 §B218 (甲) 第一百五十二刀 = 腿④ 与腿⑤ 合体那一格: 过程级**类数组**出作用域 ====
' 修前: 载体与元素两样都没人放 (本机探针 A6 局部动态 / A7 局部固定 各退 0 枚, `.build/k1650_p64.txt`);
' 修后: 出口那一发把里面每一枚交回同一位主人 (臂仍由主人自己问维度)。非类那一档是**反面证人** = 一枚不退。
' 刻意不在这一刀吃「销毁载体」: 出口那一发的形状只有「按代码顺序」那一个答案, 出厂答错的下限是静默不漏,
' 而 free 答错就是拿错步长解引用 ⇒ 载体级的销毁 (= 内存那半, 全元素档一起) 是另一格机制, 在册。
Private Sub W36OneDyn()
    Dim loc() As C445
    ReDim loc(1 To 2)
    Set loc(1) = New C445
    Set loc(2) = New C445
End Sub

Private Sub W36OneFix()
    Dim lf(1 To 2) As C445
    Set lf(1) = New C445
    Set lf(2) = New C445
End Sub

Private Sub W36OneNotClass()
    Dim lg() As Long
    ReDim lg(1 To 2)
    lg(1) = 7
End Sub

Sub W36LocalArrayScope()
    Dim t0 As Long
    Dim i As Long
    t0 = Terms
    For i = 1 To 3
        W36OneDyn
        W36OneFix
        W36OneNotClass
    Next
    Debug.Print "W36-terms=" & (Terms - t0)
End Sub

' ==== 账 #305 §B218 (甲) 第一百五十四刀 = 多维 Preserve **收缩**那一格: 掉出界的那几格也逐枚交回 ====
' 修前: 丢尾那几格的引用连旧块一起被 free (探针 A10 = 2x3 缩到 2x1 退 0 枚, .build/k1846_a10_p64.txt);
' 修后: 同一趟丢尾循环认指针档 ⇒ 2 枚 (那两枚留下来的不在这一发里退, 由紧随的 Erase 交回)。
' 同格里再走一发 As Long 的多维 Preserve = 反面证人: 那一发继续走转调, 产物一字不动。
Sub W37NDPreserveShrink()
    Dim t0 As Long
    t0 = Terms
    ReDim mND445(1 To 2, 1 To 3)
    Set mND445(1, 1) = New C445
    Set mND445(1, 2) = New C445
    Set mND445(2, 1) = New C445
    Set mND445(2, 3) = New C445
    ReDim Preserve mND445(1 To 2, 1 To 1)
    Debug.Print "W37-terms=" & (Terms - t0) & " kept=" & (Not (mND445(1, 1) Is Nothing))
    Erase mND445
    ReDim mNDLng445(1 To 2, 1 To 3)
    ReDim Preserve mNDLng445(1 To 2, 1 To 1)
    Erase mNDLng445
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
    W20ArrElemOverwrite
    W21ArrElemNothing
    W22IndexSideEffectOnce
    W23ArrFieldTeardown
    W24OutParamWriteBack
    W25ChainCarriesBack
    W26OutParamNil
    W27ByValDoesNotLand
    W28HandedOverArg
    W29EraseFixedArr
    W30EraseDynArr
    W31RedimNoPreserve
    W32PreserveKeepsElems
    W33NonClassCarriers
    W34NdCarrierUntouched
    W35PreserveShrink
    W36LocalArrayScope
    W37NDPreserveShrink
    Debug.Print "B445-DONE"
End Sub
