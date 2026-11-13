Attribute VB_Name = "ModProbe005a"
Option Explicit

' 账 #610 第342刀（富贵竹 FIX-005a 那一族）的正式夹具。
'
' 上游 FIX-005a 原文: "valueIsBareComResult 的内联名单长度全错 … 11 条里只有前两条长度
' 正确, 其余全部少写一格, 等价于整张表失效 ⇒ 改成调用既有主人 cExprIsBareComResult()"。
'
' ⚠ 本地那张内联名单(src/backend/detail/stmt/cgen_assign_value_sem.inc)不是"全部少写一格",
'   而是**只有一条写错格**:
'       value.compare(0, 20, "vb6_ComCallByDispid(")   ← 字面量 19 字符, 拿 20 去比 → 恒假
'   ⇒ 这一枚(VARIANT* 结果里唯一带 ByDispid 的)永远不被当裸 COM 结果 ⇒ 赋给数值/字符串
'   槽时**不解包**, 把 void* 直接递给 int32_t / BSTR 形参 ⇒ C2440。
'   这就是上游那一族的本地同款 bug(同源不同形)。
'
' 判据: 修前 C2440(无法从 void* 转换为 int32_t) ⇒ 编不出 exe; 修后走
'       vb6_VariantToLong(vb6_VariantFromComResult(vb6_ComCallByDispid(...))) ⇒ 链通 + 跑通。

Public Function ProbeDefaultToLong() As Long
    Dim d As Object
    Dim v As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", 7
    v = d("k")                     ' 后期绑定默认成员 (DISPID_VALUE) → vb6_ComCallByDispid
    ProbeDefaultToLong = v
End Function

Public Function ProbeDefaultToString() As String
    Dim d As Object
    Dim s As String
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", "hello"
    s = d("k")                     ' 同上, 目标是 BSTR 槽
    ProbeDefaultToString = s
End Function

Public Function ProbeDefaultToDouble() As Double
    Dim d As Object
    Dim x As Double
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k", 2.5
    x = d("k")
    ProbeDefaultToDouble = x
End Function

Public Sub Main()
    Dim n As Long
    Dim s As String
    Dim x As Double
    n = ProbeDefaultToLong()
    s = ProbeDefaultToString()
    x = ProbeDefaultToDouble()
    Debug.Print "DICT-L=" & n
    Debug.Print "DICT-S=" & s
    Debug.Print "DICT-X=" & x
End Sub
