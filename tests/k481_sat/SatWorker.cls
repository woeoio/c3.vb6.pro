VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SatWorker"
Option Explicit

' 账 #481 = 第一百八十二刀 (注销腿) 的运行时证人 (第一百八十三刀接上)。
' 这一发在 DLL 进程内: 同时建 n 枚 SatImpl (并发存活 ⇒ 地址各不相同, 表真的会长),
' 再全部拆掉 (修后: 注销腿把格腾出来), 最后建一枚**另一个类** SatOther, 经视图装进
' Variant, 用 TypeName 问身份。
'   修后: 表里没有残留 ⇒ 最后一枚登记得上 ⇒ 答 "SatOther"。
'   修前: 计数只增不减 (n > 512 ⇒ 表饱和), 最后一枚的地址若是那 88 个没登记进去的之一,
'         表里查不到 ⇒ 答空; 若地址恰好命中一条**陈旧的 SatImpl 记录**, 修前会答 "SatImpl"
'         —— 两种错答都被这一枚针钉住 (断言 = 必须 "SatOther")。
Public Function Churn(ByVal n As Long) As String
    Dim i As Long
    Dim arr() As SatImpl
    ReDim arr(1 To n)
    For i = 1 To n
        Set arr(i) = New SatImpl
    Next i
    For i = 1 To n
        Set arr(i) = Nothing
    Next i
    Dim last As SatOther
    Set last = New SatOther
    Dim view As SatIface
    Set view = last
    Dim v As Variant
    Set v = view
    Churn = TypeName(v)
End Function
