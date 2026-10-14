VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CImpl"
Option Explicit

Implements IProbe

Private m_got As Long

Private Sub Ping(ByVal n As Long)
    m_got = n * 2
End Sub

Private Property Get Got() As Long
    Got = m_got
End Property

' ai/022 B17: 一个**公有成员** —— 接口槽之外的那一半对外面。契约成员是 Private（VB6 语义里
' 接口实现不进类的公有面），所以纯接口宿主在 IDispatch 晚绑定下按名字一个成员也点不到；
' 外部 `CreateObject` + 按名调用的验收（test_p613 那一族）要有东西可点，就得有公有成员。
' 同一个类同时有两个面这件事本身也值得钉住：晚绑定走 desc->methods（这条），
' 早绑定走 IProbe 的契约槽（Ping/Got）。
Public Function Twice(ByVal n As Long) As Long
    Twice = n * 2
End Function

' Acct #305 SB144: three members that push a Decimal ACROSS the COM boundary.
' OLE keeps the 16-byte DECIMAL at the HEAD of the VARIANT on both arches, while our
' slot keeps it inline at offset 8 (x64) or behind pdecVal (x86). The server adapter
' used to memcpy the slot's first 16 bytes in and out, which answered 0 on x64 and a
' heap address on x86 without raising anything. Amt/Quarter cover the outbound
' direction, EchoVal the inbound one (the echo also proves the borrow does not steal
' the caller's bytes: the client clears its own VARIANT afterwards).
Public Function Amt() As Variant
    Amt = CDec(1234567)
End Function

Public Function Quarter() As Variant
    Quarter = CDec("1.25")
End Function

Public Function EchoVal(ByVal v As Variant) As Variant
    EchoVal = v
End Function
