Attribute VB_Name = "B597Main"
Option Explicit

' ByRef 工程类形参作为晚绑定 COM 实参 (B5 同源缺口).
' 修复前实测: d.Add "k", obj (obj 为 `ByRef ... As bHello`, C 侧 `vb6_cls_bHello**`)
' 发成 vb6_ComPackValue((*obj)) → 裸实例指针当 VT_DISPATCH → 0xC0000005,
' faulting module OLEAUT32.dll +0x9FF30, 故障读地址 0x8.
' 断言钉死: 借用打包 (FromBorrowedInstance) —— 对端 Remove 后宿主槽仍握有销毁权.
Sub Main()
    Dim d As Object
    Dim h As bHello
    Set d = CreateObject("Scripting.Dictionary")
    Set h = New bHello

    Debug.Print "S1 before call"
    PassIt d, h
    Debug.Print "S2 after add"

    Dim o As Object
    Set o = d("k")
    Debug.Print "S3 got object"
    CallByName o, "Ping", VbMethod

    d.Remove "k"
    Debug.Print "S4 after remove"
    h.Ping
    Debug.Print "DONE"
End Sub

Sub PassIt(ByRef d As Object, ByRef obj As bHello)
    d.Add "k", obj
End Sub
