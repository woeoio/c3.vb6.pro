Attribute VB_Name = "Module1"
Option Explicit

' B5 回归: 本工程类实例**变量**作为晚绑定 COM 实参 (Dictionary.Add).
' 修复前: 变量槽里的裸 vb6_cls_bHello* 被 vb6_ComPackValue 当 VT_DISPATCH
' 原样交出, 对端 AddRef 读结构体首字段 __comObj 当 vtable ⇒ 0xC0000005.
' 修复后: 借用型包装 (FromBorrowedInstance 加引用, 实例销毁权仍归 h).
' 后半段 d.Remove 后继续 h.Ping —— 若误用 FromInstance (所有权转移),
' Remove 释放 VARIANT 就会销毁实例, 这里必崩 ⇒ 钉死"借用"而非"转移".
Sub Main()
    Dim d As Object
    Dim h As bHello
    Set d = CreateObject("Scripting.Dictionary")

    Set h = New bHello
    d.Add "x", h
    Debug.Print "ADD-OK"

    d.Remove "x"
    Debug.Print "REMOVED"
    h.Ping
    Debug.Print "HOST-ALIVE"

    Dim o As Object
    Set h = New bHello
    Set o = h
    d.Add "y", h
    Debug.Print "ADD2-OK"

    Set o = d("y")
    Dim s As String
    s = CallByName(o, "Echo", VbMethod, "ok")
    Debug.Print "ECHO=" & s

    d.Remove "y"
    h.Ping
    Debug.Print "HOST2-ALIVE"

    Set h = Nothing
    Set o = Nothing
    Set d = Nothing
    Debug.Print "DONE"
End Sub
