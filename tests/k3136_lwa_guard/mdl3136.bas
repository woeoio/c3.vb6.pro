Attribute VB_Name = "mdl3136"
Option Explicit

' 账 #305 第一百九十八刀 族A 夹具: 同一个模块里两条 Declare 指向 kernel32 的同一枚导出
' (VB6 的常见写法) —— 先发那条占了 diGuard, 别名那条整块曾被预处理器吃掉, 调用点因此
' 指向一份不存在的 static __inline => 链接期 LNK2019。
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Public Sub Main()
    Dim src(3) As Byte
    Dim dst(3) As Byte
    src(0) = 65
    src(1) = 66
    src(2) = 67
    src(3) = 68
    RtlMoveMemory dst(0), src(0), 2
    CopyMemory dst(2), src(2), 2
    Debug.Print "K3136-a=" & CStr(dst(0) + dst(1))
    Debug.Print "K3136-b=" & CStr(dst(2) + dst(3))
    Debug.Print "K3136-DONE"
End Sub
