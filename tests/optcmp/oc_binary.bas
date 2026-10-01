Option Explicit

' <vbeclipse>: 本模块**没有** Option Compare ⇒ 恒 Binary。同工程的 oc_main.bas 是
' Option Compare Text, 两个模块编进同一个 exe: 如果模式是靠进程唯一的全局来切的,
' 这里就会被带跑成 "TEXT"。这条探针就是按模块隔离的负控。
Public Function OCBinaryProbe() As String
    Dim a As String
    a = "abc"
    If a = "ABC" Then
        OCBinaryProbe = "TEXT"
    Else
        OCBinaryProbe = "BINARY"
    End If
End Function
