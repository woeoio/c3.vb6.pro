Option Explicit

' 账 #305 第二百六十二刀 (丙): 晚绑定出口的「非 IDispatch 接收者」护栏 —— 三种形状钉在**一个**计数里。
' 形状一 (宿主伪对象): RTL 内建 Collection 不是 IDispatch, 但宿主模型应答 ⇒ **0 漏**。
' 形状二 (工程类进 Object 槽): 应发包装器 ⇒ **0 漏**; 这里要是响了一声, 那就是晚绑定真落空了。
' 形状三 (负控): Nothing 接收者 ⇒ **恰好 1 漏** —— 护栏不响, 就是护栏自己坏了。
Public gObj As Object

Sub Main()
    Dim col As Object
    Set col = New Collection
    Dim nCnt As Long
    nCnt = col.Count
    Debug.Print "B446-coll=" & nCnt

    Set gObj = New B446Cls
    Dim sPing As String
    sPing = gObj.Ping()
    Debug.Print "B446-ping=" & sPing

    Dim z As Object
    On Error Resume Next
    z.Ping()
    On Error GoTo 0
    Debug.Print "B446-DONE"
End Sub
