Attribute VB_Name = "Module1"
Option Explicit

' 免注册 COM 正例：没有任何 ComLib= 声明，CreateObject 的 ProgID 只存在于
' 客户端 manifest 的 dependentAssembly 里 —— 未注册的 test_activex_dll 被 SxS 解析出来。
' manifest 模式必须有 ActCtx 才成；direct 模式下这条 ProID 必然失败（作为对照）。
Sub Main()
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("TestAXDLL.Calc")
    If Err.Number = 0 Then
        Debug.Print "manifest-create=OK"
    Else
        Debug.Print "manifest-create=FAIL"
    End If
End Sub
