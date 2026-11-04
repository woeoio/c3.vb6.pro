Attribute VB_Name = "ModuleDirect"
Option Explicit

' 直加载回归：ComLib= 把 test_activex_dll 烘焙进 c3_comlibs 表，
' CreateObject 走 LoadLibrary + DllGetClassObject（不经过注册表、不需要 ActCtx）。
' 这条在 direct 模式下必须成；manifest 模式下也不受影响（表优先）。
Sub Main()
    Dim d As Object
    On Error Resume Next
    Set d = CreateObject("TestAXDLL.Calc")
    If Err.Number = 0 Then
        Debug.Print "direct-create=OK"
    Else
        Debug.Print "direct-create=FAIL"
    End If
End Sub
