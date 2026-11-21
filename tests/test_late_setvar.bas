Attribute VB_Name = "test_late_setvar"
' 第一百零八刀 (§11.87): Set 赋值位 + ByRef Variant 目标
' 真机 M2/M4 都是 111 / 438（目标一个字都不写）；M1 也是 True / 438；M3 是控制格。
Option Explicit

Private Sub PE(ByVal tag As String, ByVal desc As String, ByVal val As String, ByVal errNum As Long)
    Debug.Print tag & " = " & desc & " = [" & val & "]  err=" & errNum
End Sub

' M4: 目标是一枚 ByRef 的 Variant 形参（C 里是 (*w)）
Private Sub LateToParam(ByRef w As Variant, ByVal o As Object)
    w = o.NoSuchProperty
End Sub

Public Sub Main()
    On Error Resume Next
    Dim o As Object
    Dim o2 As Object
    Dim v As Variant
    Dim e As Long
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"

    ' M1: Set 到 Object 变量 —— 真机 o2 仍是 Nothing
    Set o2 = Nothing
    Err.Clear
    Set o2 = o.NoSuchProperty
    e = Err.Number
    PE "M1", "Set o2 = o.NoSuchProperty", "o2 is nothing=" & CStr(o2 Is Nothing), e

    ' M2: Set 到 Variant 目标 —— 真机 v 仍是 111
    v = 111
    Err.Clear
    Set v = o.NoSuchProperty
    e = Err.Number
    PE "M2", "Set v = o.NoSuchProperty", "v=" & CStr(v), e

    ' M3: 控制格 —— 成功的 Set 必须落到目标
    Set o2 = Nothing
    Err.Clear
    Set o2 = o
    e = Err.Number
    PE "M3", "Set o2 = o", "o2 is nothing=" & CStr(o2 Is Nothing), e

    ' M4: ByRef Variant 形参做目标 —— 真机调用方的 v 仍是 111
    v = 111
    Err.Clear
    LateToParam v, o
    e = Err.Number
    PE "M4", "LateToParam v", "v=" & CStr(v), e

    ' 控制格: 正常 Set 到 Variant（读成功）不许被守卫挡掉
    v = 111
    Err.Clear
    Set v = o
    e = Err.Number
    PE "M5", "Set v = o", "v is nothing=" & CStr(IsObject(v) = False), e

    Debug.Print "---- setvar done ----"
End Sub
