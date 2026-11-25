Attribute VB_Name = "test_late_varslot2"
' 第一百零七刀 (§11.85): Variant 槽的赋值位 —— 表达式内部 / 数组元素 / UDT 字段
' 真机 L6/L7/L8/L9 量下来的都是「111 / 438」: 目标一个字都不写。
Option Explicit

Private Type TV
    f As Variant
End Type

Private gHit As Long

Private Sub PE(ByVal tag As String, ByVal desc As String, ByVal val As String, ByVal errNum As Long)
    Debug.Print tag & " = " & desc & " = [" & val & "]  err=" & errNum
End Sub

Public Sub Main()
    On Error Resume Next
    Dim o As Object
    Dim v As Variant
    Dim s As String
    Dim e As Long
    Dim va(0 To 1) As Variant
    Dim uv As TV
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"

    ' L6: 晚绑定读在算术表达式内部 —— 真机 v=111
    v = 111
    Err.Clear
    v = 1 + o.NoSuchProperty
    e = Err.Number
    PE "L6", "v = 1 + o.NoSuchProperty", "v=" & CStr(v), e

    ' L7: 晚绑定读在串拼接内部 —— 真机 v=111
    v = 111
    Err.Clear
    v = "x" & o.NoSuchProperty
    e = Err.Number
    PE "L7", "v = ""x"" & o.NoSuchProperty", "v=" & CStr(v), e

    ' L8: Variant 数组元素做目标 —— 真机 va(0)=111
    va(0) = 111
    Err.Clear
    va(0) = o.NoSuchProperty
    e = Err.Number
    PE "L8", "va(0) = o.NoSuchProperty", "va(0)=" & CStr(va(0)), e

    ' L9: UDT 的 Variant 字段做目标 —— 真机 uv.f=111
    uv.f = 111
    Err.Clear
    uv.f = o.NoSuchProperty
    e = Err.Number
    PE "L9", "uv.f = o.NoSuchProperty", "uv.f=" & CStr(uv.f), e

    ' 控制格: 读成功必须照样落到目标 (真机 v=1 / v=vv)
    v = 111
    Err.Clear
    v = o.Count
    e = Err.Number
    PE "L11", "v = o.Count", "v=" & CStr(v), e

    v = 111
    Err.Clear
    v = o.Item("k")
    e = Err.Number
    PE "L12", "v = o.Item(""k"")", "v=" & CStr(v), e

    ' 控制格: 数组元素/UDT 字段上的**正常**赋值不能被守卫挡掉
    va(1) = 111
    va(1) = o.Count
    PE "L8b", "va(1) = o.Count", "va(1)=" & CStr(va(1)), Err.Number

    uv.f = 111
    uv.f = o.Count
    PE "L9b", "uv.f = o.Count", "uv.f=" & CStr(uv.f), Err.Number

    s = "done"
    Debug.Print "---- varslot2 done: " & s
End Sub
