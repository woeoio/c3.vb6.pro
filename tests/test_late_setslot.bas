Attribute VB_Name = "test_late_setslot"
' 第一百零九刀 (§11.89): Set 进 Variant 数组元素 —— 真机 N1 = va(0) 仍是 111 / 438
Option Explicit

Private Sub PE(ByVal tag As String, ByVal desc As String, ByVal val As String, ByVal errNum As Long)
    Debug.Print tag & " = " & desc & " = [" & val & "]  err=" & errNum
End Sub

Public Sub Main()
    On Error Resume Next
    Dim o As Object
    Dim va(0 To 1) As Variant
    Dim e As Long
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"

    ' N1: Set 进 Variant 数组元素 —— 真机 va(0) 仍是 111
    va(0) = 111
    Err.Clear
    Set va(0) = o.NoSuchProperty
    e = Err.Number
    PE "N1", "Set va(0) = o.NoSuchProperty", "va(0)=" & CStr(va(0)), e

    ' N1b: 下标是表达式（同一道闸，不得求值两次）
    va(0) = 111
    Err.Clear
    Set va(0 + 0) = o.NoSuchProperty
    e = Err.Number
    PE "N1b", "Set va(0 + 0) = o.NoSuchProperty", "va(0)=" & CStr(va(0)), e

    ' N3a: 控制格 —— 成功的 Set 必须落进数组元素
    va(1) = 111
    Err.Clear
    Set va(1) = o
    e = Err.Number
    PE "N3a", "Set va(1) = o", "va(1) nothing=" & CStr(IsObject(va(1)) = False), e

    Debug.Print "---- setslot done ----"
End Sub
