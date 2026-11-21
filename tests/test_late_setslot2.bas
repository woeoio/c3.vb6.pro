Attribute VB_Name = "test_late_setslot2"
' 第一百一十一刀 (§11.93): Set 的目标 —— 解引用槽 (*w) 与 With 块里的 UDT Variant 字段
' 真机（probe111 / Q 族）:
'   Q1  Set w = o.NoSuchProperty    (ByRef Variant 目标)  = v=111   err=438
'   Q1c Set w = o                   (控制格)              = isobj=True err=0
'   Q2  Set uv.f = o.NoSuchProperty (UDT 字段, 参照格)     = uv.f=111 err=438
'   Q3  With uv: Set .f = o.NoSuchProperty                 = uv.f=111 err=438
'   Q3c With uv: Set .f = o         (控制格)              = isobj=True err=0
Option Explicit

Public Type UV
    f As Variant
End Type

Private Sub PE(ByVal tag As String, ByVal desc As String, ByVal val As String, ByVal errNum As Long)
    Debug.Print tag & " = " & desc & " = [" & val & "]  err=" & errNum
End Sub

' Q1: Set 语句的目标是一枚 ByRef Variant 形参（C 里是 (*w)）
Private Sub LateSetToParam(ByRef w As Variant, ByVal o As Object)
    Set w = o.NoSuchProperty
End Sub

Private Sub SetObjToParam(ByRef w As Variant, ByVal o As Object)
    Set w = o
End Sub

Public Sub Main()
    On Error Resume Next
    Dim o As Object
    Dim v As Variant
    Dim uv As UV
    Dim e As Long
    Set o = CreateObject("Scripting.Dictionary")
    o.Add "k", "vv"

    ' Q1: 晚绑定读做右值，目标是解引用槽
    v = 111
    Err.Clear
    LateSetToParam v, o
    e = Err.Number
    PE "Q1", "Set w = o.NoSuchProperty (ByRef Variant)", "v=" & CStr(v), e

    ' Q1c: 控制格 —— 成功的 Set 必须写回调用方的 v
    v = 111
    Err.Clear
    SetObjToParam v, o
    e = Err.Number
    PE "Q1c", "Set w = o (control)", "isobj=" & CStr(IsObject(v)), e

    ' Q2: 参照格（第一百零七刀已修）
    uv.f = 111
    Err.Clear
    Set uv.f = o.NoSuchProperty
    e = Err.Number
    PE "Q2", "Set uv.f = o.NoSuchProperty (UDT field)", "uv.f=" & CStr(uv.f), e

    ' Q3: With 块里的 UDT Variant 字段
    uv.f = 111
    Err.Clear
    With uv
        Set .f = o.NoSuchProperty
    End With
    e = Err.Number
    PE "Q3", "With uv: Set .f = o.NoSuchProperty", "uv.f=" & CStr(uv.f), e

    ' Q3c: 控制格 —— With 块里成功的 Set 必须落槽
    uv.f = 111
    Err.Clear
    With uv
        Set .f = o
    End With
    e = Err.Number
    PE "Q3c", "With uv: Set .f = o (control)", "isobj=" & CStr(IsObject(uv.f)), e

    Debug.Print "---- setslot2 done ----"
End Sub
