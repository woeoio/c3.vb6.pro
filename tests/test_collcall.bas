Attribute VB_Name = "test_collcall"
' 第一百一十五刀 (§11.97): 内建 Collection 的**调用臂** —— Count / 默认成员 / 未知名 438
' 真机 probe113 读数 (VB6 IDE, 2026-11-22):
'   Q2  v = col.Count           => v=[1] err=0 VarType=3(VT_I4)
'   Q2b v = col.Item(1)         => v=[5] err=0 VarType=2(VT_I2; 5 是 Integer 字面量)
'   Q2c v = col.NoSuchProperty  => v=[111] err=438 (普通赋值读未知名, 目标不动)
'   Q1  Set v = col.NoSuchProperty => v=[111] err=438 (getprop 臂, §11.95/§11.96 已收口)
' 修前: `v = col.Count` 与 `With <UDT>: .Prop(1)` 都发 vb6_ComCall, 而宿主模型 Collection 的
'       调用臂没有 Count、不认空名字 (默认成员)、未知名还 return 1 —— 三格全答 Empty/err=0。
Option Explicit

Public Type Holder
    Prop As Collection
End Type
Public h() As Holder

Public Sub Main()
    Dim col As Object
    Dim v As Variant
    Dim e As Long

    Set col = New Collection
    col.Add 5, "a"

    ' ① Count 走调用臂 (v = col.Count 发的是 vb6_ComCall, 不是 ComGetProp)
    v = col.Count
    Debug.Print "CC-1 count v=[" & CStr(v) & "] VarType=" & CStr(VarType(v))

    ' ② Item 已知成员: 值对。VarType 本机 3(VT_I4) / 真机 2(VT_I2) 是整数字面量归档的
    '    既有口径差异, 本条针只立"搬到了槽里且值对"。
    v = 111
    v = col.Item(1)
    Debug.Print "CC-2 item v=[" & CStr(v) & "]"

    ' ③ 未知名成员用普通赋值 (Q2c): 438, 目标保持 111
    v = 111
    Err.Clear
    On Error Resume Next
    v = col.NoSuchProperty
    e = Err.Number
    On Error GoTo 0
    Debug.Print "CC-3 v=[" & CStr(v) & "] err=" & CStr(e)

    ' ④ 报过错以后 Count 照常 (不是一次性坏掉)
    Err.Clear
    On Error Resume Next
    Debug.Print "CC-4 count=" & CStr(col.Count) & " err=" & CStr(Err.Number)
    On Error GoTo 0

    ' ⑤ With <UDT>: .Prop(1) —— 空方法名 = 默认成员 Item (复现器 .temp/k113/wuc2.bas 的 H2)
    ReDim h(0)
    Set h(0).Prop = New Collection
    h(0).Prop.Add 7, "k"
    With h(0)
        Debug.Print "CC-5 withitem=[" & CStr(.Prop(1)) & "] count=" & CStr(.Prop.Count)
    End With

    ' ⑥ 语句式未知名调用: 同样 438 (调用臂 argc=0 那一档)
    Err.Clear
    On Error Resume Next
    col.AnotherMissing
    e = Err.Number
    On Error GoTo 0
    Debug.Print "CC-6 stmt err=" & CStr(e)

    Debug.Print "COLLCALL-DONE"
End Sub
