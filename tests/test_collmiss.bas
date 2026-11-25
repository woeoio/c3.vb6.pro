Attribute VB_Name = "test_collmiss"
' 第一百一十三刀 (§11.95): 内建 Collection 的未知名成员必须报 438，且不写目标
' 真机 probe112 P3: `Set v = col.NoSuchProperty` = v=[111] err=438
' 修前: 宿主模型对 Collection 的任何未知成员都 `return 1`（答了但没答）⇒
'       vb6_ComGetPropRcv 的 `hgot < 0` 438 收口走不到，本机 v=[] err=0（还把目标写了）。
Option Explicit

Public Sub Main()
    Dim col As Object
    Dim v As Variant
    Dim e As Long

    Set col = New Collection
    col.Add 5

    ' ① 已知成员: Count / Item 照旧
    Debug.Print "CM-1 count=" & CStr(col.Count)
    Debug.Print "CM-2 item=" & CStr(col.Item(1))

    ' ② 未知名成员: 438，且目标保持 111（438 不写目标）
    v = 111
    Err.Clear
    On Error Resume Next
    Set v = col.NoSuchProperty
    e = Err.Number
    On Error GoTo 0
    Debug.Print "CM-3 v=[" & CStr(v) & "] err=" & CStr(e)

    ' ③ 报过错以后已知成员仍然照常（不是一次性坏掉）
    Err.Clear
    On Error Resume Next
    Debug.Print "CM-4 count=" & CStr(col.Count) & " err=" & CStr(Err.Number)
    On Error GoTo 0

    ' ④ 换一个名字再报一次（不是「只报第一次」），目标仍保持不动
    v = 222
    Err.Clear
    On Error Resume Next
    Set v = col.AnotherMissingName
    Debug.Print "CM-5 v=[" & CStr(v) & "] err=" & CStr(Err.Number)
    On Error GoTo 0

    Debug.Print "COLLMISS-DONE"
End Sub
