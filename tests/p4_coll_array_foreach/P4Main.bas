Attribute VB_Name = "P4Main"

' P4 取证回归 (2026-11-04): 用户报的 "P4.exe 内存不能 read"。
' 形态与 .build/probe2/P4.bas 的 F5 逐字同型: Variant() 数组里装一枚 C3 伪对象
' (内建 Collection, 首字段是身份标记 0xC01C01C0, 不是 vtable), 整只数组再装进 Variant,
' 然后 For Each 遍历。改前: 循环尾清栈上 VARIANT 走 OLEAUT32 VariantClear, 它把伪对象
' 按 COM 协议解引用 vtable ⇒ 0xC0000005, 故障读地址 = 标记 + 槽偏移 (用户截图 0xC01C01C4)。
' 改后: 清理由 vb6_VariantClearWin 的对象档按 vb6_ReleaseObject 同一判据处理, 伪对象跳过。
Public Sub Main()
    Dim dv(0 To 1) As Variant
    dv(0) = DateSerial(2024, 1, 2)
    Set dv(1) = New Collection

    Dim v As Variant
    v = dv

    Dim x As Variant
    Dim n As Long
    Dim tag As String
    For Each x In v
        n = n + 1
        tag = tag & TypeName(x) & ";"
    Next

    Debug.Print "P4 count=" & n
    Debug.Print "P4 types=" & tag
    Debug.Print "DONE"
End Sub
