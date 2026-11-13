Attribute VB_Name = "mdlMain"
Option Explicit

' <vbeclipse> 回归夹子: 接口宿主类型映射。IFoo 显式 VB_Creatable=True ⇒ 进不了
' ivref 登记表, `Dim f As IFoo` 必须经 Class/isInterface 分支发 vb6_iface_IFoo
' 包装类型; Set 侧对 Variant 里的实现类实例发 vb6_iface_IFoo_wrap。两侧判据一旦
' 分叉 (rev7 首版曾让工程类名短路劫走声明侧 ⇒ 裸 vb6_cls_IFoo*), 这里 C2440 当场红。
Sub Main()
    Dim f As IFoo
    Dim o As Impl
    Set o = New Impl
    Dim v As Variant
    Set v = o
    Set f = v
    If f.Bar() = 7 Then
        Debug.Print "IFACE-CALL-OK"
    Else
        Debug.Print "IFACE-CALL-BAD"
    End If
    ' 负控 (真 VB6 探针 .temp/vb6probe_neg 定真值): Variant/Object 里装的是
    ' **没实现 IFoo 的类** 或 **外部对象** ⇒ 运行期 13 Type mismatch, 整条 Set 被放弃,
    ' 目标保持原值 (不是静默 Nothing)。
    Dim vNeg As Variant
    On Error Resume Next
    Set vNeg = New Other
    Err.Clear
    Set f = vNeg
    Debug.Print "IFACE-NEG-KIND=" & CStr(Err.Number)
    Err.Clear
    Set vNeg = CreateObject("Scripting.Dictionary")
    Set f = vNeg
    Debug.Print "IFACE-NEG-EXT=" & CStr(Err.Number)
    Debug.Print "IFACE-NEG-KEEP=" & CStr(f.Bar())
    On Error GoTo 0
    Debug.Print "IFACE-DONE"
End Sub
