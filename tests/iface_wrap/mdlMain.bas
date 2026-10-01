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
    Debug.Print "IFACE-DONE"
End Sub
