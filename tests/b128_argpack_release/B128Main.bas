Option Explicit

' 第一百二十八刀 (§11.110): 晚绑定实参"打包了、却只 free 结构体不清内容"那一族。
' 打包 (vb6_ComPackObject / vb6_ComPackVB6InstanceRaw 等) 交出一枚 VARIANT,
' 结构体是堆上自造的, 内容里那一枚引用 (与 BSTR) 只有 VariantClear 会退。
Public gTerms As Long

' 甲: 宿主对象那一岔 —— 内建 Collection 是宿主对象 (vb6_uc_isColl),
'     `c.Add o` 落进 vb6_Host_Call 那一支。修前那一支收尾只 free 打包的结构体
'     ⇒ 每发 Add 漏一枚引用, 元素到 Remove 也退不到 0, Class_Terminate 永不响。
Private Sub HostArm()
    Dim c As Collection
    Dim o As Object
    Set c = New Collection
    Set o = New B128Item
    c.Add o
    Set o = Nothing
    c.Remove 1
End Sub

' 乙: 真 IDispatch 那一岔 —— Dictionary 的 Item 赋值走 vb6_ComSetPropArg。
'     修前值那一枚与索引参数的结构体都是裸 free, 值的内容从没清过 ⇒ 同一族漏。
Private Sub DispatchArm()
    Dim d As Dictionary
    Dim o As Object
    Set d = New Dictionary
    Set o = New B128Item
    d("k") = o
    Set o = Nothing
    d.Remove "k"
End Sub

Sub Main()
    HostArm
    Debug.Print "C3ProbeB128-A1-terms=" & CStr(gTerms)
    DispatchArm
    Debug.Print "C3ProbeB128-A2-terms=" & CStr(gTerms)
    Debug.Print "C3ProbeB128-DONE"
End Sub