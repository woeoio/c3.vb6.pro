' 账 #116：`Dim X As String * N` 的 typeRef 节点是 FixedStringTypeRef，而模块级变量发码
' 处（cgen_decl_var.cpp:395 / cgen_decl_func.cpp:204 / cgen_decl_prop.cpp:209）无条件
' `static_cast<SimpleTypeRef*>(...)->name` —— 那是把节点里的 `ExprPtr length` 当成
' std::string 读，"长度"其实是一个堆指针 ⇒ operator new 张口要几十 GB ⇒ std::bad_alloc
' 没人接 → abort()（退出码 3、零诊断，调试版 CRT 还弹一个前台模态框把会话卡住）。
' 本用例只钉"这六种定长串落点都能正常发码"，刻意**不**问 VB6 的空格补齐语义
' （那条还欠着，见 ai/029 账 #116 收线里的"还欠的一格"）。
Option Explicit

Dim gT As String * 8
Dim gArr(3) As Long      ' 数组: 空值必须是 0 (即空指针), 不能是 vb6_VariantEmpty()

Type Rec
    f As String * 8
    n As Long
End Type

Function F() As String * 5
    F = "ab"
End Function

Sub S(ByVal x As String * 4)
    Debug.Print x
End Sub

Sub Main()
    Dim lT As String * 4
    Dim r As Rec
    lT = "cd"
    gT = "ab"
    r.f = "ef"
    r.n = 1
    Debug.Print gT
    Debug.Print Len(gT)
    Debug.Print lT
    Debug.Print r.f
    Debug.Print F()
    S "gh"
End Sub
