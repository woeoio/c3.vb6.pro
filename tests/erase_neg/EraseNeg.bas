Attribute VB_Name = "EraseNeg"
Option Explicit

' 账 #186 的**边界负例**: `Erase m_list(1)` 这一形 VB6 只在元素是 Variant(装着数组) 时合法。
' 本仓没有那条通路 —— 静默降级成"销毁整个数组"会把语义改掉 (别的元素的数据全没了),
' 所以这一句必须继续报诊断。修完之后这条负例是本账的"不许过头"那一头。
Private m_list(3) As Long

Sub Main()
    Erase m_list(1)
    Debug.Print "SHOULD-NOT-COMPILE"
End Sub
