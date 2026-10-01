' Negative case (VB3043): Join 的首参同样是数组槽 (RTL: vb6_Join(SafeArray1D*, BSTR))。
' 传 VarType 常量在真 VB6 里是编译期错误; C3 修复前折成整数塞进指针形参 → 段错误。
Option Explicit

Public Sub Main()
    Debug.Print "["; Join(vbNull, ","); "]"
End Sub
