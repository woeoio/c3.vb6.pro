' Negative case (VB3043): UBound 的首参在 VB6 必须是数组表达式。
' vbNull 是 VarType 常量 (=1), 真 VB6 里这条是编译期错误 (提示缺少数组),
' 不是某个返回值 —— C3 修复前把它折成整数塞进 SafeArray1D* 形参, 运行期解引用地址 1。
Option Explicit

Public Sub Main()
    Debug.Print UBound(vbNull)
End Sub
