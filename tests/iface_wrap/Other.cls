Attribute VB_Name = "Other"
Option Explicit

' 负控用类: 与 IFoo 毫无关系 (没有 Implements)。真 VB6 探针 (.temp/vb6probe_neg)
' 给这一形的真值是**运行期错误 13 Type mismatch**, 且失败后接口变量保持原值。
Public Function Bar() As Long
    Bar = -1
End Function
