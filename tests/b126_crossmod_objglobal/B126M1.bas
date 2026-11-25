Option Explicit

' 标准模块的 Public Object 变量 (Fix 089 那一臂): 消费模块必须把它当 void* 晚绑定,
' 不是当裸名不存在。类模块的 Public 字段走的是同一条注入路, 但那是成员, 不进名字表
' (本夹具同时钉住这两面: 上面这条留得住, 类字段那条不许流进来)。
Public gObj As Object

Sub PrepareIt()
    Dim it As B126Item
    Set it = New B126Item
    it.Seed = 7
    Set gObj = it
End Sub