Attribute VB_Name = "ToolsA"
Option Explicit

' C2197 第二族夹具（正面）：4 参的 NotifyMouseMove。
' ToolsB 里有一枚同名但 6 参的过程; 跨模块注入按裸名去重只留先注册者,
' 所以这里的调用点会借到 ToolsB 的表 => 补两枚缺省 => C2197。
Public Sub NotifyMouseMove(ByVal b As Boolean, ByVal sh As Integer, ByVal x As Single, ByVal y As Single)
End Sub
