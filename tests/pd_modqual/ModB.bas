Attribute VB_Name = "ModB"
Option Explicit

' C2197 钉子（反面）：本模块有一枚同名但带一枚必填参的过程 GetLevel。
' Probe 里叫的是 ModA.GetLevel() —— ModA 那一枚是 0 参, 实参表必须是空的。
' 修前: lookupModule 先命中本模块这枚 1 参过程 ⇒ 补一枚缺省 0 ⇒ C2197。
Public Function GetLevel(ByVal fmt As Long) As Long
    GetLevel = fmt
End Function

Public Function Probe() As Long
    Probe = ModA.GetLevel()
End Function

' C2197 第二族（跨模块同名成员）：ToolsA 那枚 4 参, ToolsB 有一枚 6 参的同名。
' 修前形参表借到注册在先的 ToolsB => 多补两枚缺省 => C2197。
Public Sub ProbeMouse(ByVal down As Boolean, ByVal sh As Integer, ByVal x As Single, ByVal y As Single)
    ToolsA.NotifyMouseMove down, sh, x, y
End Sub
