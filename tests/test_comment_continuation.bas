Option Explicit

' 注释续行：这条注释以「空白 + _」结尾，下一物理行仍属于同一条注释。
' 形如 PhotoDemon 里的 'Public Declare Function ... ( _
ByVal hWnd As Long, ByVal uMsg As Long) As Long
' 若注释续行没实现，上面那行会被当成代码 → 语法错。

Public Sub Main()
    Debug.Print "CC=1"
End Sub
