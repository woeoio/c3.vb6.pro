Option Explicit

' VBA.Math 的成员负控（PhotoDemon Math 定案，handoff §10.4）。
' Fix/Int 是**全局内建函数**，不是 VBA.Math 的成员：真 VB6 在这里报
' 「未找到方法或数据成员」，C3 必须在语义层报 VB3045 并停止发码
' （修前它们落进发码白名单，折成 vb6_vb6_Fix / vb6_vb6_Int，再到链接期 LNK2019）。

Public Sub Main()
    Debug.Print Math.Fix(1.7)
    Debug.Print Math.Int(1.7)
    Debug.Print "VMBAD-MEMBER PASS"
End Sub
