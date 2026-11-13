Attribute VB_Name = "ModNeg"
Option Explicit

' 探针夹具 (§B93, PhotoDemon C2064 第 2 族): 语句级无括号调用的**首参是负数**。
'   1) `o.Move -1 * w, h`  —— 右操作数是二元乘 (旧判据只认 LiteralExpr ⇒ 落空)
'   2) `o.Move -x, (y + 1)` —— 右操作数是标识符 (同样落空)
'   3) `o.Move -30, -40`    —— 存量形状 (Fix <vbeclipse> 早有, 必须逐字不变)
' 落空时整句被当成"调用一个减法结果": `(vb6_Mover_Move(o) - (1*w))(h)` ⇒ cl C2064。
' 负控: 赋值右侧的减法 `UseVal = o.Val - 1` 绝不能折叠成调用。

Public Sub UseNeg(ByRef o As Mover, ByVal w As Long, ByVal h As Long, _
                  ByVal x As Long, ByVal y As Long)
    o.Move -1 * w, h
    o.Move -x, (y + 1)
    o.Move -30, -40
End Sub

Public Function UseVal(ByRef o As Mover) As Long
    UseVal = o.Val - 1
End Function
