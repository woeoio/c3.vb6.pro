Attribute VB_Name = "ModNd"
Option Explicit

' 探针夹具 (§38, PhotoDemon C2064 主族): UDT 的**动态数组字段**被下标访问时
' 以前被当成函数调用（`ch.cBytes(1, 2)` → C 里 `ch->cBytes(1, 2)` ⇒ C2064）。
' 已批准机理只覆盖秩≥2 的成员动态数组（秩 1 的 Pattern B 早就有）。
'   1) 秩 2 动态成员：读 / 写 / ByRef UDT 形参
'   2) 秩 2 动态成员：嵌套 UDT 字段（`ly.lMask.cBytes(...)`）
'   3) 秩 1 动态成员与定长 2 维成员是负控 —— 发码必须逐字不变
'      (`VB6_SA_AT(uint8_t, ch->cLine, 5)` / `ch->cFixed[2][1]`)

Public Type Chan
    cWidth As Long
    cHeight As Long
    cBytes() As Byte
    cLine() As Byte
    cFixed(0 To 3, 0 To 3) As Long
End Type

Public Type Layer
    lMask As Chan
End Type

Public Sub TouchChan(ByRef ch As Chan)
    Dim i As Long
    ch.cBytes(1, 2) = 7
    i = ch.cBytes(3, 4)
    ch.cLine(5) = 9
    i = ch.cLine(6)
    i = ch.cFixed(1, 2)
    ch.cFixed(3, 0) = i
End Sub

Public Sub TouchLayer(ByRef ly As Layer)
    Dim k As Long
    ly.lMask.cBytes(0, 0) = 1
    k = ly.lMask.cBytes(1, 1)
End Sub
