Attribute VB_Name = "Main"
Option Explicit

' 后测循环条件里出现「需要物化的 Variant 比较」(数组元素被推断成 Variant,
' 只能先存进 _vcmp_N 再取址) 的最小形状。与 PhotoDemon pdPSDLayer.cls
' `Loop While lineLengths(y) > bytesReadLine` 同形。
Public Function SumUntil(ByVal n As Long) As Long
    Dim xs() As Long
    Dim i As Long
    Dim acc As Long
    ReDim xs(0 To 10)
    Do
        acc = acc + xs(i)
        i = i + 1
    Loop While xs(i) > n
    SumUntil = acc
End Function
